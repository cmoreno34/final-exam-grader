// Grader: takes a parsed workbook and a list of SQL screenshots, asks Claude
// to grade each question against the rubric, returns per-question results.

import { QUESTIONS, STRICTNESS, SHEETS } from "./rubric.js";
import { callClaude, textOf, parseJsonResponse } from "./api.js";
import { fullWorkbookDump } from "./extractor.js";

const SYSTEM_PROMPT = `You are a meticulous Microsoft Excel and SQL grader. \
You evaluate student answers against the provided rubric and return strict \
JSON only — no prose, no markdown fences.

Each question's verdict must follow this schema:
{
  "points_awarded": <number, between 0 and the question's max_points, may be fractional>,
  "verdict": "correct" | "partial" | "incorrect" | "missing",
  "evidence": "<short literal quote from the student's submission — formula, cell ref, or value found, max 200 chars>",
  "feedback": "<one or two sentences explaining the grade, plain text, no markdown>"
}

CRITICAL: Search the ENTIRE workbook for each answer. Students often place \
answers in unexpected cells, columns, or sheets — different from the location \
named in the rubric. Look at every sheet, every non-empty cell, every chart, \
pivot, slicer, defined name, conditional formatting rule, data validation \
rule, and Power Query connection. If the student did the requested work but \
in a different location than the rubric specifies, AWARD THE POINTS and \
mention the actual location in the feedback.

Rules:
- Treat the rubric and the strictness level as authoritative.
- If after searching the entire workbook you cannot find any evidence the student attempted the question, set verdict to "missing" and points to 0.
- Do NOT inflate points; under STRICT, missing the requested technique scores 0 even if the value is right.
- Never wrap the JSON in markdown fences. Output the JSON only.`;

// Bulk Excel grader — sends the entire workbook + the full rubric in ONE call
// per student. Claude searches the whole workbook for each answer, so an
// answer placed in an unexpected sheet/column still gets credit.
export async function gradeWorkbook({
  workbook,
  apiKey,
  model,
  strictness,
  studentName,
  signal,
  onProgress,
}) {
  const excelQuestions = QUESTIONS.filter((q) => !q.sql);
  const dump = fullWorkbookDump(workbook);
  let dumpText = JSON.stringify(dump);
  // Hard cap to avoid runaway prompts. ~280KB is fine for Opus 4.7 (1M ctx).
  if (dumpText.length > 280_000) {
    dumpText = dumpText.slice(0, 280_000) + '..."[truncated]"}';
  }

  const rubricText = excelQuestions
    .map(
      (q) =>
        `${q.id} (max_points=${q.points}, section ${q.section}): ${q.label}\n` +
        `  Description: ${q.describe()}\n` +
        `  Expected: ${q.expected}`
    )
    .join("\n\n");

  const prompt = `Grade ONE student's Excel workbook against the rubric below.

STRICTNESS: ${strictness.toUpperCase()}
${STRICTNESS[strictness].description}

You must search the ENTIRE workbook for each answer — not only the sheet or
range mentioned in the question. If the student placed the answer in a
different sheet, column, or row than the rubric expected, you should still
locate it and award points (mention where you found it in the feedback).

RUBRIC:

${rubricText}

STUDENT WORKBOOK (every non-empty cell from every sheet plus structural
metadata for charts, pivots, slicers, data validations, conditional
formatting, defined names, tables and Power Query connections — JSON):

\`\`\`json
${dumpText}
\`\`\`

Return ONE JSON object whose top-level keys are the question IDs (Q1, Q2, …, Q15).
Each value is an object with the schema {points_awarded, verdict, evidence, feedback}.
Output the JSON object only. No prose, no markdown fences.`;

  onProgress?.({
    kind: "bulk",
    student: studentName,
    questionCount: excelQuestions.length,
  });

  let resp;
  try {
    resp = await callClaude({
      apiKey,
      model,
      system: SYSTEM_PROMPT,
      messages: [{ role: "user", content: [{ type: "text", text: prompt }] }],
      maxTokens: 8192,
      signal,
    });
  } catch (err) {
    if (err.name === "AbortError") throw err;
    return excelQuestions.map((q) => ({
      questionId: q.id,
      points_awarded: 0,
      verdict: "error",
      evidence: "",
      feedback: "Bulk grading call failed: " + (err.message ?? String(err)).slice(0, 240),
    }));
  }

  let parsed;
  try {
    parsed = parseJsonResponse(textOf(resp));
  } catch (err) {
    // Fall back to per-question grading if bulk JSON parsing failed
    return await gradePerQuestion({
      excelQuestions,
      workbook,
      apiKey,
      model,
      strictness,
      studentName,
      signal,
      onProgress,
    });
  }

  return excelQuestions.map((q) => {
    const r = parsed[q.id] ?? {};
    return {
      questionId: q.id,
      points_awarded: clamp(r.points_awarded, 0, q.points),
      verdict: r.verdict ?? "missing",
      evidence: (r.evidence ?? "").toString().slice(0, 240),
      feedback: r.feedback ?? "No verdict returned for this question.",
    };
  });
}

// Fallback: one Claude call per question, but each call still receives the
// full workbook so the search-anywhere behavior is preserved.
async function gradePerQuestion({
  excelQuestions,
  workbook,
  apiKey,
  model,
  strictness,
  studentName,
  signal,
  onProgress,
}) {
  const dump = fullWorkbookDump(workbook);
  const dumpText = JSON.stringify(dump).slice(0, 220_000);
  const results = [];
  for (let i = 0; i < excelQuestions.length; i++) {
    const q = excelQuestions[i];
    onProgress?.({
      kind: "question",
      student: studentName,
      questionId: q.id,
      index: i,
      total: excelQuestions.length,
    });
    try {
      const result = await gradeOne({
        apiKey,
        model,
        strictness,
        question: q,
        evidence: { fullWorkbook: dumpText },
        signal,
      });
      results.push({ ...result, questionId: q.id });
    } catch (err) {
      if (err.name === "AbortError") throw err;
      results.push({
        questionId: q.id,
        points_awarded: 0,
        verdict: "error",
        evidence: "",
        feedback: "Grader failed: " + (err.message ?? String(err)).slice(0, 240),
      });
    }
  }
  return results;
}

export async function gradeSqlScreenshots({
  screenshots,
  apiKey,
  model,
  strictness,
  studentName,
  signal,
  onProgress,
  referenceTables,
}) {
  const results = [];
  const sqlQuestions = QUESTIONS.filter((q) => q.sql);
  if (screenshots.length === 0) {
    return sqlQuestions.map((q) => ({
      questionId: q.id,
      points_awarded: 0,
      verdict: "missing",
      evidence: "",
      feedback: "No SQL screenshot uploaded for this student.",
    }));
  }

  // Send all screenshots in a single call so Claude can match them to questions.
  onProgress?.({
    kind: "sql",
    student: studentName,
    questionCount: sqlQuestions.length,
  });

  const systemPrompt = `${SYSTEM_PROMPT}\n\nFor SQL grading you will receive multiple screenshot images that may show queries and their results from sqliteonline.com. The student's screenshots may not be in question order. Match each screenshot to the appropriate question in the rubric. For each question, return a JSON object as described above.`;

  const rubricText = sqlQuestions
    .map(
      (q) =>
        `${q.id} (max_points=${q.points}): ${q.describe()}\nExpected: ${
          q.expected
        }`
    )
    .join("\n\n---\n\n");

  const messages = [
    {
      role: "user",
      content: [
        ...screenshots.map((s) => ({
          type: "image",
          source: { type: "base64", media_type: s.mediaType, data: s.base64 },
        })),
        {
          type: "text",
          text: `STRICTNESS LEVEL: ${strictness.toUpperCase()}\n${
            STRICTNESS[strictness].description
          }\n\nREFERENCE TABLES (the source-of-truth data used by the student):\n${referenceTables}\n\nRUBRIC for SQL questions:\n\n${rubricText}\n\nReturn a single JSON object of the form: { "Q16": {points_awarded, verdict, evidence, feedback}, "Q17": {...}, ..., "Q20": {...} }. Output the JSON only, no prose.`,
        },
      ],
    },
  ];

  const resp = await callClaude({
    apiKey,
    model,
    system: systemPrompt,
    messages,
    maxTokens: 4096,
    signal,
  });
  let parsed;
  try {
    parsed = parseJsonResponse(textOf(resp));
  } catch (err) {
    return sqlQuestions.map((q) => ({
      questionId: q.id,
      points_awarded: 0,
      verdict: "error",
      evidence: "",
      feedback: "Could not parse SQL grading response: " + err.message,
    }));
  }
  for (const q of sqlQuestions) {
    const r = parsed[q.id] ?? {};
    results.push({
      questionId: q.id,
      points_awarded: clamp(r.points_awarded, 0, q.points),
      verdict: r.verdict ?? "missing",
      evidence: r.evidence ?? "",
      feedback: r.feedback ?? "No feedback returned.",
    });
  }
  return results;
}

async function gradeOne({ apiKey, model, strictness, question, evidence, signal }) {
  const prompt = buildPrompt({ question, evidence, strictness });
  const resp = await callClaude({
    apiKey,
    model,
    system: SYSTEM_PROMPT,
    messages: [{ role: "user", content: [{ type: "text", text: prompt }] }],
    maxTokens: 1024,
    signal,
  });
  const parsed = parseJsonResponse(textOf(resp));
  return {
    points_awarded: clamp(parsed.points_awarded, 0, question.points),
    verdict: parsed.verdict ?? "missing",
    evidence: (parsed.evidence ?? "").toString().slice(0, 240),
    feedback: parsed.feedback ?? "",
  };
}

function buildPrompt({ question, evidence, strictness }) {
  let evidenceText;
  if (evidence?.fullWorkbook) {
    evidenceText =
      "```json\n" + evidence.fullWorkbook.slice(0, 220_000) + "\n```";
  } else if (evidence) {
    evidenceText =
      "```json\n" + JSON.stringify(evidence, null, 2).slice(0, 14000) + "\n```";
  } else {
    evidenceText = "(no evidence extracted — sheet may be missing)";
  }
  return `QUESTION ${question.id} (max_points=${question.points}, section ${question.section})
${question.describe()}

EXPECTED:
${question.expected}

STRICTNESS LEVEL: ${strictness.toUpperCase()}
${STRICTNESS[strictness].description}

STUDENT WORKBOOK (full dump — search every sheet, every cell, every chart/pivot/slicer/dataValidation/CF/PowerQuery for the answer; if the student placed it somewhere unexpected still credit them):
${evidenceText}

Grade this single question. Return the JSON object only.`;
}

function clamp(value, min, max) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(min, Math.min(max, n));
}
