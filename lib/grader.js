// Grader: takes a parsed workbook and a list of SQL screenshots, asks Claude
// to grade each question against the rubric, returns per-question results.

import { QUESTIONS, STRICTNESS, SHEETS } from "./rubric.js";
import { callClaude, textOf, parseJsonResponse } from "./api.js";

const SYSTEM_PROMPT = `You are a meticulous Microsoft Excel and SQL grader. \
You evaluate one student answer at a time against the provided rubric and \
return a strict JSON object — nothing else, no prose. The JSON must follow \
this exact schema:

{
  "points_awarded": <number, between 0 and the question's max_points, may be fractional>,
  "verdict": "correct" | "partial" | "incorrect" | "missing",
  "evidence": "<short literal quote from the student's submission, eg the formula or value found, max 200 chars>",
  "feedback": "<one or two sentences explaining the grade, plain text, no markdown>"
}

Rules:
- Treat the rubric and the strictness level as authoritative.
- If you cannot find any evidence the student attempted the question, set verdict to "missing" and points to 0.
- Do NOT inflate points; under STRICT, missing the requested technique scores 0 even if the value is right.
- Never wrap the JSON in markdown fences. Output the JSON object only.`;

export async function gradeWorkbook({
  workbook,
  apiKey,
  model,
  strictness,
  studentName,
  signal,
  onProgress,
}) {
  const results = [];
  const excelQuestions = QUESTIONS.filter((q) => !q.sql);
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
      const evidence = q.evidence ? q.evidence(workbook) : null;
      const result = await gradeOne({
        apiKey,
        model,
        strictness,
        question: q,
        evidence,
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
  const evidenceText = evidence
    ? "```json\n" + JSON.stringify(evidence, null, 2).slice(0, 14000) + "\n```"
    : "(no evidence extracted — sheet may be missing)";
  return `QUESTION ${question.id} (max_points=${question.points}, section ${question.section})
${question.describe()}

EXPECTED:
${question.expected}

STRICTNESS LEVEL: ${strictness.toUpperCase()}
${STRICTNESS[strictness].description}

STUDENT EVIDENCE (extracted from the workbook — only the cells, formulas, validations, conditional formatting, charts, pivots and Power Query metadata relevant to this question):
${evidenceText}

Grade this single question. Return the JSON object only.`;
}

function clamp(value, min, max) {
  const n = Number(value);
  if (!Number.isFinite(n)) return 0;
  return Math.max(min, Math.min(max, n));
}
