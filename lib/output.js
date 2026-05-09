// Output: build the per-student feedback markdown, the master grading xlsx,
// and bundle everything into a single zip the user can download.

import { QUESTIONS, EXAM_TOTAL } from "./rubric.js";

export async function buildOutputZip({ studentResults, runMeta }) {
  const JSZip = window.JSZip;
  const zip = new JSZip();

  // 1. Master grading sheet (xlsx) using SheetJS
  const masterAoa = buildMasterAoa(studentResults, runMeta);
  const wb = window.XLSX.utils.book_new();
  const ws = window.XLSX.utils.aoa_to_sheet(masterAoa);
  window.XLSX.utils.book_append_sheet(wb, ws, "Grades");
  // Per-question detail sheet
  const detailAoa = buildDetailAoa(studentResults);
  const detailWs = window.XLSX.utils.aoa_to_sheet(detailAoa);
  window.XLSX.utils.book_append_sheet(wb, detailWs, "Question detail");
  const xlsxBlob = window.XLSX.write(wb, { type: "array", bookType: "xlsx" });
  zip.file("00_grades.xlsx", xlsxBlob);

  // 2. Master CSV (handy for quick paste into LMS)
  zip.file("00_grades.csv", aoaToCsv(masterAoa));

  // 3. Per-student folder with original file + feedback.md
  for (const sr of studentResults) {
    const folder = zip.folder(sanitizePath(sr.studentName));
    if (sr.originalFile) {
      folder.file(sr.originalFile.name, sr.originalFile.bytes);
    }
    folder.file("feedback.md", buildFeedbackMd(sr));
    folder.file("results.json", JSON.stringify(sr, null, 2));
  }

  // 4. Run summary
  zip.file("README.txt", buildRunSummary(studentResults, runMeta));

  return await zip.generateAsync({ type: "blob" });
}

function buildMasterAoa(studentResults, runMeta) {
  const headers = ["Student", "Total /100"];
  for (const q of QUESTIONS) headers.push(`${q.id} (/${q.points})`);
  headers.push("Strictness", "Model", "Graded at");

  const rows = [headers];
  for (const sr of studentResults) {
    const row = [sr.studentName, round2(sr.totalPoints)];
    for (const q of QUESTIONS) {
      const r = sr.results.find((x) => x.questionId === q.id);
      row.push(r ? round2(r.points_awarded) : 0);
    }
    row.push(runMeta.strictness, runMeta.model, sr.gradedAt);
    rows.push(row);
  }
  // Class average row
  const avgRow = ["Class average", classAvg(studentResults, "totalPoints")];
  for (const q of QUESTIONS) {
    avgRow.push(classAvgQ(studentResults, q.id));
  }
  avgRow.push("", "", "");
  rows.push(avgRow);
  return rows;
}

function buildDetailAoa(studentResults) {
  const rows = [
    ["Student", "Question", "Max", "Awarded", "Verdict", "Evidence", "Feedback"],
  ];
  for (const sr of studentResults) {
    for (const q of QUESTIONS) {
      const r = sr.results.find((x) => x.questionId === q.id) ?? {};
      rows.push([
        sr.studentName,
        q.id,
        q.points,
        round2(r.points_awarded ?? 0),
        r.verdict ?? "missing",
        (r.evidence ?? "").toString(),
        (r.feedback ?? "").toString(),
      ]);
    }
  }
  return rows;
}

function buildFeedbackMd(sr) {
  const lines = [];
  lines.push(`# Feedback — ${sr.studentName}`);
  lines.push("");
  lines.push(`**Total:** ${round2(sr.totalPoints)} / ${EXAM_TOTAL}`);
  lines.push(`**Graded at:** ${sr.gradedAt}`);
  lines.push("");
  let currentSection = -1;
  for (const q of QUESTIONS) {
    if (q.section !== currentSection) {
      currentSection = q.section;
      lines.push(`## Section ${currentSection}`);
      lines.push("");
    }
    const r = sr.results.find((x) => x.questionId === q.id) ?? {};
    const awarded = round2(r.points_awarded ?? 0);
    const verdict = r.verdict ?? "missing";
    const icon =
      verdict === "correct" ? "OK" :
      verdict === "partial" ? "PARTIAL" :
      verdict === "incorrect" ? "WRONG" :
      verdict === "error" ? "ERROR" : "MISSING";
    lines.push(`### ${q.id} — ${q.label}  (${awarded} / ${q.points})  [${icon}]`);
    lines.push(`*${q.describe()}*`);
    lines.push("");
    if (r.feedback) lines.push(`**Feedback:** ${r.feedback}`);
    if (r.evidence) lines.push(`**Evidence:** \`${shortEvidence(r.evidence)}\``);
    lines.push("");
  }
  return lines.join("\n");
}

function buildRunSummary(studentResults, runMeta) {
  return [
    "FINAL EXAM grading run",
    "----------------------",
    `Students graded:  ${studentResults.length}`,
    `Strictness:       ${runMeta.strictness}`,
    `Model:            ${runMeta.model}`,
    `Run started:      ${runMeta.startedAt}`,
    `Run finished:     ${runMeta.finishedAt}`,
    "",
    "Class average:    " + classAvg(studentResults, "totalPoints") + " / 100",
    "Highest score:    " + Math.max(...studentResults.map((s) => s.totalPoints || 0)),
    "Lowest score:     " + Math.min(...studentResults.map((s) => s.totalPoints || 0)),
    "",
    "Open 00_grades.xlsx for the master sheet.",
    "Open <student>/feedback.md for per-student detail.",
  ].join("\n");
}

function aoaToCsv(aoa) {
  return aoa
    .map((row) =>
      row
        .map((cell) => {
          const s = (cell ?? "").toString();
          if (/[,"\n]/.test(s)) return '"' + s.replace(/"/g, '""') + '"';
          return s;
        })
        .join(",")
    )
    .join("\n");
}

function sanitizePath(name) {
  return name.replace(/[\\/:*?"<>|]/g, "_").trim() || "student";
}

function shortEvidence(text) {
  return text.toString().replace(/\s+/g, " ").slice(0, 160);
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}

function classAvg(students, key) {
  if (!students.length) return 0;
  const sum = students.reduce((a, s) => a + (s[key] || 0), 0);
  return round2(sum / students.length);
}

function classAvgQ(students, qid) {
  if (!students.length) return 0;
  const vals = students.map((s) => {
    const r = s.results.find((x) => x.questionId === qid);
    return r?.points_awarded ?? 0;
  });
  return round2(vals.reduce((a, b) => a + b, 0) / vals.length);
}
