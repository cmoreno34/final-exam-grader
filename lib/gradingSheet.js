// Merges the AI grades into the EXISTING "GRADING" sheet of the student's xlsx
// instead of creating a new "GRADING (auto)" sheet. The template already has a
// fully-styled GRADING sheet (rows B5:K27) with question rows in B6:E26, a per-
// question "grade (/10)" column in D, and a SUMPRODUCT total in D27. We:
//
//   1. Overwrite D6:D26 with the AI-derived grade out of 10 (preserving each
//      cell's style attribute) so the existing SUMPRODUCT in D27 picks them up.
//   2. Append a justification table below row 27 with verdict / evidence /
//      feedback per question, so the teacher can audit each grade in-place.
//   3. Clean up calcChain.xml + its relationship + content-type override and
//      force Excel to recompute formulas on open. This is what made the previous
//      version produce "needs recovery" warnings — the calcChain referenced
//      formulas that had been rewritten.
//
// We work on raw XML strings instead of DOMParser to avoid the namespace-
// prefix pollution that the previous DOM approach was introducing into the
// modified workbook. The output keeps every other sheet, drawing, chart,
// pivot, slicer and Power Query intact.

import { QUESTIONS } from "./rubric.js";

const GRADING_SHEET_NAME = "GRADING";

export async function injectGradingSheet(xlsxBytes, studentResult, runMeta) {
  const JSZip = window.JSZip;
  const zip = await JSZip.loadAsync(xlsxBytes);

  const wbXml = await readText(zip, "xl/workbook.xml");
  const relsXml = await readText(zip, "xl/_rels/workbook.xml.rels");
  if (!wbXml || !relsXml) {
    throw new Error("Workbook missing xl/workbook.xml or its relationships");
  }

  const sheetPath = resolveGradingSheetPath(wbXml, relsXml);
  if (!sheetPath) {
    throw new Error(
      `No "${GRADING_SHEET_NAME}" sheet found in workbook — student may have deleted the template's grading tab.`
    );
  }
  let sheetXml = await readText(zip, sheetPath);
  if (!sheetXml) {
    throw new Error(`GRADING sheet file not found at ${sheetPath}`);
  }

  // 1. Fill D6:D26 with AI grades, converted to a 0–10 scale matching the template.
  for (let i = 0; i < QUESTIONS.length; i++) {
    const q = QUESTIONS[i];
    const row = 6 + i;
    const result = studentResult.results.find((r) => r.questionId === q.id);
    const awarded = result?.points_awarded ?? 0;
    const grade10 = q.points > 0 ? round2((awarded / q.points) * 10) : 0;
    sheetXml = writeDCell(sheetXml, row, grade10);
  }

  // 2. Update D27 cached total (the SUMPRODUCT formula stays so the teacher
  //    can still tweak any D cell and watch the total recompute).
  const totalGrade10 = round2((studentResult.totalPoints || 0) / 10);
  sheetXml = updateTotalCell(sheetXml, totalGrade10);

  // 3. Append justification table just before </sheetData>.
  const justificationXml = buildJustificationRows(studentResult, runMeta);
  sheetXml = sheetXml.replace("</sheetData>", justificationXml + "</sheetData>");

  // 4. Update the sheet's <dimension/> so Excel knows about the new rows.
  const lastRow = 32 + QUESTIONS.length; // header row 32 + one per question
  sheetXml = sheetXml.replace(
    /<dimension ref="[^"]*"\/>/,
    `<dimension ref="B1:K${lastRow}"/>`
  );

  zip.file(sheetPath, sheetXml);

  // 5. Drop calcChain.xml + its rel + its Content_Types Override so Excel
  //    rebuilds the chain from scratch on open. This is what fixes the
  //    "we found a problem, want us to recover" prompt.
  await scrubCalcChain(zip);

  // 6. Force full recalculation on open, and make the GRADING tab the active
  //    one so the teacher lands on it.
  await rewriteWorkbookXml(zip);

  return await zip.generateAsync({ type: "uint8array" });
}

// ---------------------------------------------------------------------------
// Sheet path resolution
// ---------------------------------------------------------------------------

function resolveGradingSheetPath(wbXml, relsXml) {
  // Find <sheet name="GRADING" ... r:id="rIdN"/>
  const sheetMatch = wbXml.match(
    new RegExp(
      `<sheet\\b[^>]*\\bname="${escapeForRegex(GRADING_SHEET_NAME)}"[^>]*?\\sr:id="(rId\\d+)"`
    )
  );
  if (!sheetMatch) return null;
  const rid = sheetMatch[1];
  const relMatch = relsXml.match(
    new RegExp(`<Relationship\\b[^>]*\\bId="${rid}"[^>]*\\bTarget="([^"]+)"`)
  );
  if (!relMatch) return null;
  const target = relMatch[1];
  // Targets are relative to xl/. Some workbooks use "/xl/worksheets/sheet1.xml",
  // others use "worksheets/sheet1.xml" — normalise both.
  if (target.startsWith("/")) return target.replace(/^\//, "");
  return "xl/" + target.replace(/^xl\//, "");
}

// ---------------------------------------------------------------------------
// D-column cell rewriting
// ---------------------------------------------------------------------------

// Replace the D{row} cell's contents with a plain numeric value, preserving
// the style attribute (e.g. s="13") so the cell still renders as styled.
function writeDCell(sheetXml, row, grade) {
  const full = new RegExp(`<c r="D${row}"([^/>]*?)>[\\s\\S]*?<\\/c>`);
  if (full.test(sheetXml)) {
    return sheetXml.replace(full, `<c r="D${row}"$1><v>${grade}</v></c>`);
  }
  const self = new RegExp(`<c r="D${row}"([^/>]*?)\\/>`);
  if (self.test(sheetXml)) {
    return sheetXml.replace(self, `<c r="D${row}"$1><v>${grade}</v></c>`);
  }
  // Cell missing entirely — append it inside the matching <row>. Best-effort.
  const rowOpen = new RegExp(`(<row r="${row}"[^>]*>)([^<]*)`, "");
  if (rowOpen.test(sheetXml)) {
    return sheetXml.replace(
      rowOpen,
      (m, open, after) => `${open}<c r="D${row}"><v>${grade}</v></c>${after}`
    );
  }
  return sheetXml;
}

// D27 already has `<f>SUMPRODUCT(...)</f><v>cached</v>`. Update the cached
// value so the workbook displays the correct total even if Excel skipped
// the recalc; the formula itself stays intact.
function updateTotalCell(sheetXml, total) {
  const re = /<c r="D27"([^/>]*?)>(?:<f([^>]*)>([\s\S]*?)<\/f>)?(?:<v>[^<]*<\/v>)?<\/c>/;
  if (!re.test(sheetXml)) return sheetXml;
  return sheetXml.replace(re, (_, attrs, fAttrs, fText) => {
    const f = fText || "SUMPRODUCT(C6:C26,D6:D26)";
    const fAttrsStr = fAttrs || "";
    return `<c r="D27"${attrs}><f${fAttrsStr}>${f}</f><v>${total}</v></c>`;
  });
}

// ---------------------------------------------------------------------------
// Justification table (rows 28+)
// ---------------------------------------------------------------------------

function buildJustificationRows(sr, runMeta) {
  const out = [];
  // Row 28: blank separator
  out.push(`<row r="28" spans="2:7"/>`);
  // Row 29: title
  out.push(
    `<row r="29" spans="2:7">` +
      inlineCell(`B29`, `Auto-grading justification (Claude)`) +
      `</row>`
  );
  // Row 30: metadata
  const meta = `Model: ${runMeta.model} — Strictness: ${runMeta.strictness} — Graded at: ${sr.gradedAt}`;
  out.push(
    `<row r="30" spans="2:7">` + inlineCell(`B30`, meta) + `</row>`
  );
  // Row 31: blank
  out.push(`<row r="31" spans="2:7"/>`);
  // Row 32: header
  out.push(
    `<row r="32" spans="2:7">` +
      inlineCell(`B32`, `Question`) +
      inlineCell(`C32`, `Max points`) +
      inlineCell(`D32`, `Awarded`) +
      inlineCell(`E32`, `Verdict`) +
      inlineCell(`F32`, `Evidence`) +
      inlineCell(`G32`, `Feedback`) +
      `</row>`
  );
  // Rows 33+: one per question
  for (let i = 0; i < QUESTIONS.length; i++) {
    const q = QUESTIONS[i];
    const r = sr.results.find((x) => x.questionId === q.id) ?? {};
    const row = 33 + i;
    const label = `${q.id}: ${q.label}`;
    out.push(
      `<row r="${row}" spans="2:7" ht="30" customHeight="1">` +
        inlineCell(`B${row}`, label) +
        numCell(`C${row}`, q.points) +
        numCell(`D${row}`, round2(r.points_awarded ?? 0)) +
        inlineCell(`E${row}`, r.verdict || "") +
        inlineCell(`F${row}`, (r.evidence || "").toString()) +
        inlineCell(`G${row}`, (r.feedback || "").toString()) +
        `</row>`
    );
  }
  return out.join("");
}

function inlineCell(ref, text) {
  return (
    `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">` +
    escapeXml(text) +
    `</t></is></c>`
  );
}

function numCell(ref, n) {
  const v = Number.isFinite(Number(n)) ? Number(n) : 0;
  return `<c r="${ref}"><v>${v}</v></c>`;
}

// ---------------------------------------------------------------------------
// Calc-chain + workbook-level cleanup
// ---------------------------------------------------------------------------

async function scrubCalcChain(zip) {
  if (zip.file("xl/calcChain.xml")) {
    zip.remove("xl/calcChain.xml");
  }
  const relsText = await readText(zip, "xl/_rels/workbook.xml.rels");
  if (relsText) {
    const newRels = relsText.replace(
      /<Relationship\b[^>]*\bType="[^"]*\/calcChain"[^>]*\/>/g,
      ""
    );
    if (newRels !== relsText) {
      zip.file("xl/_rels/workbook.xml.rels", newRels);
    }
  }
  const ctText = await readText(zip, "[Content_Types].xml");
  if (ctText) {
    const newCt = ctText.replace(
      /<Override\s+PartName="\/xl\/calcChain\.xml"[^>]*\/>/g,
      ""
    );
    if (newCt !== ctText) {
      zip.file("[Content_Types].xml", newCt);
    }
  }
}

async function rewriteWorkbookXml(zip) {
  let wbXml = await readText(zip, "xl/workbook.xml");
  if (!wbXml) return;

  // Ensure calcPr has fullCalcOnLoad="1" so Excel recomputes all formulas
  // (the cached values for D27 are correct, but D6:D26 had formulas that we
  // stripped, so this guarantees correctness even with stale caches).
  if (/<calcPr\b[^/>]*\/>/.test(wbXml)) {
    wbXml = wbXml.replace(/<calcPr\b([^/>]*)\/>/, (m, attrs) => {
      if (/\bfullCalcOnLoad="1"/.test(attrs)) return m;
      return `<calcPr${attrs} fullCalcOnLoad="1"/>`;
    });
  } else if (!/<calcPr\b/.test(wbXml)) {
    wbXml = wbXml.replace(
      /<\/workbook>/,
      `<calcPr fullCalcOnLoad="1"/></workbook>`
    );
  }

  // Make GRADING the active tab. activeTab/firstSheet are 0-indexed and
  // GRADING is the first sheet in the template.
  wbXml = wbXml.replace(
    /(<workbookView\b[^/>]*?)\bactiveTab="\d+"/,
    '$1activeTab="0"'
  );
  wbXml = wbXml.replace(
    /(<workbookView\b[^/>]*?)\bfirstSheet="\d+"/,
    '$1firstSheet="0"'
  );

  zip.file("xl/workbook.xml", wbXml);
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

async function readText(zip, path) {
  const f = zip.file(path);
  if (!f) return null;
  return await f.async("string");
}

function escapeXml(s) {
  return String(s ?? "")
    .replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g, "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeForRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}
