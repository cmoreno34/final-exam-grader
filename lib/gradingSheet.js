// Injects a "GRADING (auto)" sheet directly into the student's xlsx so the
// teacher can open the file, see the verdict, and tweak any "Awarded" cell —
// the total is a SUM formula that updates automatically.
//
// Strategy: we don't round-trip the workbook through SheetJS (which would
// destroy charts, pivots, slicers, Power Query, etc.). Instead we open the
// xlsx as a zip, drop in a new worksheet XML, and edit the three index files:
//   - xl/workbook.xml             (add a <sheet>)
//   - xl/_rels/workbook.xml.rels  (add a <Relationship> to the new sheet)
//   - [Content_Types].xml         (add an <Override>)
// Existing GRADING (auto) sheets from a prior run are removed first so the
// injection is idempotent.

import { QUESTIONS, EXAM_TOTAL } from "./rubric.js";

const SSML = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_OFFICEDOC =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const SHEET_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
const SHEET_CT =
  "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

const TARGET_SHEET_NAME = "GRADING (auto)";

export async function injectGradingSheet(xlsxBytes, studentResult, runMeta) {
  const JSZip = window.JSZip;
  const zip = await JSZip.loadAsync(xlsxBytes);

  const wbDoc = await readXmlFromZip(zip, "xl/workbook.xml");
  const relsDoc = await readXmlFromZip(zip, "xl/_rels/workbook.xml.rels");
  const ctDoc = await readXmlFromZip(zip, "[Content_Types].xml");
  if (!wbDoc || !relsDoc || !ctDoc) {
    throw new Error("Workbook is missing required parts (workbook/rels/content-types)");
  }

  removeExistingGradingSheet(zip, wbDoc, relsDoc, ctDoc);

  // Pick fresh ids
  const newRid = nextRid(relsDoc);
  const newSheetId = nextSheetId(wbDoc);
  const newSheetN = nextSheetFileNumber(zip);
  const newSheetPath = `xl/worksheets/sheet${newSheetN}.xml`;

  // Write the new worksheet
  zip.file(newSheetPath, buildGradingSheetXml(studentResult, runMeta));

  // Register the new sheet in workbook.xml as the FIRST sheet so it opens by default
  const sheetsParent = wbDoc.getElementsByTagNameNS(SSML, "sheets")[0];
  const newSheet = wbDoc.createElementNS(SSML, "sheet");
  newSheet.setAttribute("name", TARGET_SHEET_NAME);
  newSheet.setAttribute("sheetId", String(newSheetId));
  newSheet.setAttributeNS(REL_OFFICEDOC, "r:id", newRid);
  sheetsParent.insertBefore(newSheet, sheetsParent.firstChild);

  // Make this sheet the active tab
  const view = wbDoc.getElementsByTagNameNS(SSML, "workbookView")[0];
  if (view) {
    view.setAttribute("activeTab", "0");
    view.setAttribute("firstSheet", "0");
  }

  // Add the workbook relationship pointing at the new sheet
  const newRel = relsDoc.createElementNS(REL, "Relationship");
  newRel.setAttribute("Id", newRid);
  newRel.setAttribute("Type", SHEET_TYPE);
  newRel.setAttribute("Target", `worksheets/sheet${newSheetN}.xml`);
  relsDoc.documentElement.appendChild(newRel);

  // Add the content-types Override
  const newOverride = ctDoc.createElementNS(CT, "Override");
  newOverride.setAttribute("PartName", `/${newSheetPath}`);
  newOverride.setAttribute("ContentType", SHEET_CT);
  ctDoc.documentElement.appendChild(newOverride);

  // Force Excel to recompute formulas on open (the cached SUM value below is a hint).
  ensureFullCalcOnLoad(wbDoc);
  // Drop calcChain.xml — Excel rebuilds it. Avoids "needs recovery" warning.
  if (zip.file("xl/calcChain.xml")) zip.remove("xl/calcChain.xml");

  // Serialize back
  const ser = new XMLSerializer();
  zip.file("xl/workbook.xml", ser.serializeToString(wbDoc));
  zip.file("xl/_rels/workbook.xml.rels", ser.serializeToString(relsDoc));
  zip.file("[Content_Types].xml", ser.serializeToString(ctDoc));

  return await zip.generateAsync({ type: "uint8array" });
}

// ---------------------------------------------------------------------------

function removeExistingGradingSheet(zip, wbDoc, relsDoc, ctDoc) {
  const sheets = Array.from(wbDoc.getElementsByTagNameNS(SSML, "sheet"));
  for (const sh of sheets) {
    if (sh.getAttribute("name") !== TARGET_SHEET_NAME) continue;
    const rid = sh.getAttributeNS(REL_OFFICEDOC, "id");
    sh.parentNode.removeChild(sh);
    const rels = Array.from(relsDoc.getElementsByTagName("Relationship"));
    for (const rel of rels) {
      if (rel.getAttribute("Id") !== rid) continue;
      const target = rel.getAttribute("Target") ?? "";
      const oldPath = target.startsWith("/")
        ? target.slice(1)
        : "xl/" + target.replace(/^xl\//, "");
      if (zip.file(oldPath)) zip.remove(oldPath);
      // Drop the matching content-types override
      const overrides = Array.from(ctDoc.getElementsByTagName("Override"));
      for (const ov of overrides) {
        if (ov.getAttribute("PartName") === "/" + oldPath) {
          ov.parentNode.removeChild(ov);
        }
      }
      rel.parentNode.removeChild(rel);
      break;
    }
  }
}

function nextRid(relsDoc) {
  let max = 0;
  for (const r of relsDoc.getElementsByTagName("Relationship")) {
    const m = /^rId(\d+)$/i.exec(r.getAttribute("Id") ?? "");
    if (m) max = Math.max(max, parseInt(m[1], 10));
  }
  return `rId${max + 1}`;
}

function nextSheetId(wbDoc) {
  let max = 0;
  for (const s of wbDoc.getElementsByTagNameNS(SSML, "sheet")) {
    const v = parseInt(s.getAttribute("sheetId") ?? "0", 10);
    if (Number.isFinite(v)) max = Math.max(max, v);
  }
  return max + 1;
}

function nextSheetFileNumber(zip) {
  let n = 1;
  while (zip.file(`xl/worksheets/sheet${n}.xml`)) n++;
  return n;
}

function ensureFullCalcOnLoad(wbDoc) {
  let calcPr = wbDoc.getElementsByTagNameNS(SSML, "calcPr")[0];
  if (!calcPr) {
    calcPr = wbDoc.createElementNS(SSML, "calcPr");
    wbDoc.documentElement.appendChild(calcPr);
  }
  calcPr.setAttribute("fullCalcOnLoad", "1");
}

async function readXmlFromZip(zip, path) {
  const file = zip.file(path);
  if (!file) return null;
  const text = await file.async("string");
  return new DOMParser().parseFromString(text, "application/xml");
}

// ---------------------------------------------------------------------------
// Worksheet XML builder
// ---------------------------------------------------------------------------

function buildGradingSheetXml(sr, runMeta) {
  const allQs = QUESTIONS; // all 20 questions: Excel (Q1–Q15) + SQL (Q16–Q20)
  // Layout:
  //   row 1: title
  //   row 2: Student | <name>
  //   row 3: Strictness | <value>
  //   row 4: Model | <value>
  //   row 5: Graded at | <iso>
  //   row 6: (blank)
  //   row 7: Total /100  | =SUM(C10:C{end})  | / | 100
  //   row 8: (blank)
  //   row 9: Question | Max | Awarded | Verdict | Evidence | Feedback
  //   row 10..: per-question rows
  const rows = [];
  rows.push([txt("GRADING (auto) — edit the Awarded column to override; the Total updates via SUM")]);
  rows.push([txt("Student"), txt(sr.studentName)]);
  rows.push([txt("Strictness"), txt(runMeta.strictness)]);
  rows.push([txt("Model"), txt(runMeta.model)]);
  rows.push([txt("Graded at"), txt(sr.gradedAt)]);
  rows.push([]);

  const dataStartRow = 10;
  const dataEndRow = dataStartRow + allQs.length - 1;
  const totalCached = round2(
    allQs.reduce((acc, q) => {
      const r = sr.results.find((x) => x.questionId === q.id);
      return acc + (r?.points_awarded || 0);
    }, 0)
  );

  rows.push([
    txt("Total /100"),
    formula(`SUM(C${dataStartRow}:C${dataEndRow})`, totalCached),
    txt("/"),
    num(EXAM_TOTAL),
  ]);
  rows.push([]);
  rows.push([
    txt("Question"),
    txt("Max"),
    txt("Awarded"),
    txt("Verdict"),
    txt("Evidence"),
    txt("Feedback"),
  ]);

  for (const q of allQs) {
    const r = sr.results.find((x) => x.questionId === q.id) ?? {};
    rows.push([
      txt(q.id),
      num(q.points),
      num(round2(r.points_awarded ?? 0)),
      txt(r.verdict ?? ""),
      txt((r.evidence ?? "").toString()),
      txt((r.feedback ?? "").toString()),
    ]);
  }

  let body = "";
  rows.forEach((row, ri) => {
    if (!row.length) {
      body += `<row r="${ri + 1}"/>`;
      return;
    }
    body += `<row r="${ri + 1}">`;
    row.forEach((cell, ci) => {
      const ref = colLetter(ci + 1) + (ri + 1);
      body += renderCell(ref, cell);
    });
    body += "</row>";
  });

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="${SSML}" xmlns:r="${REL_OFFICEDOC}">
<dimension ref="A1:F${rows.length}"/>
<sheetViews><sheetView tabSelected="1" workbookViewId="0"/></sheetViews>
<sheetFormatPr defaultRowHeight="15"/>
<cols>
<col min="1" max="1" width="14" customWidth="1"/>
<col min="2" max="2" width="8" customWidth="1"/>
<col min="3" max="3" width="11" customWidth="1"/>
<col min="4" max="4" width="11" customWidth="1"/>
<col min="5" max="5" width="44" customWidth="1"/>
<col min="6" max="6" width="70" customWidth="1"/>
</cols>
<sheetData>${body}</sheetData>
</worksheet>`;
}

function txt(s) {
  return { kind: "s", value: String(s ?? "") };
}
function num(n) {
  return { kind: "n", value: Number(n) };
}
function formula(f, cachedValue) {
  return { kind: "f", formula: f, cached: cachedValue };
}

function renderCell(ref, cell) {
  if (!cell) return "";
  if (cell.kind === "s") {
    return `<c r="${ref}" t="inlineStr"><is><t xml:space="preserve">${escapeXml(cell.value)}</t></is></c>`;
  }
  if (cell.kind === "n") {
    return `<c r="${ref}"><v>${escapeNum(cell.value)}</v></c>`;
  }
  if (cell.kind === "f") {
    const v = cell.cached !== undefined && cell.cached !== null
      ? `<v>${escapeNum(cell.cached)}</v>`
      : "";
    return `<c r="${ref}"><f>${escapeXml(cell.formula)}</f>${v}</c>`;
  }
  return "";
}

function escapeXml(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
function escapeNum(n) {
  if (!Number.isFinite(n)) return "0";
  return String(n);
}

function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function round2(n) {
  return Math.round((Number(n) || 0) * 100) / 100;
}
