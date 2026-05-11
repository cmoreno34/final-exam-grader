// xlsx extractor: parses an .xlsx blob into a structured JSON workbook summary.
// Uses JSZip + DOMParser. Reads cell values, formulas, number formats, data
// validations, conditional formatting, defined names, charts, pivots, slicers
// and Power Query connections — everything needed by the rubric.

import { colLetter, colNum, parseRef } from "./rubric.js";

const NS = {
  ssml: "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
  rel: "http://schemas.openxmlformats.org/package/2006/relationships",
  drawingml: "http://schemas.openxmlformats.org/drawingml/2006/main",
  chart: "http://schemas.openxmlformats.org/drawingml/2006/chart",
  x14: "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
  x15: "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main",
};

export async function extractWorkbook(blob) {
  const JSZip = window.JSZip;
  if (!JSZip) throw new Error("JSZip not loaded");
  const zip = await JSZip.loadAsync(blob);

  const xml = (path) =>
    zip.file(path) ? zip.file(path).async("string") : Promise.resolve(null);

  const workbookXml = await xml("xl/workbook.xml");
  const relsXml = await xml("xl/_rels/workbook.xml.rels");
  if (!workbookXml) throw new Error("Not an .xlsx (no xl/workbook.xml)");

  const wbDoc = parseXml(workbookXml);
  const relsDoc = relsXml ? parseXml(relsXml) : null;

  // Sheet name → relationship id → file path
  const relsById = {};
  if (relsDoc) {
    for (const rel of relsDoc.getElementsByTagName("Relationship")) {
      relsById[rel.getAttribute("Id")] = {
        target: rel.getAttribute("Target"),
        type: rel.getAttribute("Type"),
      };
    }
  }
  const sheetRelsByName = {};
  for (const sh of wbDoc.getElementsByTagNameNS(NS.ssml, "sheet")) {
    const name = sh.getAttribute("name");
    const rid = sh.getAttributeNS(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "id"
    );
    sheetRelsByName[name] = relsById[rid]?.target;
  }
  const definedNames = [];
  for (const dn of wbDoc.getElementsByTagNameNS(NS.ssml, "definedName")) {
    definedNames.push({
      name: dn.getAttribute("name"),
      ref: dn.textContent,
    });
  }

  // sharedStrings
  const sst = await loadSharedStrings(zip);
  // styles -> numFmt lookups
  const styles = await loadStyles(zip);

  // Walk sheets
  const sheets = {};
  const allCharts = [];
  const allTables = [];
  const allPivots = [];
  const allSlicers = [];

  for (const [name, target] of Object.entries(sheetRelsByName)) {
    if (!target) continue;
    const sheetPath = "xl/" + target.replace(/^\/?xl\//, "");
    const sheetXml = await xml(sheetPath);
    if (!sheetXml) continue;
    const sheet = await parseSheet(zip, sheetPath, sheetXml, sst, styles);
    sheets[name] = sheet;
    allCharts.push(...sheet._charts);
    allTables.push(...sheet._tables);
    allPivots.push(...sheet._pivots);
    allSlicers.push(...sheet._slicers);
  }

  // Power Query
  const powerQuery = await detectPowerQuery(zip);
  const connections = await loadConnections(zip);

  // Embedded images — students often paste SQL screenshots directly into the
  // section-3 sheet instead of attaching them as separate files. We surface
  // them here so the SQL grader can use them.
  const embeddedImages = await extractEmbeddedImages(zip);

  return {
    definedNames,
    sheets,
    charts: allCharts,
    tables: allTables,
    pivotTables: allPivots,
    slicers: allSlicers,
    powerQuery,
    connections,
    embeddedImages,
  };
}

// Pulls every image in xl/media/ out as base64. Returns [{base64, mediaType, path}].
async function extractEmbeddedImages(zip) {
  const paths = [];
  zip.forEach((relPath) => {
    if (/^xl\/media\/[^/]+\.(png|jpe?g)$/i.test(relPath)) paths.push(relPath);
  });
  const out = [];
  for (const path of paths) {
    const isPng = /\.png$/i.test(path);
    try {
      const base64 = await zip.file(path).async("base64");
      out.push({
        base64,
        mediaType: isPng ? "image/png" : "image/jpeg",
        path,
      });
    } catch (e) {
      // skip unreadable images
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Shared strings + styles
// ---------------------------------------------------------------------------

async function loadSharedStrings(zip) {
  const f = zip.file("xl/sharedStrings.xml");
  if (!f) return [];
  const doc = parseXml(await f.async("string"));
  const out = [];
  for (const si of doc.getElementsByTagNameNS(NS.ssml, "si")) {
    // si may contain <t> directly or rich runs <r><t>
    let s = "";
    for (const t of si.getElementsByTagNameNS(NS.ssml, "t")) {
      s += t.textContent;
    }
    out.push(s);
  }
  return out;
}

async function loadStyles(zip) {
  const f = zip.file("xl/styles.xml");
  if (!f) return { numFmtById: {}, cellXfs: [] };
  const doc = parseXml(await f.async("string"));
  const numFmtById = { ...BUILTIN_NUM_FMT };
  for (const nf of doc.getElementsByTagNameNS(NS.ssml, "numFmt")) {
    numFmtById[nf.getAttribute("numFmtId")] = nf.getAttribute("formatCode");
  }
  const cellXfs = [];
  const xfsRoot = doc.getElementsByTagNameNS(NS.ssml, "cellXfs")[0];
  if (xfsRoot) {
    for (const xf of xfsRoot.getElementsByTagNameNS(NS.ssml, "xf")) {
      cellXfs.push({
        numFmtId: xf.getAttribute("numFmtId") ?? "0",
        applyNumberFormat: xf.getAttribute("applyNumberFormat") === "1",
      });
    }
  }
  return { numFmtById, cellXfs };
}

const BUILTIN_NUM_FMT = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0;(#,##0)",
  38: "#,##0;[Red](#,##0)",
  39: "#,##0.00;(#,##0.00)",
  40: "#,##0.00;[Red](#,##0.00)",
  44: '_("$"* #,##0.00_);_("$"* (#,##0.00);_("$"* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "mm:ss.0",
  48: "##0.0E+0",
  49: "@",
};

// ---------------------------------------------------------------------------
// Sheet parsing
// ---------------------------------------------------------------------------

async function parseSheet(zip, sheetPath, xmlText, sst, styles) {
  const doc = parseXml(xmlText);
  const cells = {};
  let maxCol = 0;
  let maxRow = 0;

  for (const c of doc.getElementsByTagNameNS(NS.ssml, "c")) {
    const ref = c.getAttribute("r");
    if (!ref) continue;
    const t = c.getAttribute("t"); // s, str, n, b, e, inlineStr, d
    const styleId = c.getAttribute("s");
    const fEl = c.getElementsByTagNameNS(NS.ssml, "f")[0];
    const vEl = c.getElementsByTagNameNS(NS.ssml, "v")[0];
    const isEl = c.getElementsByTagNameNS(NS.ssml, "is")[0];
    let v = vEl ? vEl.textContent : null;
    if (t === "s") v = sst[parseInt(v, 10)] ?? "";
    else if (t === "b") v = v === "1";
    else if (t === "n" || t === null) v = v === null ? null : Number(v);
    else if (t === "inlineStr") {
      v = isEl ? Array.from(isEl.getElementsByTagNameNS(NS.ssml, "t")).map(x => x.textContent).join("") : "";
    }
    const numFmt = lookupNumFmt(styleId, styles);
    const cell = {};
    if (fEl) cell.f = fEl.textContent;
    if (v !== null && v !== undefined) cell.v = v;
    if (numFmt && numFmt !== "General") cell.numFmt = numFmt;
    if (styleId) cell.styleId = styleId;
    cells[ref] = cell;
    const [col, row] = parseRef(ref);
    if (col > maxCol) maxCol = col;
    if (row > maxRow) maxRow = row;
  }

  // Header row (row 1)
  const headers = [];
  for (let c = 1; c <= maxCol; c++) {
    const ref = colLetter(c) + 1;
    const cell = cells[ref];
    headers.push(cell?.v ?? null);
  }

  // Data validations
  const dataValidations = [];
  for (const dv of doc.getElementsByTagNameNS(NS.ssml, "dataValidation")) {
    dataValidations.push({
      type: dv.getAttribute("type"),
      sqref: dv.getAttribute("sqref"),
      formula1:
        dv.getElementsByTagNameNS(NS.ssml, "formula1")[0]?.textContent ?? null,
      formula2:
        dv.getElementsByTagNameNS(NS.ssml, "formula2")[0]?.textContent ?? null,
    });
  }
  // x14 ext data validations
  for (const dv of doc.getElementsByTagNameNS(NS.x14, "dataValidation")) {
    dataValidations.push({
      type: dv.getAttribute("type"),
      sqref:
        dv.getElementsByTagNameNS(
          "http://schemas.microsoft.com/office/spreadsheetml/2006/main",
          "sqref"
        )[0]?.textContent ??
        dv.getAttribute("sqref"),
      formula1:
        dv.getElementsByTagNameNS(NS.x14, "formula1")[0]?.textContent ?? null,
      formula2:
        dv.getElementsByTagNameNS(NS.x14, "formula2")[0]?.textContent ?? null,
      ext: true,
    });
  }

  // Conditional formatting
  const conditionalFormattings = [];
  for (const cf of doc.getElementsByTagNameNS(NS.ssml, "conditionalFormatting")) {
    const sqref = cf.getAttribute("sqref");
    const rules = [];
    for (const rule of cf.getElementsByTagNameNS(NS.ssml, "cfRule")) {
      rules.push({
        type: rule.getAttribute("type"),
        priority: rule.getAttribute("priority"),
        operator: rule.getAttribute("operator"),
        formula: Array.from(rule.getElementsByTagNameNS(NS.ssml, "formula"))
          .map((f) => f.textContent)
          .join(" || "),
        dxfId: rule.getAttribute("dxfId"),
      });
    }
    conditionalFormattings.push({ sqref, rules });
  }

  // Sheet relationships → drawing → charts, table refs, pivot refs
  const relsPath = sheetPath.replace(
    /xl\/worksheets\/(.*)$/,
    "xl/worksheets/_rels/$1.rels"
  );
  const _charts = [];
  const _tables = [];
  const _pivots = [];
  const _slicers = [];
  const relsXml = zip.file(relsPath) ? await zip.file(relsPath).async("string") : null;
  if (relsXml) {
    const r = parseXml(relsXml);
    for (const rel of r.getElementsByTagName("Relationship")) {
      const type = rel.getAttribute("Type") ?? "";
      const target = rel.getAttribute("Target") ?? "";
      const path = resolveRel(sheetPath, target);
      if (type.endsWith("/drawing")) {
        const charts = await readDrawingCharts(zip, path);
        _charts.push(...charts);
      } else if (type.endsWith("/table")) {
        const tbl = await readTable(zip, path);
        if (tbl) _tables.push(tbl);
      } else if (type.endsWith("/pivotTable")) {
        const piv = await readPivot(zip, path);
        if (piv) _pivots.push(piv);
      } else if (type.endsWith("/slicer") || type.includes("slicer")) {
        const sl = await readSlicer(zip, path);
        if (sl) _slicers.push(sl);
      }
    }
  }

  return {
    cells,
    maxCol,
    maxRow,
    headers,
    dataValidations,
    conditionalFormattings,
    _charts,
    _tables,
    _pivots,
    _slicers,
  };
}

function lookupNumFmt(styleId, styles) {
  if (styleId === null || styleId === undefined) return null;
  const xf = styles.cellXfs[parseInt(styleId, 10)];
  if (!xf) return null;
  return styles.numFmtById[xf.numFmtId] ?? null;
}

// ---------------------------------------------------------------------------
// Charts / Tables / Pivots / Slicers
// ---------------------------------------------------------------------------

async function readDrawingCharts(zip, drawingPath) {
  const drawingDoc = parseXml((await zip.file(drawingPath)?.async("string")) ?? "");
  const drawingRelsPath = drawingPath.replace(
    /xl\/drawings\/(.*)$/,
    "xl/drawings/_rels/$1.rels"
  );
  const relsXml = zip.file(drawingRelsPath)
    ? await zip.file(drawingRelsPath).async("string")
    : null;
  if (!relsXml) return [];
  const relDoc = parseXml(relsXml);
  const chartTargets = [];
  for (const rel of relDoc.getElementsByTagName("Relationship")) {
    const t = rel.getAttribute("Type") ?? "";
    if (t.endsWith("/chart"))
      chartTargets.push(resolveRel(drawingPath, rel.getAttribute("Target")));
  }
  const charts = [];
  for (const cp of chartTargets) {
    const chartXml = (await zip.file(cp)?.async("string")) ?? "";
    if (!chartXml) continue;
    charts.push(summarizeChart(cp, chartXml));
  }
  return charts;
}

function summarizeChart(path, xmlText) {
  const doc = parseXml(xmlText);
  const types = [];
  const TYPE_TAGS = [
    "barChart",
    "lineChart",
    "pieChart",
    "areaChart",
    "scatterChart",
    "bubbleChart",
    "radarChart",
    "stockChart",
    "doughnutChart",
    "ofPieChart",
    "surfaceChart",
    "surface3DChart",
    "bar3DChart",
    "line3DChart",
    "pie3DChart",
    "area3DChart",
  ];
  for (const tag of TYPE_TAGS) {
    if (doc.getElementsByTagNameNS(NS.chart, tag).length > 0) types.push(tag);
  }
  // Trendlines
  const trendlines = [];
  for (const t of doc.getElementsByTagNameNS(NS.chart, "trendline")) {
    trendlines.push({
      trendlineType:
        t.getElementsByTagNameNS(NS.chart, "trendlineType")[0]?.getAttribute("val") ??
        null,
      forward:
        t.getElementsByTagNameNS(NS.chart, "forward")[0]?.getAttribute("val") ?? null,
      backward:
        t.getElementsByTagNameNS(NS.chart, "backward")[0]?.getAttribute("val") ??
        null,
      dispRSqr:
        t.getElementsByTagNameNS(NS.chart, "dispRSqr")[0]?.getAttribute("val") ===
        "1",
      dispEq:
        t.getElementsByTagNameNS(NS.chart, "dispEq")[0]?.getAttribute("val") === "1",
    });
  }
  // Axes
  const axes = [];
  for (const ax of doc.getElementsByTagNameNS(NS.chart, "valAx")) {
    axes.push({
      kind: "val",
      axId: ax.getElementsByTagNameNS(NS.chart, "axId")[0]?.getAttribute("val"),
      crosses:
        ax.getElementsByTagNameNS(NS.chart, "crosses")[0]?.getAttribute("val") ??
        null,
      numFmt:
        ax.getElementsByTagNameNS(NS.chart, "numFmt")[0]?.getAttribute("formatCode") ??
        null,
    });
  }
  for (const ax of doc.getElementsByTagNameNS(NS.chart, "catAx")) {
    axes.push({
      kind: "cat",
      axId: ax.getElementsByTagNameNS(NS.chart, "axId")[0]?.getAttribute("val"),
    });
  }
  // Series titles
  const series = [];
  for (const ser of doc.getElementsByTagNameNS(NS.chart, "ser")) {
    const tx = ser.getElementsByTagNameNS(NS.chart, "tx")[0];
    const v = tx?.getElementsByTagNameNS(NS.chart, "v")[0]?.textContent ?? null;
    const ref =
      tx?.getElementsByTagNameNS(NS.chart, "f")[0]?.textContent ?? null;
    series.push({ name: v, ref });
  }
  // Map chart? (there's no native node; sometimes detected as bubble or via extLst)
  const isMap = xmlText.includes("MapChart") || xmlText.includes("geoCache");
  return { path, types, trendlines, axes, series, isMap };
}

async function readTable(zip, path) {
  const xmlText = (await zip.file(path)?.async("string")) ?? "";
  if (!xmlText) return null;
  const doc = parseXml(xmlText);
  const t = doc.getElementsByTagNameNS(NS.ssml, "table")[0];
  if (!t) return null;
  const cols = [];
  for (const tc of t.getElementsByTagNameNS(NS.ssml, "tableColumn")) {
    cols.push(tc.getAttribute("name"));
  }
  return {
    name: t.getAttribute("name"),
    displayName: t.getAttribute("displayName"),
    ref: t.getAttribute("ref"),
    columns: cols,
  };
}

async function readPivot(zip, path) {
  const xmlText = (await zip.file(path)?.async("string")) ?? "";
  if (!xmlText) return null;
  const doc = parseXml(xmlText);
  const pt = doc.getElementsByTagNameNS(NS.ssml, "pivotTableDefinition")[0];
  if (!pt) return null;
  const rowFields = Array.from(
    doc.getElementsByTagNameNS(NS.ssml, "rowFields")
  ).flatMap((r) =>
    Array.from(r.getElementsByTagNameNS(NS.ssml, "field")).map((f) => f.getAttribute("x"))
  );
  const colFields = Array.from(
    doc.getElementsByTagNameNS(NS.ssml, "colFields")
  ).flatMap((r) =>
    Array.from(r.getElementsByTagNameNS(NS.ssml, "field")).map((f) => f.getAttribute("x"))
  );
  const dataFields = [];
  for (const df of doc.getElementsByTagNameNS(NS.ssml, "dataField")) {
    dataFields.push({
      name: df.getAttribute("name"),
      fld: df.getAttribute("fld"),
      subtotal: df.getAttribute("subtotal"),
    });
  }
  const fields = [];
  for (const pf of doc.getElementsByTagNameNS(NS.ssml, "pivotField")) {
    fields.push({
      name: pf.getAttribute("name"),
      axis: pf.getAttribute("axis"),
      dataField: pf.getAttribute("dataField") === "1",
    });
  }
  return {
    name: pt.getAttribute("name"),
    cacheId: pt.getAttribute("cacheId"),
    rowFields,
    colFields,
    dataFields,
    fields,
  };
}

async function readSlicer(zip, path) {
  const xmlText = (await zip.file(path)?.async("string")) ?? "";
  if (!xmlText) return null;
  const m = xmlText.match(/<x14:slicer[^>]*name="([^"]+)"[^>]*cache="([^"]+)"/);
  return m ? { name: m[1], cache: m[2] } : { rawXmlPreview: xmlText.slice(0, 200) };
}

// ---------------------------------------------------------------------------
// Power Query detection
// ---------------------------------------------------------------------------

async function detectPowerQuery(zip) {
  // Power Query lives in customXml/itemX.xml as a DataMashup base64 blob,
  // and/or in xl/queries/ + xl/connections.xml for newer versions.
  const items = [];
  zip.forEach((relativePath) => {
    if (relativePath.startsWith("customXml/item") && relativePath.endsWith(".xml")) {
      items.push(relativePath);
    }
  });
  let mashupFound = false;
  let mCodeSnippets = [];
  for (const path of items) {
    const text = await zip.file(path).async("string");
    if (text.includes("DataMashup")) {
      mashupFound = true;
      // Try to decode the embedded zip in the DataMashup base64 blob to extract M code
      const m = text.match(/<\?mso-DataMashup>([^<]+)<\/\?mso-DataMashup>?/);
      // The blob actually appears in <DataMashup ...>BASE64</DataMashup>
      const m2 = text.match(/<DataMashup[^>]*>([\s\S]*?)<\/DataMashup>/);
      const b64 = m2?.[1]?.trim();
      if (b64 && window.JSZip) {
        try {
          const bytes = base64ToBytes(b64);
          // First 4 bytes = version, then a 4-byte length prefix for the embedded zip
          // followed by the zip itself. We just try to find the zip header (PK\x03\x04).
          let start = -1;
          for (let i = 0; i < bytes.length - 4; i++) {
            if (
              bytes[i] === 0x50 &&
              bytes[i + 1] === 0x4b &&
              bytes[i + 2] === 0x03 &&
              bytes[i + 3] === 0x04
            ) {
              start = i;
              break;
            }
          }
          if (start >= 0) {
            const inner = await window.JSZip.loadAsync(bytes.slice(start));
            const mPaths = [];
            inner.forEach((p) => {
              if (p.endsWith(".m") || p.endsWith("Section1.m")) mPaths.push(p);
            });
            for (const p of mPaths) {
              mCodeSnippets.push(await inner.file(p).async("string"));
            }
          }
        } catch (e) {
          // best-effort only
        }
      }
    }
  }
  // Also check xl/queries/
  let queryFiles = [];
  zip.forEach((relativePath) => {
    if (relativePath.startsWith("xl/queries/")) queryFiles.push(relativePath);
  });

  if (!mashupFound && queryFiles.length === 0) return null;
  return {
    detected: true,
    customXmlItems: items,
    queryFiles,
    mCodeSnippets,
  };
}

async function loadConnections(zip) {
  const f = zip.file("xl/connections.xml");
  if (!f) return [];
  const doc = parseXml(await f.async("string"));
  const out = [];
  for (const c of doc.getElementsByTagNameNS(NS.ssml, "connection")) {
    out.push({
      id: c.getAttribute("id"),
      name: c.getAttribute("name"),
      description: c.getAttribute("description"),
      type: c.getAttribute("type"),
      refreshedVersion: c.getAttribute("refreshedVersion"),
    });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function parseXml(text) {
  return new DOMParser().parseFromString(text, "application/xml");
}

function resolveRel(fromPath, target) {
  if (target.startsWith("/")) return target.slice(1);
  const parts = fromPath.split("/").slice(0, -1);
  for (const seg of target.split("/")) {
    if (seg === "..") parts.pop();
    else parts.push(seg);
  }
  return parts.join("/");
}

function base64ToBytes(b64) {
  const clean = b64.replace(/\s+/g, "");
  const bin = atob(clean);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
  return out;
}

// Small helper to build a compact textual workbook summary for the LLM.
export function summarizeForLLM(workbook) {
  const out = {
    definedNames: workbook.definedNames,
    sheetNames: Object.keys(workbook.sheets),
    perSheet: {},
    powerQuery: !!workbook.powerQuery?.detected,
    connections: workbook.connections,
    charts: workbook.charts.map((c) => ({
      types: c.types,
      trendlines: c.trendlines,
      axes: c.axes,
      isMap: c.isMap,
      series: c.series.map((s) => s.name).filter(Boolean),
    })),
    pivotTables: workbook.pivotTables,
    slicers: workbook.slicers,
  };
  for (const [name, sh] of Object.entries(workbook.sheets)) {
    out.perSheet[name] = {
      maxCol: sh.maxCol,
      maxRow: sh.maxRow,
      headers: sh.headers,
      dataValidationCount: sh.dataValidations.length,
      cfCount: sh.conditionalFormattings.length,
    };
  }
  return out;
}

// FULL workbook dump for the bulk grader. Includes every non-empty cell on
// "answer" sheets where students place formulas/values, plus headers + a
// handful of sample rows for pure-data sheets where the body is just source
// records. This keeps the prompt small enough to be cheap on Sonnet/Haiku.
export function fullWorkbookDump(workbook, options = {}) {
  // Per-sheet row caps. Tight on bulk-data sheets, generous on answer sheets.
  // These defaults are aggressive (cost-tuned for Haiku); pass options to
  // loosen them when running on Opus/Sonnet for tricky cases.
  const dataSampleRows = options.dataSampleRows ?? 3;     // section 2 source data: header + 3 sample rows
  const consolidatedSampleRows = options.consolidatedSampleRows ?? 2; // Q10 result table: header + 2 rows (only schema matters)
  const chartSheetMaxRow = options.chartSheetMaxRow ?? 50; // section 2 charts page: top of sheet only
  // PRINT-only sheets contain no answers, just static question text from the
  // template — skip them entirely so we don't bill tokens for what we already
  // have in the rubric.
  const skipSheetRe = /^PRINT/i;
  const out = {
    definedNames: workbook.definedNames,
    powerQuery: workbook.powerQuery
      ? {
          detected: true,
          // Most of the M-code signal is in the first ~1.5KB (let, Source =
          // Excel.CurrentWorkbook, Table.Combine/Merge calls). Trimmed.
          mCodeSnippets: (workbook.powerQuery.mCodeSnippets || [])
            .slice(0, 3)
            .map((s) => s.slice(0, 1800)),
          customXmlItemCount: workbook.powerQuery.customXmlItems?.length ?? 0,
          queryFiles: workbook.powerQuery.queryFiles ?? [],
        }
      : null,
    connections: workbook.connections,
    tables: workbook.tables,
    charts: workbook.charts,
    pivotTables: workbook.pivotTables,
    slicers: workbook.slicers,
    sheets: {},
  };
  for (const [name, sh] of Object.entries(workbook.sheets)) {
    if (skipSheetRe.test(name)) continue;
    const cells = {};
    let cap = sh.maxRow + 1;
    let truncationReason = null;
    if (/section\s*2_DATA/i.test(name)) {
      cap = dataSampleRows + 1;
      truncationReason = "source-data sheet — header + sample rows only";
    } else if (/sect\s*2\s*quesion\s*10/i.test(name)) {
      cap = consolidatedSampleRows + 1;
      truncationReason = "consolidated table — header + few rows (only schema matters)";
    } else if (/section\s*2_\s*quesi?tons?\s*11/i.test(name)) {
      cap = chartSheetMaxRow;
      truncationReason = "chart/pivot page — only top portion needed";
    } else if (/^(Table|Merge|Append)\d+$/i.test(name)) {
      // Power-Query-loaded staging tables that students may leave behind.
      // Only the schema (header row + a couple of samples) matters.
      cap = dataSampleRows + 1;
      truncationReason = "Power-Query staging table — header + sample rows only";
    }
    for (const [ref, cell] of Object.entries(sh.cells)) {
      if (cell.f === undefined && cell.v === undefined) continue;
      const m = /^([A-Z]+)(\d+)$/.exec(ref);
      if (!m) continue;
      const row = parseInt(m[2], 10);
      if (row > cap) continue;
      const c = {};
      if (cell.f) c.f = cell.f;
      if (cell.v !== undefined && cell.v !== null && cell.v !== "") c.v = cell.v;
      if (cell.numFmt) c.fmt = cell.numFmt;
      cells[ref] = c;
    }
    out.sheets[name] = {
      maxCol: sh.maxCol,
      maxRow: sh.maxRow,
      headers: sh.headers,
      cells,
      dataValidations: sh.dataValidations,
      conditionalFormattings: sh.conditionalFormattings,
      truncationReason,
    };
  }
  return out;
}
