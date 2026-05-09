// Rubric for FINAL EXAM v05 2026
// Each question describes what to grade, the max points, and where to look in the workbook.
// The grader sends question + student-evidence + strictness level to Claude and asks for JSON.

export const EXAM_TOTAL = 100;

export const SHEETS = {
  GRADING: "GRADING",
  S1: "section 1 ( 1 to 9)",
  S2_DATA: "section 2_DATA",
  S2_Q10: "sect 2 quesion 10",
  S2_Q11_15: "section 2_ quesitons 11 to 15",
  S3: "section 3 SQL  16 A 20",
  PRINT: "PRINT - All Questions",
};

// Strictness presets — sent to Claude verbatim so it knows how to allocate partial credit.
export const STRICTNESS = {
  lenient: {
    label: "Lenient",
    description:
      "Reward effort. Give full credit if the result is essentially correct even when the formula is unusual or partially manual. Give 50–75% if the answer is close (off by a value, off by formatting, used SUM instead of SUMIF, etc.). Only give 0 when nothing was attempted.",
  },
  normal: {
    label: "Normal",
    description:
      "Standard university grading. Full credit only if the requested method was used and the result is correct. Award 50% if the answer is correct but the requested method was NOT used (e.g. manual numbers instead of SUMIF, hardcoded value instead of LET). Award proportional partial credit for partially correct work.",
  },
  strict: {
    label: "Strict",
    description:
      "No partial credit unless the answer is structurally correct AND uses the exact technique requested. Hardcoded values, manual lists, missing headers, wrong cell references, or missing formula features (LET, SUMIF, etc.) score 0 even if the displayed result is right.",
  },
};

// Reference snapshot of the source-of-truth tables (so Claude can verify SQL outputs).
// Tables are loaded by the app at startup from /reference/*.csv
export const SQL_TABLES = {
  p1: "table1.csv",
  p2: "table2.csv",
  p3: "table3.csv",
};

// ----------------------------------------------------------------------------
// QUESTIONS
// ----------------------------------------------------------------------------
// Each question defines:
//   id, points, section, sheet, label
//   describe(): the question text
//   evidence(workbook): returns relevant slice of the structured workbook
//   expected: what a correct answer looks like (sent to Claude as the rubric)

export const QUESTIONS = [
  // ============== SECTION 1 ==============
  {
    id: "Q1",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Extract Project Name from column A",
    describe: () =>
      "Extract Project Name from column A (which contains noisy strings with codes/spaces/hidden chars) into column B (B2:B56). Result must be a clean project name.",
    expected:
      "Column B should contain the cleaned project name only. Acceptable methods: TRIM/CLEAN/SUBSTITUTE chains, TEXTSPLIT, MID/FIND, LEFT/RIGHT, or any formula. Manual typing scores 0 under STRICT, 30% under NORMAL, full under LENIENT only if values are correct. Result must be a non-empty string per row.",
    evidence: (wb) => sliceCells(wb, SHEETS.S1, ["A1:A60", "B1:B60"]),
  },
  {
    id: "Q2",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Extract Email from column A (lowercase)",
    describe: () =>
      "Extract the email from column A into column C (C2:C56). Output must be lowercase, no extra characters or spaces.",
    expected:
      "Column C must contain a valid lowercase email per row. Must be derived by formula (TEXTSPLIT/MID/FIND/REGEX/LOWER chain). Hardcoding is 0 under STRICT, 30% under NORMAL.",
    evidence: (wb) => sliceCells(wb, SHEETS.S1, ["A1:A60", "C1:C60"]),
  },
  {
    id: "Q3",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "End Date column formatted as m/d/yyyy",
    describe: () =>
      "Column E (End Date) must be formatted as m/d/yyyy (a clean date number-format).",
    expected:
      "Each cell in E2:E56 must have a numFmt that resolves to m/d/yyyy (or equivalent date pattern). Text strings that LOOK like dates do not count.",
    evidence: (wb) => ({
      formats: sliceFormats(wb, SHEETS.S1, "E2:E56"),
      values: sliceCells(wb, SHEETS.S1, ["E1:E60"]),
    }),
  },
  {
    id: "Q4",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Total Duration in days",
    describe: () =>
      "Column F (F2:F56) must compute the total Duration in days = End Date − Start Date.",
    expected:
      "F2:F56 must contain a formula equivalent to =Ex-Dx (or DAYS(Ex,Dx)). Values must be positive integers. Hardcoded numbers = 0 under STRICT.",
    evidence: (wb) =>
      sliceCells(wb, SHEETS.S1, ["D1:D60", "E1:E60", "F1:F60"]),
  },
  {
    id: "Q5",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Working days excluding holidays M10:M29",
    describe: () =>
      "Column G (G2:G56) must compute working days (NETWORKDAYS) between Start Date and End Date, excluding holidays in M10:M29.",
    expected:
      "G2:G56 must use NETWORKDAYS or NETWORKDAYS.INTL with the holiday range $M$10:$M$29 (absolute reference required for full credit). Plain DAYS / E-D = 0–25%.",
    evidence: (wb) =>
      sliceCells(wb, SHEETS.S1, [
        "D1:D60",
        "E1:E60",
        "G1:G60",
        "M9:M30",
      ]),
  },
  {
    id: "Q6",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Modules per Category summary (SUMIF/SUMIFS)",
    describe: () =>
      "There is a small summary table on the sheet listing total modules per category. Cells must use SUMIF / SUMIFS — not manual numbers.",
    expected:
      "Each summary cell must contain a SUMIF or SUMIFS formula referencing the Modules and Category columns. Manual numeric values = 0 under STRICT, 25% under NORMAL.",
    evidence: (wb) =>
      sliceCells(wb, SHEETS.S1, ["I1:N40", "A1:K60"]),
  },
  {
    id: "Q7",
    section: 1,
    points: 2,
    sheet: SHEETS.S1,
    label: "Data validation drop-downs for Category and Status",
    describe: () =>
      "Two data validation list drop-downs: Category restricted to {Mobile, Web} and Status restricted to {Active, Completed}.",
    expected:
      "The sheet's dataValidations must contain at least two list-type rules whose formula1 contains 'Mobile' & 'Web' for one and 'Active' & 'Completed' for the other. Otherwise 0.",
    evidence: (wb) => ({
      dataValidations: wb.sheets?.[SHEETS.S1]?.dataValidations ?? [],
    }),
  },
  {
    id: "Q8",
    section: 1,
    points: 4,
    sheet: SHEETS.S1,
    label: "Cross-hair conditional formatting (active row + active column)",
    describe: () =>
      "Apply conditional formatting on A1:K56 that highlights the entire row AND the entire column of the currently active cell (cross-hair). Typical implementation uses two CF rules with formulas like =ROW()=CELL(\"row\") and =COLUMN()=CELL(\"col\").",
    expected:
      "Sheet conditionalFormattings must contain at least 2 formula-type CF rules referencing CELL(\"row\") / CELL(\"col\") (or ROW()/COLUMN() comparisons that achieve cross-hair). Single rule highlighting only one axis = 50%.",
    evidence: (wb) => ({
      conditionalFormatting:
        wb.sheets?.[SHEETS.S1]?.conditionalFormattings ?? [],
    }),
  },
  {
    id: "Q9",
    section: 1,
    points: 4,
    sheet: SHEETS.S1,
    label: "Category with HIGHEST total Revenue (Start Date in 2023)",
    describe: () =>
      "Find the Category with the highest total Revenue among projects whose Start Date is in 2023. Formulas only — no PivotTable, no Excel Tables.",
    expected:
      "There must be a labelled cell on the sheet containing the answer derived by formulas (SUMIFS over Category and YEAR(StartDate)=2023, then INDEX/MATCH or XLOOKUP on the MAX). Manual or pivot-derived = 0 under STRICT.",
    evidence: (wb) =>
      sliceCells(wb, SHEETS.S1, ["I1:R40", "A1:K60"]),
  },
  {
    id: "Q9b",
    section: 1,
    points: 3,
    sheet: SHEETS.S1,
    label: "LET function: Category with highest AVG Revenue in 2023",
    describe: () =>
      "Using the LET function, compute the AVERAGE Revenue per Category for projects whose Start Date is in 2023, and return the Category with the highest average. Must use LET to name intermediate variables in a single final expression. No helper cells.",
    expected:
      "Cell must contain a single formula starting with =LET(...) with at least 2 named variables and a final expression that returns a Category name. If LET is missing → 0.",
    evidence: (wb) =>
      sliceCells(wb, SHEETS.S1, ["I1:R40", "A1:K60"]),
  },

  // ============== SECTION 2 ==============
  {
    id: "Q10",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q10,
    label: "Power Query consolidated table (23 fields)",
    describe: () =>
      "Use POWER QUERY ONLY to combine Tables 1–7 from 'section 2_DATA' into a single consolidated table on a new sheet. Append Table6+Table7, add 'DeliveryYear' = YEAR(ActualDelivery), merge through ProductID/OrderID/CustomerID/ShipmentID/SupplierID. Final table must contain EXACTLY these 23 fields:\n  ProductID, ProductName, Category, UnitPrice, CustomerID, CustomerName, CustomerCountry, OrderDate, ShipmentID, DispatchDate, EstimatedDelivery, ActualDelivery, ShippingCarrier, SupplierID, SupplierName, SupplierCountry, Quantity, SalePrice, TotalAmount, Discount, FinalAmount, DeliveryYear, PaymentMethod",
    expected:
      "Workbook must contain Power Query connections (DataMashup / customXml). The sheet 'sect 2 quesion 10' (or a new sheet) must contain a loaded table whose header row matches the 23 required fields exactly (order may vary). Column count and presence of DeliveryYear are critical.",
    evidence: (wb) => ({
      powerQuery: wb.powerQuery ?? null,
      sheetHeaders: wb.sheets?.[SHEETS.S2_Q10]?.headers ?? [],
      sheetCells: sliceCells(wb, SHEETS.S2_Q10, ["A1:Z3"]),
      hasConnections: !!wb.connections?.length,
      tablesInBook: wb.tables ?? [],
    }),
  },
  {
    id: "Q11",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q11_15,
    label: "Time series + linear trendline (R², 3-period forecast)",
    describe: () =>
      "Build a PivotTable from the Data Model (Rows=OrderDate grouped by Month, Values=Sum(Quantity)), then a line chart with a LINEAR trendline showing R² and forecasting forward 3 periods.",
    expected:
      "Workbook must contain at least one pivotTable on the Data Model AND at least one chart (lineChart) with a trendline whose trendlineType = 'linear', dispRSqr = true, and forecast forward = 3.",
    evidence: (wb) => ({
      pivotTables: wb.pivotTables ?? [],
      charts: wb.charts ?? [],
    }),
  },
  {
    id: "Q12",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q11_15,
    label: "Column chart of TotalAmount by Category + 3 connected slicers",
    describe: () =>
      "PivotTable Rows=Category, Values=Sum(TotalAmount). Column/bar chart from this pivot. Three slicers (CustomerCountry, ShippingCarrier, PaymentMethod) all connected to the pivot.",
    expected:
      "Must have a column/bar chart bound to a pivot with Category rows and TotalAmount values, plus 3 slicers in the workbook for those exact field names, all connected to the same pivot.",
    evidence: (wb) => ({
      pivotTables: wb.pivotTables ?? [],
      charts: wb.charts ?? [],
      slicers: wb.slicers ?? [],
    }),
  },
  {
    id: "Q13",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q11_15,
    label: "Geographic Map chart with 2 shared slicers",
    describe: () =>
      "Map chart from a pivot on the Data Model: Locations=CustomerCountry (or SupplierCountry), Values=Sum(FinalAmount). Two slicers (Category, PaymentMethod) connected to BOTH this pivot and Q14's pivot.",
    expected:
      "Must have a map/geographic chart and two slicers (Category, PaymentMethod) whose pivotTable connections include both Q13's and Q14's pivots.",
    evidence: (wb) => ({
      pivotTables: wb.pivotTables ?? [],
      charts: wb.charts ?? [],
      slicers: wb.slicers ?? [],
    }),
  },
  {
    id: "Q14",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q11_15,
    label: "Bubble chart Quantity vs UnitPrice vs TotalAmount + shared slicers",
    describe: () =>
      "PivotTable Rows=ProductName, Values: Sum(Quantity) → X, Avg(UnitPrice) → Y, Sum(TotalAmount) → bubble size. Bubble chart from this pivot. Slicers Category & PaymentMethod connected to BOTH Q13 and Q14 pivots.",
    expected:
      "Must have a bubble chart referencing those three measures and slicers (Category, PaymentMethod) connected to this pivot AND Q13's pivot.",
    evidence: (wb) => ({
      pivotTables: wb.pivotTables ?? [],
      charts: wb.charts ?? [],
      slicers: wb.slicers ?? [],
    }),
  },
  {
    id: "Q15",
    section: 2,
    points: 10,
    sheet: SHEETS.S2_Q11_15,
    label: "Combo chart: Quantity (column) + FinalAmount (line, secondary axis)",
    describe: () =>
      "Combo chart from a pivot on the Data Model: Rows=Category, Values=Sum(Quantity) clustered column on PRIMARY axis + Sum(FinalAmount) line on SECONDARY axis. Currency formatting on the secondary axis. Data labels on the LINE series only.",
    expected:
      "Must have a combo chart with at least one barChart/columnChart series (primary axis) and at least one lineChart series on a secondary axis (axId differs). Insight text box optional but counts toward presentation.",
    evidence: (wb) => ({
      pivotTables: wb.pivotTables ?? [],
      charts: wb.charts ?? [],
    }),
  },

  // ============== SECTION 3 - SQL (graded from screenshots, not workbook) ==============
  {
    id: "Q16",
    section: 3,
    points: 3,
    label: "COUNT DISTINCT genres in p2",
    describe: () =>
      "Count the number of DISTINCT genres in Table 2 (p2). Return a single number with column alias 'genre_count'.",
    expected:
      "Query: SELECT COUNT(DISTINCT genre) AS genre_count FROM p2;\nExpected result: 10 distinct genres (Pop, Jazz, Rock, Blues, EDM, Country, Rap, Latin, Folk, Indie).\nFull credit if alias is exactly genre_count and the result is 10.",
    sql: true,
  },
  {
    id: "Q17",
    section: 3,
    points: 3,
    label: "INNER JOIN p1 + p2, Sales > 520000",
    describe: () =>
      "INNER JOIN p1 and p2. Return ID, Status, Sales, Genre WHERE Sales > 520000.",
    expected:
      "Query must INNER JOIN p1 and p2 on ID, project Status from p1, Sales+Genre from p2, filtered by Sales > 520000. Result must show only rows with Sales 521000..524000.",
    sql: true,
  },
  {
    id: "Q18",
    section: 3,
    points: 3,
    label: "Add p3, filter UK + Sales > 520000",
    describe: () =>
      "INNER JOIN previous result with p3. Return Project Name, Status, Sales, Genre, Country WHERE Sales > 520000 AND Country = 'UK'.",
    expected:
      "Three-table inner join on ID and country chain (artist_Location -> p3.country). Final result must only include rows where Country='UK' AND Sales>520000. Expected: rows with sales=521000 (Jazz, Jazz Collection).",
    sql: true,
  },
  {
    id: "Q19",
    section: 3,
    points: 3,
    label: "All 3 tables, Sales > 600000 AND Status = 'Active', ORDER BY Sales DESC",
    describe: () =>
      "INNER JOIN all THREE tables (p1, p2, p3) in a single query. Return Project Name, Status, Sales, Genre, Country WHERE Sales > 600000 AND Status = 'Active'. Order results by Sales DESC.",
    expected:
      "Single query with three inner joins, two filters, and ORDER BY Sales DESC. Looking at the data, no rows in the supplied tables have Sales > 600000 (max is 524000) so result is empty — query must STILL be syntactically correct and the WHERE clause precise.",
    sql: true,
  },
  {
    id: "Q20",
    section: 3,
    points: 3,
    label: "LEFT vs RIGHT vs FULL OUTER JOIN comparison + explanation",
    describe: () =>
      "Three separate queries (LEFT, RIGHT, FULL OUTER) between p1 and p2 on ID, returning ID, Project Name, Genre. Plus a 1–2 sentence written explanation of why row counts differ and what NULLs mean.",
    expected:
      "Three syntactically valid queries (SQLite uses LEFT and FULL OUTER; RIGHT requires emulation). Written explanation must mention: LEFT keeps all p1 rows (NULL where p2 missing), RIGHT keeps all p2 rows, FULL keeps both. NULLs indicate the absent side.",
    sql: true,
  },
];

// ---------------- Helpers used by evidence() ---------------------------------

function sliceCells(workbook, sheetName, ranges) {
  const sheet = workbook?.sheets?.[sheetName];
  if (!sheet) return { _missing: sheetName };
  const out = {};
  for (const range of ranges) {
    out[range] = pluckRange(sheet.cells, range);
  }
  return out;
}

function sliceFormats(workbook, sheetName, range) {
  const sheet = workbook?.sheets?.[sheetName];
  if (!sheet) return { _missing: sheetName };
  return pluckRangeFormats(sheet.cells, range);
}

function pluckRange(cells, range) {
  const [start, end] = range.split(":");
  const [c1, r1] = parseRef(start);
  const [c2, r2] = parseRef(end ?? start);
  const out = {};
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const ref = colLetter(c) + r;
      const cell = cells?.[ref];
      if (cell) out[ref] = compactCell(cell);
    }
  }
  return out;
}

function pluckRangeFormats(cells, range) {
  const [start, end] = range.split(":");
  const [c1, r1] = parseRef(start);
  const [c2, r2] = parseRef(end ?? start);
  const out = {};
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) {
      const ref = colLetter(c) + r;
      const cell = cells?.[ref];
      if (cell?.numFmt || cell?.styleId !== undefined) {
        out[ref] = { numFmt: cell.numFmt ?? null, styleId: cell.styleId };
      }
    }
  }
  return out;
}

function compactCell(cell) {
  const o = {};
  if (cell.f) o.f = cell.f;
  if (cell.v !== undefined && cell.v !== null && cell.v !== "") o.v = cell.v;
  if (cell.numFmt) o.fmt = cell.numFmt;
  return o;
}

export function parseRef(ref) {
  const m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref);
  if (!m) return [0, 0];
  return [colNum(m[1]), parseInt(m[2], 10)];
}

export function colNum(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

export function colLetter(n) {
  let s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
