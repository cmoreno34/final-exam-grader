# FINAL EXAM Grader

A static, browser-only grading app that uses Claude (Anthropic API) to grade
the FINAL EXAM (Excel + SQL) the same way Claude does in chat.

- **Excel grading** — unzips each `.xlsx`, parses cells, formulas, number-formats,
  data-validations, conditional formatting, charts, pivots, slicers, defined names
  and Power Query connections, then asks Claude to grade each rubric question.
- **SQL grading** — sends student screenshots to Claude with vision plus the
  source CSV tables, and asks for one verdict per question.
- **Strictness slider** — Lenient, Normal, Strict.
- **Output** — a single `grading-results.zip` containing:
  - `00_grades.xlsx` — master sheet with totals and per-question scores
  - `00_grades.csv` — same as CSV
  - `<student>/<name>_GRADED.xlsx` — the student's original workbook **with
    a new `GRADING (auto)` sheet injected inside**. The sheet shows max points,
    awarded points, verdict, evidence and feedback per question. The total is
    a `=SUM(...)` formula, so if you edit any *Awarded* cell manually the
    total updates automatically.
  - `<student>/feedback.md` — same content as a markdown report.

Everything runs in your browser. Files never go anywhere except the Anthropic
API endpoint.

## Deploy on GitHub Pages (account `cmoreno34`)

```powershell
# 1) Initialize a repo from this folder
cd "C:\Users\Usuario\OneDrive - UFV\32.- SLU BTM\FINAL\grader-web"
git init -b main
git add -A
git commit -m "FINAL exam grader"

# 2) Create the repo on GitHub (replace REPO_NAME if you like)
gh repo create cmoreno34/final-exam-grader --public --source=. --push
# (or, without gh CLI:)
#   create the repo at https://github.com/new
#   git remote add origin https://github.com/cmoreno34/final-exam-grader.git
#   git push -u origin main

# 3) Enable Pages
#    Settings → Pages → Source: Deploy from a branch → main → / (root) → Save
```

After ~1 minute the site will be live at:

```
https://cmoreno34.github.io/final-exam-grader/
```

## Local development

The app is plain HTML + ES modules. Any static server works:

```powershell
cd "C:\Users\Usuario\OneDrive - UFV\32.- SLU BTM\FINAL\grader-web"
python -m http.server 8000
# open http://localhost:8000
```

You can also just open `index.html` over `file://`, but the reference CSV
fetches will fail in some browsers.

## How to use

1. Paste your Anthropic API key (`sk-ant-…`). Click **Test key**. The key
   stays in `localStorage` on this device — it is never sent anywhere except
   `api.anthropic.com`.
2. Pick a strictness level.
3. Drop one or more files:
   - a single `.xlsx`
   - a `.zip` of many `.xlsx` (one per student)
   - a `.zip` with one folder per student containing the student's `.xlsx`
     and any `.png` / `.jpg` SQL screenshots
   - or loose `.png` / `.jpg` screenshots (will be matched to a student by
     filename prefix)
4. Click **Start grading** and wait. Claude is called once per Excel question
   per student plus once for the SQL screenshots.
5. Click **Download grading-results.zip**.

## Cost guidance

Per student, the grader makes ~16 short Claude calls (Q1–Q15 + a single SQL
multimodal call). With Opus 4.7 expect roughly $0.20–$0.50 per student;
Sonnet 4.6 is ~5× cheaper. You can switch model in the UI.

## Files

```
grader-web/
├── index.html
├── styles.css
├── app.js                  # UI orchestrator
├── lib/
│   ├── rubric.js           # the 20-question rubric
│   ├── extractor.js        # xlsx → structured JSON
│   ├── api.js              # Claude API wrapper
│   ├── grader.js           # rubric × workbook → grades
│   └── output.js           # builds the result zip
├── reference/
│   ├── template.xlsx       # the source-of-truth template
│   ├── table1.csv          # p1 for SQL grading
│   ├── table2.csv          # p2 for SQL grading
│   └── table3.csv          # p3 for SQL grading
├── .nojekyll               # disable Jekyll on GitHub Pages
└── README.md
```

## Limitations

- Microsoft `.xlsb` and `.xls` are not supported (xlsx only).
- Some Power Query M code is wrapped inside an undocumented `DataMashup`
  base64 blob; the extractor does its best to decode it but the grader can
  also rely on the structural evidence (presence of connections, headers of
  the loaded table, the 23 expected fields).
- Charts that Excel saves as embedded `.bin` (legacy) cannot be inspected.
- Cross-hair conditional formatting is detected by inspecting the CF rule
  formulas; if a student used a non-standard pattern the grader's evidence
  will be limited and Claude is asked to judge accordingly.
