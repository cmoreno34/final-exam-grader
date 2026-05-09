// Main application: wires UI to extractor + grader + output.

import { extractWorkbook } from "./lib/extractor.js";
import { gradeWorkbook, gradeSqlScreenshots } from "./lib/grader.js";
import { buildOutputZip } from "./lib/output.js";
import { QUESTIONS, EXAM_TOTAL } from "./lib/rubric.js";
import { callClaude } from "./lib/api.js";

const els = {
  apiKey: document.getElementById("apiKey"),
  model: document.getElementById("model"),
  testKey: document.getElementById("testKey"),
  keyStatus: document.getElementById("keyStatus"),
  dropzone: document.getElementById("dropzone"),
  fileInput: document.getElementById("fileInput"),
  browseBtn: document.getElementById("browseBtn"),
  fileList: document.getElementById("fileList"),
  gradeBtn: document.getElementById("gradeBtn"),
  cancelBtn: document.getElementById("cancelBtn"),
  progress: document.getElementById("progress"),
  log: document.getElementById("log"),
  resultsCard: document.getElementById("resultsCard"),
  resultsTable: document.getElementById("resultsTable"),
  downloadBtn: document.getElementById("downloadBtn"),
};

const state = {
  files: [],
  results: null,
  zipBlob: null,
  abort: null,
  referenceTablesText: null,
};

// --- API key persistence -----------------------------------------------------
els.apiKey.value = localStorage.getItem("anthropic_key") ?? "";
els.model.value = localStorage.getItem("anthropic_model") ?? "claude-opus-4-7";
els.apiKey.addEventListener("change", () =>
  localStorage.setItem("anthropic_key", els.apiKey.value.trim())
);
els.model.addEventListener("change", () =>
  localStorage.setItem("anthropic_model", els.model.value)
);

els.testKey.addEventListener("click", async () => {
  const key = els.apiKey.value.trim();
  if (!key) {
    els.keyStatus.textContent = "Enter a key first.";
    return;
  }
  els.keyStatus.textContent = "Testing…";
  try {
    await callClaude({
      apiKey: key,
      model: els.model.value,
      messages: [{ role: "user", content: [{ type: "text", text: "ping" }] }],
      maxTokens: 8,
    });
    els.keyStatus.innerHTML = "<span style=\"color:var(--good)\">Key works.</span>";
    localStorage.setItem("anthropic_key", key);
  } catch (err) {
    els.keyStatus.innerHTML =
      '<span style="color:var(--bad)">Key failed: ' +
      escapeHtml((err.message ?? String(err)).slice(0, 200)) +
      "</span>";
  }
});

// --- File input --------------------------------------------------------------
els.browseBtn.addEventListener("click", () => els.fileInput.click());
els.fileInput.addEventListener("change", (e) => addFiles(e.target.files));

["dragenter", "dragover"].forEach((evt) =>
  els.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    els.dropzone.classList.add("dragover");
  })
);
["dragleave", "drop"].forEach((evt) =>
  els.dropzone.addEventListener(evt, (e) => {
    e.preventDefault();
    els.dropzone.classList.remove("dragover");
  })
);
els.dropzone.addEventListener("drop", (e) => addFiles(e.dataTransfer.files));

function addFiles(fileList) {
  for (const f of fileList) state.files.push(f);
  renderFileList();
  els.gradeBtn.disabled = state.files.length === 0;
}
function renderFileList() {
  els.fileList.innerHTML = state.files
    .map(
      (f, i) =>
        `<li>${escapeHtml(f.name)} <span class="muted small">(${formatSize(f.size)})</span> <button data-rm="${i}" class="link">remove</button></li>`
    )
    .join("");
  els.fileList.querySelectorAll("button[data-rm]").forEach((b) =>
    b.addEventListener("click", () => {
      state.files.splice(Number(b.dataset.rm), 1);
      renderFileList();
      els.gradeBtn.disabled = state.files.length === 0;
    })
  );
}

// --- Grading -----------------------------------------------------------------
els.gradeBtn.addEventListener("click", grade);
els.cancelBtn.addEventListener("click", () => state.abort?.abort());

async function grade() {
  const apiKey = els.apiKey.value.trim();
  if (!apiKey) {
    showLog("Enter and save your Claude API key first.", "err");
    return;
  }
  const model = els.model.value;
  const strictness =
    document.querySelector('input[name="strictness"]:checked')?.value ?? "normal";

  els.gradeBtn.disabled = true;
  els.cancelBtn.disabled = false;
  els.log.classList.add("visible");
  els.log.innerHTML = "";
  els.resultsCard.hidden = true;
  state.abort = new AbortController();

  try {
    if (!state.referenceTablesText) {
      state.referenceTablesText = await loadReferenceTables();
    }
    showLog("Unpacking student submissions…");
    const submissions = await splitIntoStudentSubmissions(state.files);
    showLog(`Found ${submissions.length} student submission(s).`, "ok");

    const startedAt = new Date().toISOString();
    const studentResults = [];
    setProgress(0);
    let perStudentSteps = QUESTIONS.length; // Excel + 1 SQL pass
    let totalSteps = submissions.length * (perStudentSteps + 1);
    let stepsDone = 0;

    for (const sub of submissions) {
      showLog(`▶ ${sub.studentName}: parsing workbook…`);
      const workbook = await extractWorkbook(sub.xlsxBytes);

      const excelResults = await gradeWorkbook({
        workbook,
        apiKey,
        model,
        strictness,
        studentName: sub.studentName,
        signal: state.abort.signal,
        onProgress: (e) => {
          stepsDone++;
          setProgress((stepsDone / totalSteps) * 100);
          showLog(
            `   · ${e.questionId} (${e.index + 1}/${e.total})`,
            "muted"
          );
        },
      });

      showLog(`   · grading SQL screenshots (${sub.screenshots.length} image(s))…`);
      const sqlResults = await gradeSqlScreenshots({
        screenshots: sub.screenshots,
        apiKey,
        model,
        strictness,
        studentName: sub.studentName,
        signal: state.abort.signal,
        referenceTables: state.referenceTablesText,
      });
      stepsDone++;
      setProgress((stepsDone / totalSteps) * 100);

      const allResults = [...excelResults, ...sqlResults];
      const totalPoints = allResults.reduce(
        (a, r) => a + (r.points_awarded || 0),
        0
      );
      studentResults.push({
        studentName: sub.studentName,
        results: allResults,
        totalPoints,
        gradedAt: new Date().toISOString(),
        originalFile: { name: sub.originalName, bytes: sub.xlsxBytes },
      });
      showLog(
        `   = ${sub.studentName}: ${totalPoints.toFixed(1)} / ${EXAM_TOTAL}`,
        "ok"
      );
    }

    setProgress(100);
    showLog("Building results zip…");
    const finishedAt = new Date().toISOString();
    const zipBlob = await buildOutputZip({
      studentResults,
      runMeta: { strictness, model, startedAt, finishedAt },
    });
    state.results = studentResults;
    state.zipBlob = zipBlob;
    renderResultsTable(studentResults);
    els.resultsCard.hidden = false;
    showLog("Done. Click Download.", "ok");
  } catch (err) {
    if (err.name === "AbortError") showLog("Cancelled.", "warn");
    else showLog("Error: " + (err.message ?? String(err)), "err");
  } finally {
    els.gradeBtn.disabled = state.files.length === 0;
    els.cancelBtn.disabled = true;
    state.abort = null;
  }
}

els.downloadBtn.addEventListener("click", () => {
  if (!state.zipBlob) return;
  window.saveAs(state.zipBlob, "grading-results.zip");
});

// --- Submission splitting ----------------------------------------------------
// Accepts a mix of: a single .xlsx, a .zip with many .xlsx, individual screenshots.
// Returns array of { studentName, originalName, xlsxBytes (Uint8Array), screenshots: [{base64, mediaType}] }

async function splitIntoStudentSubmissions(files) {
  const submissions = new Map();

  const ensure = (name) => {
    if (!submissions.has(name))
      submissions.set(name, {
        studentName: name,
        originalName: null,
        xlsxBytes: null,
        screenshots: [],
      });
    return submissions.get(name);
  };

  for (const file of files) {
    const lower = file.name.toLowerCase();
    if (lower.endsWith(".zip")) {
      const inner = await window.JSZip.loadAsync(file);
      const entries = [];
      inner.forEach((relPath, entry) => {
        if (!entry.dir) entries.push({ relPath, entry });
      });
      // Group by top-level folder. Entries directly at the root each become
      // their own student (one xlsx per student). A folder containing an
      // xlsx + images is a single student.
      const groups = new Map();
      for (const f of entries) {
        const parts = f.relPath.split("/");
        let key;
        if (parts.length > 1) key = parts[0];
        else if (f.relPath.toLowerCase().endsWith(".xlsx")) key = stripExt(f.relPath);
        else key = "loose-screenshots-" + parts[0];
        if (!groups.has(key)) groups.set(key, []);
        groups.get(key).push(f);
      }
      for (const [groupName, members] of groups.entries()) {
        const xlsx = members.find((m) => m.relPath.toLowerCase().endsWith(".xlsx"));
        const imgs = members.filter((m) => /\.(png|jpe?g)$/i.test(m.relPath));
        if (!xlsx && imgs.length === 0) continue;
        const sub = ensure(groupName);
        if (xlsx) {
          sub.xlsxBytes = await xlsx.entry.async("uint8array");
          sub.originalName = xlsx.relPath.split("/").pop();
        }
        for (const img of imgs) {
          const data = await img.entry.async("base64");
          sub.screenshots.push({
            base64: data,
            mediaType: img.relPath.toLowerCase().endsWith(".png")
              ? "image/png"
              : "image/jpeg",
          });
        }
      }
    } else if (lower.endsWith(".xlsx")) {
      const studentName = stripExt(file.name);
      const sub = ensure(studentName);
      sub.xlsxBytes = new Uint8Array(await file.arrayBuffer());
      sub.originalName = file.name;
    } else if (/\.(png|jpe?g)$/i.test(lower)) {
      // Loose screenshot — attach to a "screenshots" pseudo-student or, better,
      // try matching by name prefix to an existing student.
      const stem = stripExt(file.name);
      const matched =
        [...submissions.keys()].find((n) =>
          stem.toLowerCase().includes(n.toLowerCase())
        ) ?? "loose-screenshots";
      const sub = ensure(matched);
      const data = arrayBufferToBase64(await file.arrayBuffer());
      sub.screenshots.push({
        base64: data,
        mediaType: lower.endsWith(".png") ? "image/png" : "image/jpeg",
      });
    }
  }
  // Drop submissions with no xlsx (just loose screenshots without a workbook).
  return [...submissions.values()].filter((s) => s.xlsxBytes);
}

// --- Reference tables --------------------------------------------------------
async function loadReferenceTables() {
  const fetchCsv = async (path) => {
    const r = await fetch(path);
    if (!r.ok) return null;
    return await r.text();
  };
  const [t1, t2, t3] = await Promise.all([
    fetchCsv("reference/table1.csv"),
    fetchCsv("reference/table2.csv"),
    fetchCsv("reference/table3.csv"),
  ]);
  return [
    "p1 (Table 1):\n" + (t1 ?? "(missing)"),
    "p2 (Table 2):\n" + (t2 ?? "(missing)"),
    "p3 (Table 3):\n" + (t3 ?? "(missing)"),
  ].join("\n\n");
}

// --- UI helpers --------------------------------------------------------------
function showLog(msg, kind) {
  const cls = kind === "ok" ? "ok" : kind === "warn" ? "warn" : kind === "err" ? "err" : "";
  const div = document.createElement("div");
  if (cls) div.className = cls;
  div.textContent = msg;
  els.log.appendChild(div);
  els.log.scrollTop = els.log.scrollHeight;
}
function setProgress(pct) {
  if (!els.progress.querySelector(".bar")) {
    const bar = document.createElement("div");
    bar.className = "bar";
    els.progress.appendChild(bar);
  }
  els.progress.querySelector(".bar").style.width = pct.toFixed(1) + "%";
}

function renderResultsTable(students) {
  const head = `<thead><tr><th>Student</th><th>Total /100</th>` +
    QUESTIONS.map((q) => `<th>${q.id}</th>`).join("") +
    `</tr></thead>`;
  const body =
    "<tbody>" +
    students
      .map((s) => {
        const cells = QUESTIONS.map((q) => {
          const r = s.results.find((x) => x.questionId === q.id);
          const v = r ? r.points_awarded.toFixed(1) : "0";
          return `<td title="${escapeHtml(r?.feedback ?? '')}">${v}</td>`;
        }).join("");
        return `<tr><td>${escapeHtml(s.studentName)}</td><td class="score">${s.totalPoints.toFixed(1)}</td>${cells}</tr>`;
      })
      .join("") +
    "</tbody>";
  els.resultsTable.innerHTML = head + body;
}

function escapeHtml(s) {
  return (s ?? "")
    .toString()
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}
function stripExt(name) {
  return name.replace(/\.[^.]+$/, "");
}
function formatSize(bytes) {
  if (bytes < 1024) return bytes + " B";
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + " KB";
  return (bytes / 1024 / 1024).toFixed(2) + " MB";
}
function arrayBufferToBase64(buf) {
  const bytes = new Uint8Array(buf);
  let bin = "";
  for (let i = 0; i < bytes.length; i++) bin += String.fromCharCode(bytes[i]);
  return btoa(bin);
}
