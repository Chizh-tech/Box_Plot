/* =========================================================
   Box Plot Tool – Frontend Logic
   ========================================================= */

"use strict";

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------
const state = {
  files: [],          // [{fileId, filename, stageName, sheetNames, selectedSheet, config, dimensions}]
  dimOverrides: {},   // {dimName: {displayName, nominal, usl, lsl}}
  plotData: null,     // raw data from /api/plot_data
  configTarget: null, // index into state.files being configured
};

// ---------------------------------------------------------------------------
// Bootstrap modal handles
// ---------------------------------------------------------------------------
let configModal, dimEditorModal;

document.addEventListener("DOMContentLoaded", () => {
  configModal    = new bootstrap.Modal(document.getElementById("configModal"));
  dimEditorModal = new bootstrap.Modal(document.getElementById("dimEditorModal"));
  initUpload();
});

// ---------------------------------------------------------------------------
// Drag-and-drop / click upload
// ---------------------------------------------------------------------------
function initUpload() {
  const zone  = document.getElementById("dropZone");
  const input = document.getElementById("fileInput");

  zone.addEventListener("click", () => input.click());
  input.addEventListener("change", () => handleFiles(Array.from(input.files)));

  zone.addEventListener("dragover",  (e) => { e.preventDefault(); zone.classList.add("drag-over"); });
  zone.addEventListener("dragleave", ()  => zone.classList.remove("drag-over"));
  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    zone.classList.remove("drag-over");
    handleFiles(Array.from(e.dataTransfer.files));
  });
}

async function handleFiles(files) {
  for (const f of files) {
    await uploadFile(f);
  }
}

async function uploadFile(file) {
  const fd = new FormData();
  fd.append("file", file);

  let resp;
  try {
    resp = await fetch("/api/upload", { method: "POST", body: fd });
  } catch (err) {
    alert("上传失败：" + err.message);
    return;
  }
  const data = await resp.json();
  if (data.error) { alert("上传错误：" + data.error); return; }

  const entry = {
    fileId:        data.file_id,
    filename:      data.filename,
    stageName:     "阶段 " + (state.files.length + 1),
    sheetNames:    data.sheet_names,
    selectedSheet: data.sheet_names[0] || "",
    config:        defaultConfig(),
    dimensions:    [],
  };
  state.files.push(entry);
  renderFileList();
  showEmptyState(false);
}

function defaultConfig() {
  return {
    format: "col_per_dim",
    header_row: 0,      // 0-based
    skip_cols: 0,
    nominal_row: null,
    usl_row: null,
    lsl_row: null,
    upper_tol_row: null,
    lower_tol_row: null,
    data_start_row: 1,
    // row_per_dim
    name_col: 0,
    first_data_row: 1,
    nominal_col: null,
    usl_col: null,
    lsl_col: null,
    upper_tol_col: null,
    lower_tol_col: null,
    data_start_col: 4,
  };
}

// ---------------------------------------------------------------------------
// Render uploaded file list
// ---------------------------------------------------------------------------
function renderFileList() {
  const container = document.getElementById("fileList");
  container.innerHTML = "";

  state.files.forEach((entry, idx) => {
    const div = document.createElement("div");
    div.className = "file-item";
    div.innerHTML = `
      <div class="d-flex justify-content-between align-items-start mb-1">
        <span class="file-name"><i class="bi bi-file-earmark-excel text-success me-1"></i>${esc(entry.filename)}</span>
        <button class="btn btn-sm btn-link text-danger p-0 ms-1" onclick="removeFile(${idx})" title="删除">
          <i class="bi bi-x-lg"></i>
        </button>
      </div>
      <div class="mb-1">
        <label class="form-label small mb-0">阶段名称</label>
        <input type="text" class="form-control form-control-sm stage-input"
               value="${esc(entry.stageName)}"
               onchange="state.files[${idx}].stageName=this.value" />
      </div>
      <div class="mb-2">
        <label class="form-label small mb-0">选择 Sheet</label>
        <select class="form-select form-select-sm" onchange="onSheetChange(${idx},this.value)">
          ${entry.sheetNames.map(s => `<option value="${esc(s)}" ${s===entry.selectedSheet?'selected':''}>${esc(s)}</option>`).join("")}
        </select>
      </div>
      <button class="btn btn-sm btn-outline-primary w-100" onclick="openConfig(${idx})">
        <i class="bi bi-gear me-1"></i>配置数据格式 / 提取尺寸
      </button>
      ${entry.dimensions.length > 0
        ? `<div class="mt-1 text-success small"><i class="bi bi-check-circle me-1"></i>已提取 ${entry.dimensions.length} 个尺寸</div>`
        : ""}
    `;
    container.appendChild(div);
  });
}

function removeFile(idx) {
  state.files.splice(idx, 1);
  renderFileList();
  mergeDimensions();
  if (state.files.length === 0) showEmptyState(true);
}

function onSheetChange(idx, sheet) {
  state.files[idx].selectedSheet = sheet;
  state.files[idx].dimensions = [];
  renderFileList();
  mergeDimensions();
}

// ---------------------------------------------------------------------------
// Config modal
// ---------------------------------------------------------------------------
function openConfig(idx) {
  state.configTarget = idx;
  const entry = state.files[idx];

  document.getElementById("configModalFilename").textContent = entry.filename;

  // Restore config into form
  const cfg = entry.config;
  const fmt = cfg.format || "col_per_dim";
  document.getElementById(fmt === "col_per_dim" ? "fmtColPerDim" : "fmtRowPerDim").checked = true;

  // col_per_dim fields (1-based display)
  setVal("cfgHeaderRow",    toDisplay(cfg.header_row));
  setVal("cfgSkipCols",     cfg.skip_cols ?? 0);
  setVal("cfgNominalRow",   toDisplay(cfg.nominal_row));
  setVal("cfgUslRow",       toDisplay(cfg.usl_row));
  setVal("cfgLslRow",       toDisplay(cfg.lsl_row));
  setVal("cfgUpperTolRow",  toDisplay(cfg.upper_tol_row));
  setVal("cfgLowerTolRow",  toDisplay(cfg.lower_tol_row));
  setVal("cfgDataStartRow", toDisplay(cfg.data_start_row));

  // row_per_dim fields
  setVal("cfgNameCol",      toDisplay(cfg.name_col));
  setVal("cfgFirstDataRow", toDisplay(cfg.first_data_row));
  setVal("cfgNominalCol",   toDisplay(cfg.nominal_col));
  setVal("cfgUslCol",       toDisplay(cfg.usl_col));
  setVal("cfgLslCol",       toDisplay(cfg.lsl_col));
  setVal("cfgUpperTolCol",  toDisplay(cfg.upper_tol_col));
  setVal("cfgLowerTolCol",  toDisplay(cfg.lower_tol_col));
  setVal("cfgDataStartCol", toDisplay(cfg.data_start_col));

  onFormatChange();
  loadPreview(entry.fileId, entry.selectedSheet);
  configModal.show();
}

// Convert 0-based index to 1-based display (null -> "")
function toDisplay(v) { return v === null || v === undefined ? "" : v + 1; }
// Convert 1-based display back to 0-based (empty -> null)
function fromDisplay(v) {
  const n = parseInt(v, 10);
  return isNaN(n) ? null : n - 1;
}

function setVal(id, v) {
  const el = document.getElementById(id);
  if (el) el.value = v === null || v === undefined ? "" : v;
}

function onFormatChange() {
  const fmt = document.querySelector('input[name="cfgFormat"]:checked').value;
  document.getElementById("cfgColPerDim").classList.toggle("d-none", fmt !== "col_per_dim");
  document.getElementById("cfgRowPerDim").classList.toggle("d-none", fmt !== "row_per_dim");
}

async function loadPreview(fileId, sheetName) {
  const resp = await fetch("/api/preview", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ file_id: fileId, sheet_name: sheetName }),
  });
  const data = await resp.json();
  if (data.error) { console.error(data.error); return; }

  const tbody = document.getElementById("previewBody");
  tbody.innerHTML = "";

  data.rows.forEach((row, ri) => {
    const tr = document.createElement("tr");
    // Row number cell
    const th = document.createElement("th");
    th.textContent = ri + 1;
    th.style.background = "#f8f9fa";
    th.style.fontSize = "0.72rem";
    th.style.minWidth = "28px";
    tr.appendChild(th);

    row.forEach(cell => {
      const td = document.createElement("td");
      td.textContent = cell === null ? "" : cell;
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

async function applyConfig() {
  const idx  = state.configTarget;
  const entry = state.files[idx];
  const fmt  = document.querySelector('input[name="cfgFormat"]:checked').value;

  // Build config object (convert 1-based UI values → 0-based)
  const cfg = { format: fmt };
  if (fmt === "col_per_dim") {
    cfg.header_row     = fromDisplay(document.getElementById("cfgHeaderRow").value);
    cfg.skip_cols      = parseInt(document.getElementById("cfgSkipCols").value, 10) || 0;
    cfg.nominal_row    = fromDisplay(document.getElementById("cfgNominalRow").value);
    cfg.usl_row        = fromDisplay(document.getElementById("cfgUslRow").value);
    cfg.lsl_row        = fromDisplay(document.getElementById("cfgLslRow").value);
    cfg.upper_tol_row  = fromDisplay(document.getElementById("cfgUpperTolRow").value);
    cfg.lower_tol_row  = fromDisplay(document.getElementById("cfgLowerTolRow").value);
    cfg.data_start_row = fromDisplay(document.getElementById("cfgDataStartRow").value) ?? 1;
  } else {
    cfg.name_col       = fromDisplay(document.getElementById("cfgNameCol").value) ?? 0;
    cfg.first_data_row = fromDisplay(document.getElementById("cfgFirstDataRow").value) ?? 1;
    cfg.nominal_col    = fromDisplay(document.getElementById("cfgNominalCol").value);
    cfg.usl_col        = fromDisplay(document.getElementById("cfgUslCol").value);
    cfg.lsl_col        = fromDisplay(document.getElementById("cfgLslCol").value);
    cfg.upper_tol_col  = fromDisplay(document.getElementById("cfgUpperTolCol").value);
    cfg.lower_tol_col  = fromDisplay(document.getElementById("cfgLowerTolCol").value);
    cfg.data_start_col = fromDisplay(document.getElementById("cfgDataStartCol").value) ?? 4;
  }

  entry.config = cfg;

  // Fetch dimension names
  const resp = await fetch("/api/dimensions", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ file_id: entry.fileId, sheet_name: entry.selectedSheet, config: cfg }),
  });
  const data = await resp.json();
  if (data.error) { alert("提取尺寸出错：" + data.error); return; }

  entry.dimensions = data.dimensions;
  configModal.hide();
  renderFileList();
  mergeDimensions();
}

// ---------------------------------------------------------------------------
// Merge dimensions from all files and show checkboxes
// ---------------------------------------------------------------------------
function mergeDimensions() {
  const allDims = new Set();
  state.files.forEach(f => f.dimensions.forEach(d => allDims.add(d)));

  const card    = document.getElementById("dimCard");
  const optCard = document.getElementById("optCard");
  const list    = document.getElementById("dimList");

  if (allDims.size === 0) {
    card.style.setProperty("display", "none", "important");
    optCard.style.setProperty("display", "none", "important");
    return;
  }

  card.style.removeProperty("display");
  optCard.style.removeProperty("display");
  list.innerHTML = "";

  allDims.forEach(dim => {
    const id = "dim_" + CSS.escape(dim);
    const div = document.createElement("div");
    div.className = "form-check";
    div.innerHTML = `
      <input class="form-check-input dim-checkbox" type="checkbox" id="${id}"
             value="${esc(dim)}" checked />
      <label class="form-check-label dim-check-label" for="${id}">${esc(dim)}</label>
    `;
    list.appendChild(div);
  });

  // Add "Edit display info" button
  const btn = document.createElement("button");
  btn.className = "btn btn-sm btn-outline-secondary w-100 mt-2";
  btn.innerHTML = '<i class="bi bi-pencil me-1"></i>编辑显示信息';
  btn.onclick = openDimEditor;
  list.appendChild(btn);
}

function getSelectedDims() {
  return Array.from(document.querySelectorAll(".dim-checkbox:checked")).map(el => el.value);
}

function selectAllDims() {
  document.querySelectorAll(".dim-checkbox").forEach(el => { el.checked = true; });
}
function clearAllDims() {
  document.querySelectorAll(".dim-checkbox").forEach(el => { el.checked = false; });
}

// ---------------------------------------------------------------------------
// Dimension editor modal
// ---------------------------------------------------------------------------
function openDimEditor() {
  const allDims = new Set();
  state.files.forEach(f => f.dimensions.forEach(d => allDims.add(d)));

  const tbody = document.getElementById("dimEditorBody");
  tbody.innerHTML = "";

  allDims.forEach(dim => {
    const ov = state.dimOverrides[dim] || {};
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td class="align-middle fw-semibold">${esc(dim)}</td>
      <td><input type="text" class="form-control form-control-sm ov-display"
                 data-dim="${esc(dim)}" value="${esc(ov.displayName || "")}"
                 placeholder="${esc(dim)}" /></td>
      <td><input type="number" class="form-control form-control-sm ov-nominal"
                 data-dim="${esc(dim)}" value="${ov.nominal ?? ""}" placeholder="（来自数据）" step="any" /></td>
      <td><input type="number" class="form-control form-control-sm ov-usl"
                 data-dim="${esc(dim)}" value="${ov.usl ?? ""}" placeholder="（来自数据）" step="any" /></td>
      <td><input type="number" class="form-control form-control-sm ov-lsl"
                 data-dim="${esc(dim)}" value="${ov.lsl ?? ""}" placeholder="（来自数据）" step="any" /></td>
    `;
    tbody.appendChild(tr);
  });

  dimEditorModal.show();
}

function applyDimEdits() {
  document.querySelectorAll("#dimEditorBody tr").forEach(tr => {
    const dim = tr.querySelector(".ov-display").dataset.dim;
    const displayName = tr.querySelector(".ov-display").value.trim();
    const nominal     = parseFloatOrNull(tr.querySelector(".ov-nominal").value);
    const usl         = parseFloatOrNull(tr.querySelector(".ov-usl").value);
    const lsl         = parseFloatOrNull(tr.querySelector(".ov-lsl").value);
    state.dimOverrides[dim] = { displayName, nominal, usl, lsl };
  });
}

// ---------------------------------------------------------------------------
// Generate plot
// ---------------------------------------------------------------------------
async function generatePlot() {
  const selectedDims = getSelectedDims();
  if (selectedDims.length === 0) { alert("请先选择至少一个尺寸！"); return; }
  if (state.files.length === 0)  { alert("请先上传文件！"); return; }

  const filesReady = state.files.filter(f => f.dimensions.length > 0 || selectedDims.length > 0);
  if (filesReady.every(f => f.dimensions.length === 0)) {
    alert("请先在每个文件中配置数据格式并提取尺寸！");
    return;
  }

  const payload = {
    files: state.files.map(f => ({
      file_id:    f.fileId,
      stage_name: f.stageName,
      sheet_name: f.selectedSheet,
      config:     f.config,
    })),
    selected_dims: selectedDims,
  };

  const resp = await fetch("/api/plot_data", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  const result = await resp.json();
  if (result.error) { alert("获取数据出错：" + result.error); return; }

  state.plotData = result.data;
  renderPlot(result.data, selectedDims);
}

// ---------------------------------------------------------------------------
// Plotly rendering
// ---------------------------------------------------------------------------
const STAGE_COLORS = [
  "#1f77b4", "#ff7f0e", "#2ca02c", "#d62728",
  "#9467bd", "#8c564b", "#e377c2", "#17becf",
];

function renderPlot(data, selectedDims) {
  const normalize  = document.getElementById("chkNormalize").checked;
  const showSpec   = document.getElementById("chkShowSpec").checked;
  const showPoints = document.getElementById("chkShowPoints").checked;
  const title      = document.getElementById("plotTitle").value || "Box Plot";
  const yTitle     = document.getElementById("yAxisTitle").value || "测量值";

  const traces = [];
  const shapes = [];
  const annotations = [];

  // Collect spec info for each dim (use first stage's values as reference)
  const specInfo = {}; // {dimName: {nominal, usl, lsl}}
  data.forEach(stage => {
    stage.dimensions.forEach(dim => {
      if (!specInfo[dim.name]) {
        const ov = state.dimOverrides[dim.name] || {};
        specInfo[dim.name] = {
          nominal: ov.nominal ?? dim.nominal,
          usl:     ov.usl     ?? dim.usl,
          lsl:     ov.lsl     ?? dim.lsl,
        };
      }
    });
  });

  // For each stage, build a box trace
  data.forEach((stage, si) => {
    const color = STAGE_COLORS[si % STAGE_COLORS.length];
    const xVals = [];
    const yVals = [];

    stage.dimensions.forEach(dim => {
      if (!selectedDims.includes(dim.name)) return;
      const ov = state.dimOverrides[dim.name] || {};
      const displayName = (ov.displayName && ov.displayName.trim()) ? ov.displayName : dim.name;
      const nominal     = ov.nominal ?? dim.nominal;

      let meas = dim.measurements || [];
      if (normalize && nominal != null) {
        meas = meas.map(v => v - nominal);
      }

      meas.forEach(v => {
        xVals.push(displayName);
        yVals.push(v);
      });
    });

    if (xVals.length === 0) return;

    traces.push({
      type:       "box",
      name:       stage.stage,
      x:          xVals,
      y:          yVals,
      boxpoints:  showPoints ? "all" : "outliers",
      jitter:     0.35,
      pointpos:   -1.6,
      marker:     { color, opacity: 0.6, size: 5 },
      line:       { color },
      fillcolor:  hexToRgba(color, 0.18),
      whiskerwidth: 0.8,
    });
  });

  // Spec limit markers (one scatter trace per limit type)
  if (showSpec) {
    const nomX = [], nomY = [], uslX = [], uslY = [], lslX = [], lslY = [];

    selectedDims.forEach(dimName => {
      const spec = specInfo[dimName] || {};
      const ov   = state.dimOverrides[dimName] || {};
      const displayName = (ov.displayName && ov.displayName.trim()) ? ov.displayName : dimName;
      const nominal = spec.nominal;
      let usl = spec.usl;
      let lsl = spec.lsl;

      if (normalize && nominal != null) {
        if (usl != null) usl = usl - nominal;
        if (lsl != null) lsl = lsl - nominal;
        if (nominal != null) { nomX.push(displayName); nomY.push(0); }
      } else {
        if (nominal != null) { nomX.push(displayName); nomY.push(nominal); }
      }
      if (usl != null) { uslX.push(displayName); uslY.push(usl); }
      if (lsl != null) { lslX.push(displayName); lslY.push(lsl); }
    });

    if (uslX.length > 0) {
      traces.push({
        type: "scatter", mode: "markers",
        name: "USL", x: uslX, y: uslY,
        marker: { symbol: "line-ew-open", size: 22, color: "red", line: { color: "red", width: 2.5 } },
        showlegend: true,
      });
    }
    if (lslX.length > 0) {
      traces.push({
        type: "scatter", mode: "markers",
        name: "LSL", x: lslX, y: lslY,
        marker: { symbol: "line-ew-open", size: 22, color: "#28a745", line: { color: "#28a745", width: 2.5 } },
        showlegend: true,
      });
    }
    if (nomX.length > 0) {
      traces.push({
        type: "scatter", mode: "markers",
        name: "理论值", x: nomX, y: nomY,
        marker: { symbol: "line-ew-open", size: 22, color: "#fd7e14", line: { color: "#fd7e14", width: 2.5 } },
        showlegend: true,
      });
    }
  }

  const layout = {
    title:    { text: title, font: { size: 18 } },
    boxmode:  "group",
    yaxis:    { title: normalize ? "偏差值" : yTitle, zeroline: normalize },
    xaxis:    { title: "尺寸", type: "category" },
    legend:   { orientation: "h", y: -0.15 },
    margin:   { t: 60, l: 60, r: 20, b: 100 },
    paper_bgcolor: "#fff",
    plot_bgcolor:  "#fafafa",
    shapes,
    annotations,
  };

  const config = {
    responsive:  true,
    displaylogo: false,
    toImageButtonOptions: { format: "png", filename: "boxplot", scale: 2 },
    modeBarButtonsToAdd: [],
  };

  document.getElementById("plotWrapper").classList.remove("d-none");
  document.getElementById("emptyState").classList.add("d-none");
  Plotly.newPlot("plotDiv", traces, layout, config);

  renderStatsTable(data, selectedDims, normalize);
}

// ---------------------------------------------------------------------------
// Stats table
// ---------------------------------------------------------------------------
function renderStatsTable(data, selectedDims, normalize) {
  const tbody = document.getElementById("statsBody");
  tbody.innerHTML = "";

  data.forEach(stage => {
    stage.dimensions.forEach(dim => {
      if (!selectedDims.includes(dim.name)) return;
      const ov      = state.dimOverrides[dim.name] || {};
      const nominal = ov.nominal ?? dim.nominal;
      const usl     = ov.usl     ?? dim.usl;
      const lsl     = ov.lsl     ?? dim.lsl;
      const meas    = dim.measurements || [];

      if (meas.length === 0) return;

      const min  = Math.min(...meas);
      const max  = Math.max(...meas);
      const mean = meas.reduce((a, b) => a + b, 0) / meas.length;
      const std  = Math.sqrt(meas.reduce((a, b) => a + (b - mean) ** 2, 0) / meas.length);

      let cpk = "—";
      if (usl != null && lsl != null && std > 0) {
        const cpu = (usl - mean) / (3 * std);
        const cpl = (mean - lsl) / (3 * std);
        cpk = Math.min(cpu, cpl).toFixed(3);
      }

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${esc(stage.stage)}</td>
        <td>${esc(dim.name)}</td>
        <td>${fmt4(nominal)}</td>
        <td>${fmt4(usl)}</td>
        <td>${fmt4(lsl)}</td>
        <td>${meas.length}</td>
        <td>${fmt4(min)}</td>
        <td>${fmt4(max)}</td>
        <td>${fmt4(mean)}</td>
        <td>${fmt4(std)}</td>
        <td class="${cpk !== '—' && parseFloat(cpk) < 1.33 ? 'text-danger fw-bold' : ''}">${cpk}</td>
      `;
      tbody.appendChild(tr);
    });
  });

  document.getElementById("statsWrapper").classList.remove("d-none");
}

// ---------------------------------------------------------------------------
// Download
// ---------------------------------------------------------------------------
function downloadPNG() {
  Plotly.downloadImage("plotDiv", { format: "png", filename: "boxplot", width: 1400, height: 800, scale: 2 });
}

function downloadHTML() {
  const plotDiv = document.getElementById("plotDiv");
  // Capture current data and layout from Plotly
  const data   = plotDiv.data;
  const layout = plotDiv.layout;
  const dataJson   = JSON.stringify(data);
  const layoutJson = JSON.stringify(layout);
  const html = `<!DOCTYPE html>
<html><head><meta charset="UTF-8"/>
<title>Box Plot</title>
<script src="/static/vendor/plotly.min.js"><\/script>
</head>
<body style="margin:0;padding:8px">
<div id="plt"></div>
<script>
var d=${dataJson};
var l=${layoutJson};
Plotly.newPlot("plt",d,l,{responsive:true,displaylogo:false});
<\/script>
</body></html>`;
  const blob = new Blob([html], { type: "text/html" });
  const a    = document.createElement("a");
  a.href     = URL.createObjectURL(blob);
  a.download = "boxplot.html";
  a.click();
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------
function esc(str) {
  return String(str ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function fmt4(v) {
  if (v === null || v === undefined) return "—";
  return Number(v).toFixed(4);
}

function parseFloatOrNull(s) {
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function hexToRgba(hex, alpha) {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

function showEmptyState(show) {
  document.getElementById("emptyState").classList.toggle("d-none", !show);
  if (show) {
    document.getElementById("plotWrapper").classList.add("d-none");
    document.getElementById("statsWrapper").classList.add("d-none");
  }
}
