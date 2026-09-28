// AI-Assisted UCS Synthetic Data Generator Controller
// 100% faithful to UCS_Multiple_Depth_Generator_EXCEL_IMPORT.py

const SOIL_TYPES = [
  "Soft Clay",
  "Medium Clay",
  "Stiff Clay",
  "Very Stiff Clay",
  "Silty Clay",
  "Sandy Clay"
];

let depthRows = [];
let generatedData = {}; // keyed by row index
let selectedRowIndex = 0;
let previewTimeout = null;

// Initial Load & Bindings
document.addEventListener('DOMContentLoaded', () => {
  loadSavedState();
  if (depthRows.length === 0) {
    addDepthRow({ depth: "1.50", diameter: "3.792", height: "7.611", density: "1.84", ucs: "1.733", ring_constant: "0.325", soil: "Medium Clay", sample_type: "UDS", seed: 1001 });
    addDepthRow({ depth: "3.00", diameter: "3.792", height: "7.611", density: "1.84", ucs: "1.733", ring_constant: "0.325", soil: "Medium Clay", sample_type: "UDS", seed: 1002 });
  } else {
    renderDepthTable();
  }

  bindEvents();
  updateStatus("Ready. Edit depths and use Preview Selected to inspect realistic stress-strain curve.");
});

function bindEvents() {
  // Depth table controls
  document.getElementById('btnAddDepth').addEventListener('click', () => addDepthRow());
  document.getElementById('btnRemoveSelected').addEventListener('click', removeSelectedRow);
  document.getElementById('btnClearAll').addEventListener('click', clearAllRows);
  document.getElementById('btnAdd5Depths').addEventListener('click', add5DepthRows);

  // Excel import
  const fileInput = document.getElementById('ucsFileInput');
  document.getElementById('btnImportExcel').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleExcelImport);

  // Template browser
  const templateInput = document.getElementById('ucsTemplateInput');
  document.getElementById('btnBrowseTemplate').addEventListener('click', () => templateInput.click());
  templateInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files[0]) {
      document.getElementById('templatePath').value = e.target.files[0].name;
      alert(`Selected template: ${e.target.files[0].name}`);
    }
  });

  // Action Bar Buttons
  document.getElementById('btnGenerateAll').addEventListener('click', generateAllDepths);
  document.getElementById('btnGenerateSelected').addEventListener('click', generateSelectedDepth);
  document.getElementById('btnPreviewSelected').addEventListener('click', previewSelectedDepth);
  document.getElementById('btnPopulateSelected').addEventListener('click', populateSelectedXlsm);
  document.getElementById('btnPopulateAll').addEventListener('click', populateAllXlsm);
  document.getElementById('btnExportPlot').addEventListener('click', exportSelectedPlot);
  document.getElementById('btnExportAllPlots').addEventListener('click', exportAllPlots);

  // Modal Controls
  document.getElementById('btnModalClose').addEventListener('click', () => {
    document.getElementById('reportPlotModal').style.display = 'none';
  });
  document.getElementById('btnModalViewReport').addEventListener('click', viewModalReport);
  document.getElementById('btnModalSavePng').addEventListener('click', saveModalReportPng);
}

// -------------------------------------------------------------
// State Persistence
// -------------------------------------------------------------
function saveState() {
  try {
    const bh = document.getElementById('bh_id').value.trim();
    localStorage.setItem('ucs_bh_id', bh);
    localStorage.setItem('ucs_depth_rows', JSON.stringify(depthRows));
  } catch (e) {
    console.warn("Storage error", e);
  }
}

function loadSavedState() {
  try {
    const savedBh = localStorage.getItem('ucs_bh_id');
    if (savedBh) {
      document.getElementById('bh_id').value = savedBh;
    }
    const savedRows = localStorage.getItem('ucs_depth_rows');
    if (savedRows) {
      depthRows = JSON.parse(savedRows);
      renderDepthTable();
    }
  } catch (e) {
    console.warn("Storage load error", e);
  }
}

function updateStatus(msg) {
  const el = document.getElementById('statusVar');
  if (el) el.textContent = msg;
}

// -------------------------------------------------------------
// Row Management
// -------------------------------------------------------------
function addDepthRow(customVals = null) {
  const n = depthRows.length + 1;
  const row = {
    id: 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5),
    depth: customVals ? customVals.depth : (n * 1.50).toFixed(2),
    diameter: customVals ? customVals.diameter : "3.792",
    height: customVals ? customVals.height : "7.611",
    density: customVals ? customVals.density : "1.84",
    ucs: customVals ? customVals.ucs : "1.733",
    ring_constant: customVals ? customVals.ring_constant : "0.325",
    soil: customVals ? customVals.soil : "Medium Clay",
    sample_type: customVals ? customVals.sample_type : "UDS",
    seed: customVals ? customVals.seed : (1000 + n),
    status: customVals && customVals.status ? customVals.status : "Not generated"
  };

  depthRows.push(row);
  renderDepthTable();
  saveState();
}

function add5DepthRows() {
  const currentCount = depthRows.length;
  for (let i = 1; i <= 5; i++) {
    const depthVal = ((currentCount + i) * 1.50).toFixed(2);
    depthRows.push({
      id: 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5) + '_' + i,
      depth: depthVal,
      diameter: "3.792",
      height: "7.611",
      density: "1.84",
      ucs: "1.733",
      ring_constant: "0.325",
      soil: "Medium Clay",
      sample_type: "UDS",
      seed: 1000 + currentCount + i,
      status: "Not generated"
    });
  }
  renderDepthTable();
  saveState();
}

function removeSelectedRow() {
  if (depthRows.length <= 1) {
    alert("UCS\n\nAt least one depth row must remain.");
    return;
  }
  depthRows.splice(selectedRowIndex, 1);
  selectedRowIndex = Math.max(0, selectedRowIndex - 1);
  delete generatedData[selectedRowIndex];
  renderDepthTable();
  saveState();
  if (depthRows.length > 0) {
    selectRow(selectedRowIndex);
  }
}

function clearAllRows() {
  depthRows = [];
  generatedData = {};
  selectedRowIndex = 0;
  addDepthRow();
  renderObsTable([]);
  clearCanvas();
  updateStatus("All depth rows cleared.");
}

function renderDepthTable() {
  const tbody = document.getElementById('ucsTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  depthRows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.id = r.id;
    if (idx === selectedRowIndex) {
      tr.style.backgroundColor = "#e0f2fe";
    }

    let soilOpts = '';
    SOIL_TYPES.forEach(s => {
      soilOpts += `<option value="${s}" ${s === r.soil ? 'selected' : ''}>${s}</option>`;
    });

    tr.innerHTML = `
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 65px;" value="${r.depth}" onchange="updateRowField('${r.id}', 'depth', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 80px;" value="${r.diameter}" onchange="updateRowField('${r.id}', 'diameter', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 80px;" value="${r.height}" onchange="updateRowField('${r.id}', 'height', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 85px;" value="${r.density}" onchange="updateRowField('${r.id}', 'density', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 110px; font-weight: bold;" value="${r.ucs}" onchange="updateRowField('${r.id}', 'ucs', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 100px;" value="${r.ring_constant}" onchange="updateRowField('${r.id}', 'ring_constant', this.value)"></td>
      <td style="text-align: center;">
        <select class="tk-table-select" style="width: 120px;" onchange="updateRowField('${r.id}', 'soil', this.value)">
          ${soilOpts}
        </select>
      </td>
      <td style="text-align: center;">
        <select class="tk-table-select" style="width: 80px;" onchange="updateRowField('${r.id}', 'sample_type', this.value)">
          <option value="UDS" ${r.sample_type === 'UDS' ? 'selected' : ''}>UDS</option>
          <option value="SPT" ${r.sample_type === 'SPT' ? 'selected' : ''}>SPT</option>
        </select>
      </td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.seed}" onchange="updateRowField('${r.id}', 'seed', this.value)"></td>
      <td style="text-align: center; font-size: 8pt; color: ${r.status.startsWith('Generated') ? '#16a34a' : '#555555'}; font-weight: 500;">${r.status}</td>
      <td style="text-align: center;">
        <input type="radio" name="selectedDepthRow" ${idx === selectedRowIndex ? 'checked' : ''} onchange="selectRow(${idx})">
      </td>
    `;

    tr.addEventListener('click', (e) => {
      if (e.target.tagName !== 'INPUT' && e.target.tagName !== 'SELECT') {
        selectRow(idx);
      }
    });

    tbody.appendChild(tr);
  });
}

function updateRowField(id, field, val) {
  const r = depthRows.find(x => x.id === id);
  if (r) {
    r[field] = val;
    r.status = "Not generated";
    renderDepthTable();
    saveState();
  }
}

function selectRow(idx) {
  selectedRowIndex = idx;
  renderDepthTable();

  if (generatedData[idx]) {
    renderObsTable(generatedData[idx].points);
    drawStressStrainCurve(generatedData[idx]);
  } else {
    previewSelectedDepth();
  }
}

// -------------------------------------------------------------
// Live Stress-Strain Curve Drawing (Canvas Fallback + Backend Plot)
// -------------------------------------------------------------
function clearCanvas() {
  const canvas = document.getElementById('ucsCanvas');
  const img = document.getElementById('ucsPreviewImage');
  if (canvas) {
    const ctx = canvas.getContext('2d');
    ctx.clearRect(0, 0, canvas.width, canvas.height);
    ctx.fillStyle = "#888";
    ctx.font = "10pt 'Segoe UI'";
    ctx.textAlign = "center";
    ctx.fillText("Generate or preview a depth to display Stress–Strain Curve", canvas.width / 2, canvas.height / 2);
    canvas.style.display = 'block';
  }
  if (img) img.style.display = 'none';
}

function drawStressStrainCurve(dataItem) {
  const canvas = document.getElementById('ucsCanvas');
  const img = document.getElementById('ucsPreviewImage');
  if (!canvas || !dataItem || !dataItem.points) return;

  const ctx = canvas.getContext('2d');
  const w = canvas.width;
  const h = canvas.height;
  ctx.clearRect(0, 0, w, h);

  const specimen = dataItem.specimen || {};
  const points = dataItem.points;
  const strains = points.map(p => parseFloat(p.strain));
  const stresses = points.map(p => parseFloat(p.stress));

  const maxStrain = Math.max(...strains, 1.0);
  const maxStress = Math.max(...stresses, 0.1);
  const xmax = Math.max(20.0, maxStrain * 1.05);
  const ymax = Math.max(3.0, maxStress * 1.18);

  const padLeft = 85;
  const padBottom = 45;
  const padTop = 35;
  const padRight = 50;
  const pw = w - padLeft - padRight;
  const ph = h - padTop - padBottom;

  function xp(x) { return padLeft + (x / xmax) * pw; }
  function yp(y) { return (h - padBottom) - (y / ymax) * ph; }

  // Title
  ctx.fillStyle = "#123B8A";
  ctx.font = "bold 9.5pt 'Segoe UI'";
  ctx.textAlign = "center";
  const bh = specimen.bh_id || "BH-01";
  const depth = parseFloat(specimen.depth || 0.0).toFixed(2);
  const soil = specimen.soil || "Medium Clay";
  const sample = specimen.sample || specimen.sample_type || "UDS";
  ctx.fillText(`UCS Stress–Strain Curve | ${bh} | Depth ${depth} m | ${soil} | ${sample}`, padLeft + pw / 2, 22);

  // Grid
  ctx.strokeStyle = "#e2e8f0";
  ctx.lineWidth = 1;
  ctx.setLineDash([2, 2]);

  for (let x = 0; x <= xmax; x += 5.0) {
    ctx.beginPath();
    ctx.moveTo(xp(x), yp(0));
    ctx.lineTo(xp(x), yp(ymax));
    ctx.stroke();

    ctx.fillStyle = "#444";
    ctx.font = "8pt 'Segoe UI'";
    ctx.textAlign = "center";
    ctx.fillText(x.toFixed(0), xp(x), h - padBottom + 15);
  }

  const yStep = ymax > 5.0 ? 1.0 : 0.5;
  for (let y = 0; y <= ymax; y += yStep) {
    ctx.beginPath();
    ctx.moveTo(xp(0), yp(y));
    ctx.lineTo(xp(xmax), yp(y));
    ctx.stroke();

    ctx.fillStyle = "#444";
    ctx.font = "8pt 'Segoe UI'";
    ctx.textAlign = "right";
    ctx.fillText(y.toFixed(1), padLeft - 8, yp(y) + 3);
  }

  // Axes
  ctx.setLineDash([]);
  ctx.strokeStyle = "#000";
  ctx.lineWidth = 1.5;
  ctx.beginPath();
  ctx.moveTo(xp(0), yp(ymax));
  ctx.lineTo(xp(0), yp(0));
  ctx.lineTo(xp(xmax), yp(0));
  ctx.stroke();

  // Axis Labels
  ctx.fillStyle = "#123B8A";
  ctx.font = "bold 9pt 'Segoe UI'";
  ctx.textAlign = "center";
  ctx.fillText("Axial Strain, ε (%)", padLeft + pw / 2, h - 8);

  ctx.save();
  ctx.translate(22, padTop + ph / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText("Compressive Stress, σ₀ (kg/cm²)", 0, 0);
  ctx.restore();

  // Curve Line
  ctx.strokeStyle = "#1f77b4";
  ctx.lineWidth = 2.0;
  ctx.beginPath();
  points.forEach((p, i) => {
    const sx = xp(parseFloat(p.strain));
    const sy = yp(parseFloat(p.stress));
    if (i === 0) ctx.moveTo(sx, sy);
    else ctx.lineTo(sx, sy);
  });
  ctx.stroke();

  // Markers
  ctx.fillStyle = "#1f77b4";
  points.forEach(p => {
    ctx.beginPath();
    ctx.arc(xp(parseFloat(p.strain)), yp(parseFloat(p.stress)), 3, 0, 2 * Math.PI);
    ctx.fill();
  });

  // Peak Point Marker & Annotation
  const pIdx = dataItem.peak_index || 0;
  const peakStrain = parseFloat(points[pIdx].strain);
  const peakStress = parseFloat(points[pIdx].stress);

  ctx.fillStyle = "#dc2626";
  ctx.beginPath();
  ctx.arc(xp(peakStrain), yp(peakStress), 5.5, 0, 2 * Math.PI);
  ctx.fill();

  // Info Annotation Box
  ctx.fillStyle = "rgba(255, 255, 255, 0.94)";
  ctx.strokeStyle = "#adadad";
  ctx.lineWidth = 0.8;
  const bx = Math.min(w - padRight - 170, xp(peakStrain) + 12);
  const by = Math.max(padTop + 10, yp(peakStress) - 40);
  ctx.fillRect(bx, by, 165, 45);
  ctx.strokeRect(bx, by, 165, 45);

  ctx.fillStyle = "#dc2626";
  ctx.font = "bold 8.5pt 'Segoe UI'";
  ctx.textAlign = "left";
  ctx.fillText(`UCS (qu) = ${peakStress.toFixed(3)} kg/cm²`, bx + 8, by + 18);
  ctx.fillStyle = "#123B8A";
  ctx.fillText(`Peak Strain = ${peakStrain.toFixed(2)}%`, bx + 8, by + 34);

  // Asynchronously request server-side high-res preview plot for exact parity
  if (previewTimeout) clearTimeout(previewTimeout);
  previewTimeout = setTimeout(async () => {
    try {
      const res = await fetch('/api/ucs/preview_plot', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ specimen: specimen, result: dataItem.result })
      });
      const data = await res.json();
      if (data.success && data.image_b64 && img && canvas) {
        img.src = data.image_b64;
        img.onload = () => {
          img.style.display = 'block';
          canvas.style.display = 'none';
        };
      }
    } catch (e) {
      if (img && canvas) {
        img.style.display = 'none';
        canvas.style.display = 'block';
      }
    }
  }, 60);
}

// -------------------------------------------------------------
// 32 Observations Table Renderer
// -------------------------------------------------------------
function renderObsTable(points) {
  const tbody = document.getElementById('obsTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (!points || points.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align: center; color: #888; font-style: italic; padding: 40px;">
          Generate a depth to preview 32 observations.
        </td>
      </tr>
    `;
    return;
  }

  points.forEach(p => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">${p.sr}</td>
      <td style="text-align: center;">${p.dial}</td>
      <td style="text-align: center;">${p.cm}</td>
      <td style="text-align: center;">${p.strain}</td>
      <td style="text-align: center; font-weight: bold;">${p.pr}</td>
      <td style="text-align: center;">${p.load}</td>
      <td style="text-align: center; font-weight: bold; color: #1e3a8a;">${p.stress}</td>
    `;
    tbody.appendChild(tr);
  });
}

// -------------------------------------------------------------
// Generation Engine Calls
// -------------------------------------------------------------
async function generateAllDepths() {
  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  if (depthRows.length === 0) {
    alert("Generation\n\nAdd at least one depth row.");
    return;
  }

  const payload = {
    bh_id: bhId,
    rows: depthRows
  };

  updateStatus("Generating 32 PR observations and compressive stress curves...");

  try {
    const res = await fetch('/api/ucs/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Generation Error\n\n" + (data.error || "Failed to generate UCS data."));
      updateStatus("Error during generation.");
      return;
    }

    data.results.forEach((item, idx) => {
      generatedData[idx] = item;
      depthRows[idx].status = item.status;
    });

    renderDepthTable();
    saveState();

    if (data.results.length > 0) {
      selectRow(0);
    }

    updateStatus(`Generated ${data.count} depth(s) successfully.`);
    alert(`Generation Complete\n\nGenerated 32 observations for ${data.count} depth(s) successfully.`);
  } catch (e) {
    alert("Generation Error\n\n" + e.message);
    updateStatus("Error");
  }
}

async function generateSelectedDepth() {
  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  if (depthRows.length === 0) return;

  const targetRow = depthRows[selectedRowIndex];
  const payload = {
    bh_id: bhId,
    rows: [targetRow]
  };

  try {
    const res = await fetch('/api/ucs/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Generation Error\n\n" + (data.error || "Failed to generate depth."));
      return;
    }

    const item = data.results[0];
    generatedData[selectedRowIndex] = item;
    targetRow.status = item.status;

    renderDepthTable();
    saveState();
    renderObsTable(item.points);
    drawStressStrainCurve(item);

    updateStatus(`Depth ${parseFloat(targetRow.depth).toFixed(2)} m | Soil: ${targetRow.soil} | Generated UCS: ${item.actual_ucs.toFixed(3)} kg/cm²`);
  } catch (e) {
    alert("Generation Error\n\n" + e.message);
  }
}

async function previewSelectedDepth() {
  if (generatedData[selectedRowIndex]) {
    renderObsTable(generatedData[selectedRowIndex].points);
    drawStressStrainCurve(generatedData[selectedRowIndex]);
  } else {
    await generateSelectedDepth();
  }
}

// -------------------------------------------------------------
// Excel Import
// -------------------------------------------------------------
async function handleExcelImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  updateStatus("Importing depths from Excel...");

  try {
    const res = await fetch('/api/ucs/import_excel', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Excel Import Error\n\n" + (data.error || "Failed to import Excel."));
      updateStatus("Import error.");
      return;
    }

    if (data.bh_id) {
      document.getElementById('bh_id').value = data.bh_id;
    }

    depthRows = data.rows.map((r, i) => ({
      ...r,
      id: 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5) + '_' + i
    }));

    generatedData = {};
    selectedRowIndex = 0;
    renderDepthTable();
    saveState();
    clearCanvas();
    renderObsTable([]);

    updateStatus(`Imported ${data.count} depth(s) from Excel.`);
    alert(`Excel Import Complete\n\n${data.count} depth(s) imported successfully.`);
  } catch (err) {
    alert("Excel Import Error\n\n" + err.message);
    updateStatus("Import error.");
  } finally {
    e.target.value = '';
  }
}

// -------------------------------------------------------------
// XLSM Populating
// -------------------------------------------------------------
async function populateSelectedXlsm() {
  if (!generatedData[selectedRowIndex]) {
    await generateSelectedDepth();
  }

  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  const resultsList = Object.keys(generatedData).map(k => generatedData[k]);

  if (resultsList.length === 0) {
    alert("XLSM Population\n\nGenerate selected depth first.");
    return;
  }

  updateStatus("Populating XLSM template for selected depth...");

  try {
    const res = await fetch('/api/ucs/populate_xlsm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, results: [generatedData[selectedRowIndex]] })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("XLSM Error\n\n" + (data.error || "Failed to populate XLSM template."));
      return;
    }

    updateStatus(`Populated XLSM: ${data.filename}`);
    triggerDownload(data.url, data.filename);
  } catch (e) {
    alert("XLSM Error\n\n" + e.message);
  }
}

async function populateAllXlsm() {
  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  const allResults = Object.keys(generatedData).map(k => generatedData[k]);

  if (allResults.length !== depthRows.length) {
    await generateAllDepths();
  }

  const updatedResults = Object.keys(generatedData).map(k => generatedData[k]);
  if (updatedResults.length === 0) return;

  updateStatus("Populating XLSM workbooks for all depths...");

  try {
    const res = await fetch('/api/ucs/populate_xlsm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, results: updatedResults })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("XLSM Error\n\n" + (data.error || "Failed to populate XLSM templates."));
      return;
    }

    updateStatus(`Created ${data.count} populated XLSM workbook(s) and ${data.plot_count || data.count} report plot(s).`);
    triggerDownload(data.url, data.filename);
    alert(`Export Complete\n\n${data.count} populated XLSM workbook(s) + ${data.plot_count || data.count} report plot(s) bundled together into:\n${data.filename}`);
  } catch (e) {
    alert("Export Error\n\n" + e.message);
  }
}

// -------------------------------------------------------------
// Plot Exports & Modal Viewer
// -------------------------------------------------------------
async function exportSelectedPlot() {
  if (!generatedData[selectedRowIndex]) {
    await generateSelectedDepth();
  }

  const dataItem = generatedData[selectedRowIndex];
  if (!dataItem) return;

  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";

  // Open modal and show spinner
  const modal = document.getElementById('reportPlotModal');
  const img = document.getElementById('modalReportImage');
  const loading = document.getElementById('modalLoading');
  const select = document.getElementById('modalDepthSelect');

  modal.style.display = 'flex';
  img.style.display = 'none';
  loading.style.display = 'block';

  // Populate dropdown
  select.innerHTML = '';
  depthRows.forEach((r, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = `Depth ${parseFloat(r.depth).toFixed(2)} m (${r.soil} - ${r.sample_type})`;
    if (i === selectedRowIndex) opt.selected = true;
    select.appendChild(opt);
  });

  try {
    const res = await fetch('/api/ucs/export_plot', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, specimen: dataItem.specimen, result: dataItem.result })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Plot Export Error\n\n" + (data.error || "Failed to export report plot."));
      modal.style.display = 'none';
      return;
    }

    img.src = data.url;
    img.onload = () => {
      loading.style.display = 'none';
      img.style.display = 'block';
    };
    updateStatus(`Report plot exported: ${data.filename}`);
  } catch (e) {
    alert("Plot Export Error\n\n" + e.message);
    modal.style.display = 'none';
  }
}

async function viewModalReport() {
  const select = document.getElementById('modalDepthSelect');
  const idx = parseInt(select.value);
  selectedRowIndex = idx;
  renderDepthTable();

  if (!generatedData[idx]) {
    await generateSelectedDepth();
  }

  const dataItem = generatedData[idx];
  if (!dataItem) return;

  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  const img = document.getElementById('modalReportImage');
  const loading = document.getElementById('modalLoading');

  img.style.display = 'none';
  loading.style.display = 'block';

  try {
    const res = await fetch('/api/ucs/export_plot', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, specimen: dataItem.specimen, result: dataItem.result })
    });
    const data = await res.json();
    if (data.url) {
      img.src = data.url;
      img.onload = () => {
        loading.style.display = 'none';
        img.style.display = 'block';
      };
    }
  } catch (e) {
    alert("Report Error\n\n" + e.message);
  }
}

function saveModalReportPng() {
  const img = document.getElementById('modalReportImage');
  if (!img || !img.src) {
    alert("No report image rendered yet.");
    return;
  }
  const select = document.getElementById('modalDepthSelect');
  const idx = parseInt(select.value);
  const row = depthRows[idx] || depthRows[0];
  const bh = document.getElementById('bh_id').value.trim() || "BH-01";
  const fn = `UCS_Report_${bh}_Depth_${parseFloat(row.depth).toFixed(2)}m_${row.sample_type}.png`;
  triggerDownload(img.src, fn);
}

async function exportAllPlots() {
  const bhId = document.getElementById('bh_id').value.trim() || "BH-01";
  const allResults = Object.keys(generatedData).map(k => generatedData[k]);

  if (allResults.length !== depthRows.length) {
    await generateAllDepths();
  }

  const updatedResults = Object.keys(generatedData).map(k => generatedData[k]);
  if (updatedResults.length === 0) return;

  updateStatus("Exporting A4 landscape report plots for all depths...");

  try {
    const res = await fetch('/api/ucs/export_all_plots', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, results: updatedResults })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Plot Export Error\n\n" + (data.error || "Failed to export report plots."));
      return;
    }

    updateStatus(`Exported ${data.count} report plot(s).`);
    triggerDownload(data.zip_url, data.filename);
    alert(`Plot Export Complete\n\nExported ${data.count} report plot(s) into:\n${data.filename}`);
  } catch (e) {
    alert("Plot Export Error\n\n" + e.message);
  }
}

function triggerDownload(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  if (filename) a.setAttribute('download', filename);
  a.target = '_blank';
  document.body.appendChild(a);
  a.click();
  a.remove();
}
