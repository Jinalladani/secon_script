// AI-Assisted UU Triaxial Shear Test Controller
// 100% faithful to AI_UU_Triaxial_Data_Generator_MultiDepth_EXCEL_IMPORT_ALL_DEPTHS.py

const SOILS = [
  "Soft Clay",
  "Medium Clay",
  "Stiff Clay",
  "Very Stiff Clay",
  "Silty Clayey Soil",
  "Silty Clayey Sandy Soil"
];

let depthRows = [];
let generatedResults = null;
let allGeneratedRows = [];
let previewItem = null;
let currentPlotFile = null;

const STORAGE_KEY_TRIAX_ROWS = "goma_triaxial_depth_rows";
const STORAGE_KEY_TRIAX_COMMON = "goma_triaxial_common_settings";

document.addEventListener('DOMContentLoaded', () => {
  loadSavedState();

  if (depthRows.length === 0) {
    addDepthRow({ depth: "5.00", diameter: "3.80", height: "7.64", density: "1.80", sample_type: "UDS", soil: "Medium Clay", c: "0.50", phi: "10.00", p1: "0.50", p2: "1.00", p3: "1.50" });
    addDepthRow({ depth: "12.00", diameter: "3.80", height: "7.64", density: "1.85", sample_type: "UDS", soil: "Stiff Clay", c: "0.80", phi: "8.00", p1: "1.00", p2: "2.00", p3: "3.00" });
  }

  renderTable();
  bindEvents();

  if (depthRows.length > 0) {
    previewItem = depthRows[0];
    drawMohrPreview(previewItem);
  }
});

function loadSavedState() {
  const savedCommon = localStorage.getItem(STORAGE_KEY_TRIAX_COMMON);
  if (savedCommon) {
    try {
      const c = JSON.parse(savedCommon);
      if (c.bh_id) document.getElementById('commonBhId').value = c.bh_id;
      if (c.datetime) document.getElementById('commonDateTime').value = c.datetime;
      if (c.interval) document.getElementById('commonInterval').value = c.interval;
      if (c.seed) document.getElementById('commonSeed').value = c.seed;
    } catch (e) {}
  }

  const savedRows = localStorage.getItem(STORAGE_KEY_TRIAX_ROWS);
  if (savedRows) {
    try {
      depthRows = JSON.parse(savedRows) || [];
    } catch (e) {
      depthRows = [];
    }
  }
}

function autoSaveState() {
  localStorage.setItem(STORAGE_KEY_TRIAX_ROWS, JSON.stringify(depthRows));
  const common = {
    bh_id: document.getElementById('commonBhId').value,
    datetime: document.getElementById('commonDateTime').value,
    interval: document.getElementById('commonInterval').value,
    seed: document.getElementById('commonSeed').value
  };
  localStorage.setItem(STORAGE_KEY_TRIAX_COMMON, JSON.stringify(common));
}

function bindEvents() {
  ['commonBhId', 'commonDateTime', 'commonInterval', 'commonSeed'].forEach(id => {
    document.getElementById(id).addEventListener('input', () => {
      generatedResults = null;
      allGeneratedRows = [];
      autoSaveState();
      if (previewItem) drawMohrPreview(previewItem);
    });
  });

  document.getElementById('btnAddRow').addEventListener('click', () => addDepthRow());
  document.getElementById('btnAdd5Rows').addEventListener('click', () => {
    for (let i = 0; i < 5; i++) addDepthRow();
  });
  document.getElementById('btnRemoveSelected').addEventListener('click', removeSelectedRows);
  document.getElementById('btnClearAll').addEventListener('click', clearAllRows);
  document.getElementById('btnPreviewSelected').addEventListener('click', previewSelected);
  document.getElementById('btnGenerateAll').addEventListener('click', generateAllData);
  document.getElementById('btnExportAllPlots').addEventListener('click', exportAllMohrPlots);
  document.getElementById('btnSaveCsv').addEventListener('click', saveCsv);
  document.getElementById('btnPopulateXlsm').addEventListener('click', populateXlsm);
  document.getElementById('btnExportAll').addEventListener('click', exportAllPackage);
  document.getElementById('btnFinalMohrPlot').addEventListener('click', openFinalMohrPlotModal);

  // Modal events
  document.getElementById('btnModalClose').addEventListener('click', () => {
    document.getElementById('finalPlotModal').style.display = 'none';
  });
  document.getElementById('btnModalPlotDepth').addEventListener('click', updateModalReportView);
  document.getElementById('btnModalSavePng').addEventListener('click', downloadCurrentModalPlot);

  const fileInput = document.getElementById('triaxFileInput');
  document.getElementById('btnImportExcel').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', handleExcelImport);
}

function addDepthRow(defaults = {}) {
  const rowId = 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
  const num = depthRows.length + 1;
  const dVal = defaults.depth !== undefined ? defaults.depth : (num * 3.00).toFixed(2);

  const row = {
    id: rowId,
    depth: dVal,
    diameter: defaults.diameter || "3.80",
    height: defaults.height || "7.64",
    density: defaults.density || "1.80",
    sample_type: defaults.sample_type || "UDS",
    soil: defaults.soil || "Medium Clay",
    c: defaults.c || "0.50",
    phi: defaults.phi || "10.00",
    p1: defaults.p1 || "0.50",
    p2: defaults.p2 || "1.00",
    p3: defaults.p3 || "1.50",
    selected: false
  };

  depthRows.push(row);
  renderTable();
  autoSaveState();
}

function renderTable() {
  const tbody = document.getElementById('triaxialTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (depthRows.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="13" style="text-align: center; color: #888; font-style: italic; padding: 15px;">
          No depth rows. Click "+ Add Depth" or "Import Excel Input".
        </td>
      </tr>
    `;
    return;
  }

  depthRows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.id = r.id;

    let soilOpts = '';
    SOILS.forEach(s => {
      soilOpts += `<option value="${s}" ${s === r.soil ? 'selected' : ''}>${s}</option>`;
    });

    tr.innerHTML = `
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 65px;" value="${r.depth}" onchange="updateField('${r.id}', 'depth', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.diameter}" onchange="updateField('${r.id}', 'diameter', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 70px;" value="${r.height}" onchange="updateField('${r.id}', 'height', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.density}" onchange="updateField('${r.id}', 'density', this.value)"></td>
      <td style="text-align: center;">
        <select class="tk-table-select" style="width: 80px;" onchange="updateField('${r.id}', 'sample_type', this.value)">
          <option value="UDS" ${r.sample_type === 'UDS' ? 'selected' : ''}>UDS</option>
          <option value="SPT" ${r.sample_type === 'SPT' ? 'selected' : ''}>SPT</option>
        </select>
      </td>
      <td style="text-align: center;">
        <select class="tk-table-select" style="width: 160px;" onchange="updateField('${r.id}', 'soil', this.value)">
          ${soilOpts}
        </select>
      </td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.c}" onchange="updateField('${r.id}', 'c', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 60px;" value="${r.phi}" onchange="updateField('${r.id}', 'phi', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.p1}" onchange="updateField('${r.id}', 'p1', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.p2}" onchange="updateField('${r.id}', 'p2', this.value)"></td>
      <td style="text-align: center;"><input type="text" class="tk-table-entry" style="width: 75px;" value="${r.p3}" onchange="updateField('${r.id}', 'p3', this.value)"></td>
      <td style="text-align: center;"><input type="checkbox" ${r.selected ? 'checked' : ''} onchange="updateField('${r.id}', 'selected', this.checked)"></td>
      <td style="text-align: center;"><button type="button" class="tk-btn" style="padding: 1px 6px; font-size: 8pt;" onclick="selectAndPreview(${idx})">Preview</button></td>
    `;
    tbody.appendChild(tr);
  });
}

function updateField(id, field, val) {
  const r = depthRows.find(x => x.id === id);
  if (r) {
    r[field] = val;
    generatedResults = null;
    allGeneratedRows = [];
    autoSaveState();
    if (previewItem && previewItem.id === id) {
      drawMohrPreview(r);
    }
  }
}

function selectAndPreview(idx) {
  const r = depthRows[idx];
  if (r) {
    previewItem = r;
    const trs = document.querySelectorAll('#triaxialTableBody tr');
    trs.forEach((tr, i) => {
      tr.style.backgroundColor = (i === idx) ? '#dbeafe' : '';
    });
    drawMohrPreview(r);
  }
}

function previewSelected() {
  const selIdx = depthRows.findIndex(r => r.selected);
  if (selIdx === -1) {
    if (depthRows.length === 0) return;
    alert("Preview\n\nTick the checkbox of a depth, then click Preview Selected.");
    return;
  }
  selectAndPreview(selIdx);
}

function removeSelectedRows() {
  const selected = depthRows.filter(r => r.selected);
  if (selected.length === 0) {
    alert("Remove Depth\n\nTick the checkbox at the right of the depth row(s) to remove.");
    return;
  }
  depthRows = depthRows.filter(r => !r.selected);
  renderTable();
  autoSaveState();
  if (depthRows.length > 0) {
    selectAndPreview(0);
  } else {
    clearCanvas();
  }
}

function clearAllRows() {
  if (depthRows.length === 0) return;
  if (!confirm("Clear All\n\nRemove all depth rows?")) return;
  depthRows = [];
  generatedResults = null;
  allGeneratedRows = [];
  previewItem = null;
  renderTable();
  autoSaveState();
  clearCanvas();
  renderObsTable([]);
}

function clearCanvas() {
  const canvas = document.getElementById('mohrCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  ctx.fillStyle = "#888";
  ctx.font = "10pt 'Segoe UI'";
  ctx.textAlign = "center";
  ctx.fillText("No depth selected", canvas.width / 2, canvas.height / 2);
}

// -------------------------------------------------------------
// Canvas & Matplotlib Mohr Preview (Exact Physics & Semicircles)
// -------------------------------------------------------------
let previewPlotTimeout = null;

function drawMohrPreview(row) {
  if (!row) return;

  // If results were already generated, immediately update observation table with this depth's observations
  if (generatedResults && generatedResults.length > 0) {
    const matched = generatedResults.find(dr => Math.abs(parseFloat(dr.depth) - parseFloat(row.depth)) < 0.001);
    if (matched && matched.rows) {
      renderObsTable(matched.rows, matched.depth);
    }
  }

  const canvas = document.getElementById('mohrCanvas');
  const img = document.getElementById('mohrPreviewImage');

  // 1. Immediate client-side fallback canvas with generous margins (no clipping!)
  if (canvas) {
    const ctx = canvas.getContext('2d');
    const w = canvas.width;
    const h = canvas.height;

    ctx.clearRect(0, 0, w, h);

    const depth = parseFloat(row.depth) || 0.0;
    const dia = parseFloat(row.diameter) || 3.80;
    const height = parseFloat(row.height) || 7.64;
    const density = parseFloat(row.density) || 1.80;
    const cVal = parseFloat(row.c) || 0.50;
    const phiVal = parseFloat(row.phi) || 10.0;
    const p1 = parseFloat(row.p1) || 0.50;
    const p2 = parseFloat(row.p2) || 1.00;
    const p3 = parseFloat(row.p3) || 1.50;
    const pressures = [p1, p2, p3];

    const rad = (phiVal * Math.PI) / 180.0;
    const sinP = Math.sin(rad);
    const cosP = Math.cos(rad);
    const N_phi = (1 + sinP) / (1 - sinP);

    // Semicircles
    const circles = pressures.map((s3, i) => {
      const s1_ideal = s3 * N_phi + (2 * cVal * cosP) / (1 - sinP);
      const varFactor = 1.0 + (i === 0 ? -0.04 : i === 1 ? 0.02 : 0.05);
      const s1 = s1_ideal * varFactor;
      return {
        test: i + 1,
        s3: s3,
        s1: s1,
        center: (s1 + s3) / 2.0,
        radius: (s1 - s3) / 2.0
      };
    });

    const sigma_n = circles.map(c => c.center);
    const tau = circles.map(c => c.radius);
    const xbar = sigma_n.reduce((a, b) => a + b, 0) / 3.0;
    const ybar = tau.reduce((a, b) => a + b, 0) / 3.0;
    const den = sigma_n.reduce((acc, x) => acc + Math.pow(x - xbar, 2), 0);
    const slope = den > 1e-6 ? sigma_n.reduce((acc, x, i) => acc + (x - xbar) * (tau[i] - ybar), 0) / den : Math.tan(rad);
    const intercept = Math.max(0.0, ybar - slope * xbar);
    const fittedPhi = Math.max(0.0, (Math.atan(slope) * 180.0) / Math.PI);
    const fittedCu = intercept;

    const maxS1 = Math.max(...circles.map(c => c.s1));
    const maxTau = Math.max(...circles.map(c => c.radius));
    const xmax = Math.max(2.5, maxS1 * 1.35);
    const ymax = Math.max(1.2, maxTau * 1.70);

    // Generous paddings so left y-axis label and right parameter box never clip
    const padLeft = 85;
    const padBottom = 48;
    const padTop = 38;
    const padRight = 50;
    const pw = w - padLeft - padRight;
    const ph = h - padTop - padBottom;

    function xp(x) { return padLeft + (x / xmax) * pw; }
    function yp(y) { return (h - padBottom) - (y / ymax) * ph; }

    // Title
    ctx.fillStyle = "#000000";
    ctx.font = "bold 9.5pt 'Segoe UI'";
    ctx.textAlign = "center";
    ctx.fillText(`Realistic Mohr Failure Envelope Preview | Depth = ${depth.toFixed(2)} m | ${row.soil}`, padLeft + pw / 2, 20);

    // Grid lines
    ctx.strokeStyle = "#e2e8f0";
    ctx.lineWidth = 1;
    ctx.setLineDash([2, 2]);

    const xStep = xmax > 4.0 ? 1.0 : 0.5;
    for (let x = 0; x <= xmax; x += xStep) {
      ctx.beginPath();
      ctx.moveTo(xp(x), yp(0));
      ctx.lineTo(xp(x), yp(ymax));
      ctx.stroke();

      ctx.fillStyle = "#444444";
      ctx.font = "8pt 'Segoe UI'";
      ctx.textAlign = "center";
      ctx.fillText(x.toFixed(1), xp(x), h - padBottom + 16);
    }

    for (let y = 0; y <= ymax; y += 0.5) {
      ctx.beginPath();
      ctx.moveTo(xp(0), yp(y));
      ctx.lineTo(xp(xmax), yp(y));
      ctx.stroke();

      ctx.fillStyle = "#444444";
      ctx.font = "8pt 'Segoe UI'";
      ctx.textAlign = "right";
      ctx.fillText(y.toFixed(1), padLeft - 8, yp(y) + 3);
    }

    // Axes
    ctx.setLineDash([]);
    ctx.strokeStyle = "#000000";
    ctx.lineWidth = 1.5;
    ctx.beginPath();
    ctx.moveTo(xp(0), yp(ymax));
    ctx.lineTo(xp(0), yp(0));
    ctx.lineTo(xp(xmax), yp(0));
    ctx.stroke();

    // Axis Labels
    ctx.fillStyle = "#000000";
    ctx.font = "bold 9pt 'Segoe UI'";
    ctx.textAlign = "center";
    ctx.fillText("Normal Stress, σ (kg/cm²)", padLeft + pw / 2, h - 8);

    ctx.save();
    ctx.translate(25, padTop + ph / 2);
    ctx.rotate(-Math.PI / 2);
    ctx.fillText("Shear Stress, τ (kg/cm²)", 0, 0);
    ctx.restore();

    // 3 Test Semicircles
    const testColors = ["#1f77b4", "#2ca02c", "#9467bd"];
    circles.forEach((c, i) => {
      ctx.strokeStyle = testColors[i];
      ctx.lineWidth = 2.0;

      const cx = xp(c.center);
      const cy = yp(0);
      const rx = (c.radius / xmax) * pw;
      const ry = (c.radius / ymax) * ph;

      ctx.beginPath();
      ctx.ellipse(cx, cy, rx, ry, 0, Math.PI, 0, false);
      ctx.stroke();

      // Marker
      ctx.fillStyle = testColors[i];
      ctx.beginPath();
      ctx.arc(xp(c.s1), yp(0), 4, 0, 2 * Math.PI);
      ctx.fill();

      // Peak point
      ctx.fillStyle = "#000000";
      ctx.beginPath();
      ctx.arc(cx, yp(c.radius), 4, 0, 2 * Math.PI);
      ctx.fill();

      // Perpendicular line
      ctx.strokeStyle = "#555555";
      ctx.lineWidth = 0.9;
      ctx.setLineDash([3, 3]);
      ctx.beginPath();
      ctx.moveTo(cx, yp(0));
      ctx.lineTo(cx, yp(c.radius));
      ctx.stroke();
      ctx.setLineDash([]);
    });

    // Failure Envelope Line
    ctx.strokeStyle = "#dc2626";
    ctx.lineWidth = 2.2;
    ctx.beginPath();
    ctx.moveTo(xp(0), yp(fittedCu));
    ctx.lineTo(xp(xmax), yp(fittedCu + slope * xmax));
    ctx.stroke();

    // Legend on Top Left
    ctx.fillStyle = "rgba(255, 255, 255, 0.92)";
    ctx.strokeStyle = "#adadad";
    ctx.lineWidth = 0.8;
    ctx.fillRect(padLeft + 10, padTop + 8, 165, 78);
    ctx.strokeRect(padLeft + 10, padTop + 8, 165, 78);

    circles.forEach((c, i) => {
      ctx.strokeStyle = testColors[i];
      ctx.lineWidth = 2;
      ctx.beginPath();
      ctx.moveTo(padLeft + 16, padTop + 22 + i * 16);
      ctx.lineTo(padLeft + 36, padTop + 22 + i * 16);
      ctx.stroke();

      ctx.fillStyle = "#000000";
      ctx.font = "8pt 'Segoe UI'";
      ctx.textAlign = "left";
      ctx.fillText(`Test ${i + 1} (σ₃ = ${c.s3.toFixed(2)})`, padLeft + 42, padTop + 25 + i * 16);
    });

    ctx.strokeStyle = "#dc2626";
    ctx.lineWidth = 2;
    ctx.beginPath();
    ctx.moveTo(padLeft + 16, padTop + 70);
    ctx.lineTo(padLeft + 36, padTop + 70);
    ctx.stroke();

    ctx.fillStyle = "#000000";
    ctx.font = "8pt 'Segoe UI'";
    ctx.fillText("Failure Envelope", padLeft + 42, padTop + 73);

    // Parameter Info Box (Top Right)
    ctx.fillStyle = "rgba(255, 255, 255, 0.93)";
    ctx.strokeStyle = "#adadad";
    ctx.lineWidth = 0.8;
    const boxW = 185;
    const boxH = 96;
    const bx = w - padRight - boxW - 5;
    const by = padTop + 8;
    ctx.fillRect(bx, by, boxW, boxH);
    ctx.strokeRect(bx, by, boxW, boxH);

    ctx.fillStyle = "#000000";
    ctx.font = "8pt 'Segoe UI'";
    ctx.textAlign = "left";
    ctx.fillText(`Preview fitted Cu = ${fittedCu.toFixed(2)} kg/cm²`, bx + 8, by + 15);
    ctx.fillText(`Preview fitted φ = ${fittedPhi.toFixed(2)}°`, bx + 8, by + 28);
    ctx.fillText(`Input Cu = ${cVal.toFixed(2)} kg/cm²`, bx + 8, by + 41);
    ctx.fillText(`Input φ = ${phiVal.toFixed(2)}°`, bx + 8, by + 54);
    ctx.fillText(`D = ${dia.toFixed(2)} cm`, bx + 8, by + 67);
    ctx.fillText(`H = ${height.toFixed(2)} cm`, bx + 8, by + 80);
    ctx.fillText(`Density = ${density.toFixed(2)} gm/cc`, bx + 8, by + 92);
  }

  // 2. Fetch server-side Matplotlib chart with exact tight layout
  if (previewPlotTimeout) clearTimeout(previewPlotTimeout);
  previewPlotTimeout = setTimeout(async () => {
    try {
      const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
      const startDt = document.getElementById('commonDateTime').value.trim() || "05-08-2026 15:47:00";
      const interval = parseFloat(document.getElementById('commonInterval').value) || 0.10;
      const seed = parseInt(document.getElementById('commonSeed').value) || 20260816;

      const payload = {
        bh_id: bhId,
        initial_datetime: startDt,
        interval_mm: interval,
        random_seed: seed,
        depth: row.depth,
        diameter: row.diameter,
        height: row.height,
        density: row.density,
        sample_type: row.sample_type,
        soil: row.soil,
        c: row.c,
        phi: row.phi,
        p1: row.p1,
        p2: row.p2,
        p3: row.p3
      };

      const res = await fetch('/api/triaxial/preview_plot', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload)
      });
      const data = await res.json();

      if (data.success) {
        if (data.image_b64 && img && canvas) {
          img.src = data.image_b64;
          img.onload = () => {
            img.style.display = 'block';
            canvas.style.display = 'none';
          };
        }
        if (data.status_text) {
          document.getElementById('statusVar').textContent = data.status_text;
        }
        // Always show all observations for this previewed depth
        if (data.rows && data.rows.length > 0) {
          renderObsTable(data.rows, row.depth);
        }
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
// Generate All Data Backend Execution
// -------------------------------------------------------------
async function generateAllData() {
  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  const startDt = document.getElementById('commonDateTime').value.trim() || "05-08-2026 15:47:00";
  const interval = parseFloat(document.getElementById('commonInterval').value) || 0.10;
  const seed = parseInt(document.getElementById('commonSeed').value) || 20260816;

  if (depthRows.length === 0) {
    alert("Generation\n\nAdd at least one depth row.");
    return;
  }

  const payload = {
    bh_id: bhId,
    initial_datetime: startDt,
    interval_mm: interval,
    random_seed: seed,
    rows: depthRows
  };

  document.getElementById('statusVar').textContent = "Generating continuous stress-strain curves and Mohr circles...";

  try {
    const res = await fetch('/api/triaxial/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Generation Error\n\n" + (data.error || "Failed to generate triaxial data."));
      document.getElementById('statusVar').textContent = "Error";
      return;
    }

    generatedResults = data.depth_results;
    allGeneratedRows = data.all_rows || [];

    const activeDepth = previewItem ? previewItem.depth : (depthRows.length > 0 ? depthRows[0].depth : null);
    let matched = null;
    if (generatedResults && generatedResults.length > 0) {
      matched = generatedResults.find(dr => Math.abs(parseFloat(dr.depth) - parseFloat(activeDepth)) < 0.001) || generatedResults[0];
    }

    if (matched && matched.rows) {
      renderObsTable(matched.rows, matched.depth);
    } else {
      renderObsTable(allGeneratedRows, activeDepth);
    }

    if (previewItem) {
      drawMohrPreview(previewItem);
    } else if (depthRows.length > 0) {
      selectAndPreview(0);
    }

    document.getElementById('statusVar').textContent =
      `Generated ${data.total_observations} observations | ${generatedResults.length} depth(s) | ${generatedResults.length * 3} UU tests.`;

    alert(`Generation Complete\n\nGenerated ${data.total_observations} continuous observations for ${generatedResults.length} depth(s).`);
  } catch (e) {
    alert("Generation Error\n\n" + e.message);
    document.getElementById('statusVar').textContent = "Error";
  }
}

function renderObsTable(rows, depth) {
  const tbody = document.getElementById('obsTableBody');
  const legend = document.getElementById('obsLegend');

  if (legend) {
    if (depth !== undefined && depth !== null) {
      legend.textContent = `Generated Observations - Depth ${parseFloat(depth).toFixed(2)} m (${rows ? rows.length : 0} Records)`;
    } else {
      legend.textContent = `Generated Observations (${rows ? rows.length : 0} Records)`;
    }
  }

  if (!tbody) return;
  tbody.innerHTML = '';

  if (!rows || rows.length === 0) {
    tbody.innerHTML = `
      <tr>
        <td colspan="7" style="text-align: center; color: #888; font-style: italic; padding: 30px;">
          No observations available. Click "Generate All Data" or "Preview".
        </td>
      </tr>
    `;
    return;
  }

  let html = '';
  rows.forEach(r => {
    html += `
      <tr>
        <td style="text-align: center;">${r["Seq#"]}</td>
        <td style="text-align: center;">${parseFloat(r["depth"]).toFixed(2)}</td>
        <td style="text-align: center;">${r["test"]}</td>
        <td style="text-align: center;">${r["sample_type"]}</td>
        <td style="text-align: center;">${r["DateTime"]}</td>
        <td style="text-align: center;">${parseFloat(r["STRAIN"]).toFixed(3)}</td>
        <td style="text-align: center; font-weight: bold;">${parseFloat(r["LOAD"]).toFixed(3)}</td>
      </tr>
    `;
  });
  tbody.innerHTML = html;
}

// -------------------------------------------------------------
// Final Mohr Plot Report View
// -------------------------------------------------------------
async function openFinalMohrPlotModal() {
  if (!generatedResults || generatedResults.length === 0) {
    await generateAllData();
    if (!generatedResults || generatedResults.length === 0) return;
  }

  const modal = document.getElementById('finalPlotModal');
  const select = document.getElementById('modalDepthSelect');
  select.innerHTML = '';

  generatedResults.forEach((dr, idx) => {
    const opt = document.createElement('option');
    opt.value = idx;
    opt.textContent = `${idx + 1}: Depth ${parseFloat(dr.depth).toFixed(2)} m | ${dr.sample_type} | ${dr.soil}`;
    select.appendChild(opt);
  });

  modal.style.display = 'flex';
  updateModalReportView();
}

async function updateModalReportView() {
  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  const select = document.getElementById('modalDepthSelect');
  const selIdx = parseInt(select.value) || 0;
  const dr = generatedResults[selIdx];

  if (!dr) return;

  document.getElementById('modalTitle').textContent =
    `Final UU Mohr's Circle Plot - Depth ${parseFloat(dr.depth).toFixed(2)} m`;
  document.getElementById('modalLoading').style.display = 'block';
  document.getElementById('modalReportImage').style.display = 'none';

  try {
    const res = await fetch('/api/triaxial/export_plots', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bh_id: bhId,
        depth_results: [dr]
      })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "Failed to render report plot.");
    }

    if (data.files && data.files.length > 0) {
      currentPlotFile = data.files[0];
      const img = document.getElementById('modalReportImage');
      img.src = currentPlotFile.url + '?t=' + Date.now();
      img.onload = () => {
        document.getElementById('modalLoading').style.display = 'none';
        img.style.display = 'block';
      };
    }
  } catch (e) {
    document.getElementById('modalLoading').innerHTML = `<span style="color: #dc2626;">Error: ${e.message}</span>`;
  }
}

function downloadCurrentModalPlot() {
  if (currentPlotFile && currentPlotFile.url) {
    triggerSafeDownload(currentPlotFile.url, currentPlotFile.filename);
  } else {
    alert("Please view a report plot first.");
  }
}

// -------------------------------------------------------------
// Export Handlers (ZIP Plots, CSV, XLSM)
// -------------------------------------------------------------
function triggerSafeDownload(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  if (filename) a.setAttribute('download', filename);
  a.target = '_blank';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function exportAllMohrPlots() {
  if (!generatedResults || generatedResults.length === 0) {
    alert("Export\n\nPlease click Generate All Data first.");
    return;
  }

  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  try {
    const res = await fetch('/api/triaxial/export_plots', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, depth_results: generatedResults })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "Plot export failed.");
    }

    if (data.download_zip_url) {
      triggerSafeDownload(data.download_zip_url, data.zip_filename);
    }

    alert(`Plot Export Complete\n\n${data.count} final Mohr plot(s) exported successfully into ZIP archive:\n${data.zip_filename}`);
  } catch (e) {
    alert("Plot Export Error\n\n" + e.message);
  }
}

async function saveCsv() {
  if (!allGeneratedRows || allGeneratedRows.length === 0) {
    alert("Save CSV\n\nPlease click Generate All Data first.");
    return;
  }

  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  try {
    const res = await fetch('/api/triaxial/export_csv', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ bh_id: bhId, all_rows: allGeneratedRows })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "CSV export failed.");
    }

    if (data.csv_url) {
      triggerSafeDownload(data.csv_url, data.filename);
    }

    alert(`CSV Export Complete\n\nCSV exported successfully with columns:\nSeq# | DateTime | STRAIN (mm) | LOAD (kg)`);
  } catch (e) {
    alert("CSV Export Error\n\n" + e.message);
  }
}

async function populateXlsm() {
  if (!generatedResults || generatedResults.length === 0) {
    alert("XLSM Template\n\nPlease click Generate All Data first.");
    return;
  }

  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  try {
    const res = await fetch('/api/triaxial/export_xlsm', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bh_id: bhId,
        depth_results: generatedResults,
        all_rows: allGeneratedRows
      })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "XLSM population failed.");
    }

    if (data.download_zip_url) {
      triggerSafeDownload(data.download_zip_url, data.zip_filename);
    }

    alert(`All Depths Completed\n\n${data.count} XLSM file(s) created successfully and packed into:\n${data.zip_filename}`);
  } catch (e) {
    alert("XLSM Error\n\n" + e.message);
  }
}

async function exportAllPackage() {
  if (!generatedResults || generatedResults.length === 0) {
    alert("All Export\n\nPlease click Generate All Data first.");
    return;
  }

  const bhId = document.getElementById('commonBhId').value.trim() || "BH-01";
  document.getElementById('statusVar').textContent = "Exporting all Mohr plots, CSV observations, and XLSM templates into a single ZIP archive...";

  try {
    const res = await fetch('/api/triaxial/export_all', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bh_id: bhId,
        depth_results: generatedResults,
        all_rows: allGeneratedRows
      })
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "All Export failed.");
    }

    if (data.download_zip_url) {
      triggerSafeDownload(data.download_zip_url, data.zip_filename);
    }

    document.getElementById('statusVar').textContent = `All Export complete: ${data.zip_filename}`;
    alert(`All Export Complete\n\nAll 3 components packaged into a single ZIP file:\n\n` +
          `• ${data.count_plots} Mohr Circle Plot(s) (.png)\n` +
          `• 1 Observations Data File (.csv)\n` +
          `• ${data.count_xlsm} Populated Excel Template(s) (.xlsm)\n\n` +
          `Archive: ${data.zip_filename}`);
  } catch (e) {
    document.getElementById('statusVar').textContent = "All Export Error";
    alert("All Export Error\n\n" + e.message);
  }
}

// -------------------------------------------------------------
// Excel Input Template Import
// -------------------------------------------------------------
async function handleExcelImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/triaxial/import_excel', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Excel Import Error\n\n" + (data.error || "Failed to parse Excel input."));
      return;
    }

    if (data.bh_id) document.getElementById('commonBhId').value = data.bh_id;
    if (data.initial_datetime) document.getElementById('commonDateTime').value = data.initial_datetime;
    if (data.interval_mm) document.getElementById('commonInterval').value = data.interval_mm;
    if (data.random_seed) document.getElementById('commonSeed').value = data.random_seed;

    depthRows = (data.rows || []).map((r, i) => ({
      id: 'row_' + Date.now() + '_' + i,
      depth: r.depth,
      diameter: r.diameter,
      height: r.height,
      density: r.density,
      sample_type: r.sample_type,
      soil: r.soil,
      c: r.c,
      phi: r.phi,
      p1: r.p1,
      p2: r.p2,
      p3: r.p3,
      selected: false
    }));

    renderTable();
    autoSaveState();

    if (depthRows.length > 0) {
      previewItem = depthRows[0];
      drawMohrPreview(previewItem);
    }

    alert(`Excel Import Complete\n\n${data.count} depth row(s) imported successfully.\nThe imported values are now loaded into the generator.`);
  } catch (err) {
    alert("Excel Import Error\n\n" + err.message);
  } finally {
    e.target.value = '';
  }
}
