// Direct Shear Test (DST) Desktop-Matched Controller
// 100% exact replica of DST_MultiDepth_Excel_Import.py logic, validations, and canvas report

let soilNames = [
  "Loose Sand",
  "Medium Dense Sand",
  "Dense Sand",
  "Very Dense Sand",
  "Soft Clay",
  "Medium Stiff Clay",
  "Stiff Clay",
  "Very Stiff Clay",
  "Hard Clay",
  "Silty Sand (SM)",
  "Clayey Sand (SC)",
  "Silt (ML/MI/MH)",
  "Gravelly Sand",
  "Sandy Gravel"
];

const SAMPLE_VOLUME_CM3 = 90.0;
let depthRows = [];
let results = [];
let selectedResult = null;

document.addEventListener('DOMContentLoaded', async () => {
  try {
    const res = await fetch('/api/dst/soil_names');
    const data = await res.json();
    if (data.soil_names && data.soil_names.length > 0) {
      soilNames = data.soil_names;
    }
  } catch (e) {
    console.warn("Using default soil profiles");
  }

  // Exact initial state from original desktop script (1 default row)
  addDepthRow();
  bindEvents();
  redraw();
});

function bindEvents() {
  document.getElementById('btnAddDepth').addEventListener('click', () => addDepthRow());
  document.getElementById('btnValidate').addEventListener('click', validateInputs);
  document.getElementById('btnGenerate').addEventListener('click', generateAll);
  document.getElementById('btnExportExcel').addEventListener('click', exportAllExcel);
  document.getElementById('btnExportPlots').addEventListener('click', exportPlots);
  document.getElementById('btnReset').addEventListener('click', resetForm);

  const fileInput = document.getElementById('excelFileInput');
  document.getElementById('btnImportExcel').addEventListener('click', () => fileInput.click());
  fileInput.addEventListener('change', importFromExcel);
}

function addDepthRow() {
  const rowId = 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);
  const row = {
    id: rowId,
    depth: "",
    soil: "Medium Dense Sand",
    sample_type: "Undisturbed Sample",
    density: "",
    weight: "",
    c: "",
    phi: "",
    behaviour: "Enter density"
  };
  depthRows.push(row);
  renderDepthTable();
}

function renderDepthTable() {
  const tbody = document.getElementById('depthTableBody');
  tbody.innerHTML = '';

  depthRows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.id = row.id;

    let soilOptions = '';
    soilNames.forEach(s => {
      soilOptions += `<option value="${s}" ${s === row.soil ? 'selected' : ''}>${s}</option>`;
    });

    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>
        <input type="text" class="tk-table-entry" value="${row.depth}" onchange="updateField('${row.id}', 'depth', this.value)">
      </td>
      <td>
        <select class="tk-table-select" onchange="updateSoil('${row.id}', this.value)">
          ${soilOptions}
        </select>
      </td>
      <td>
        <select class="tk-table-select" onchange="updateField('${row.id}', 'sample_type', this.value)">
          <option value="Undisturbed Sample" ${row.sample_type === 'Undisturbed Sample' ? 'selected' : ''}>Undisturbed Sample</option>
          <option value="Remolded" ${row.sample_type === 'Remolded' ? 'selected' : ''}>Remolded</option>
        </select>
      </td>
      <td>
        <input type="text" class="tk-table-entry" value="${row.density}" oninput="updateDensity('${row.id}', this.value)">
      </td>
      <td>
        <input type="text" class="tk-table-entry" value="${row.weight}" readonly style="background-color: #f0f0f0;">
      </td>
      <td>
        <input type="text" class="tk-table-entry" value="${row.c}" onchange="updateField('${row.id}', 'c', this.value)">
      </td>
      <td>
        <input type="text" class="tk-table-entry" value="${row.phi}" onchange="updateField('${row.id}', 'phi', this.value)">
      </td>
      <td>
        <span id="beh_${row.id}">${row.behaviour}</span>
      </td>
      <td>
        <button class="tk-btn-remove" onclick="removeDepthRow('${row.id}')">X</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function updateField(id, field, val) {
  const row = depthRows.find(r => r.id === id);
  if (row) row[field] = val;
}

async function updateSoil(id, soil) {
  const row = depthRows.find(r => r.id === id);
  if (row) {
    row.soil = soil;
    await checkBehaviour(row);
    const el = document.getElementById(`beh_${id}`);
    if (el) el.textContent = row.behaviour;
  }
}

async function updateDensity(id, densityStr) {
  const row = depthRows.find(r => r.id === id);
  if (row) {
    row.density = densityStr;
    const d = parseFloat(densityStr);
    const tr = document.getElementById(id);
    const weightInput = tr ? tr.querySelectorAll('input')[2] : null;

    if (!isNaN(d) && d > 0) {
      row.weight = (d * SAMPLE_VOLUME_CM3).toFixed(2);
      if (weightInput) weightInput.value = row.weight;
      await checkBehaviour(row);
    } else {
      row.weight = "";
      row.behaviour = "Enter density";
      if (weightInput) weightInput.value = "";
    }
    const el = document.getElementById(`beh_${id}`);
    if (el) el.textContent = row.behaviour;
  }
}

async function checkBehaviour(row) {
  try {
    const res = await fetch('/api/dst/classify_density', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ soil: row.soil, density: row.density })
    });
    const data = await res.json();
    row.behaviour = data.behaviour || "Enter density";
  } catch (e) {
    row.behaviour = "Enter density";
  }
}

function removeDepthRow(id) {
  if (depthRows.length <= 1) {
    alert("At least one depth row is required.");
    return;
  }
  depthRows = depthRows.filter(r => r.id !== id);
  renderDepthTable();
}

function getPayload() {
  return {
    bh_id: document.getElementById('bh_id').value.trim(),
    pr_constant: document.getElementById('pr_constant').value.trim(),
    normal_1: document.getElementById('normal_1').value.trim(),
    normal_2: document.getElementById('normal_2').value.trim(),
    normal_3: document.getElementById('normal_3').value.trim(),
    random_seed: document.getElementById('random_seed').value.trim(),
    rows: depthRows
  };
}

async function validateInputs() {
  const payload = getPayload();
  const statusEl = document.getElementById('statusVar');

  try {
    const res = await fetch('/api/dst/validate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json().catch(() => ({}));

    if (!res.ok || !data.valid) {
      const errMsg = (data.errors && data.errors.length > 0) ? data.errors.join("\n") : (data.error || "Validation failed.");
      alert("Input Error\n\n" + errMsg);
      statusEl.textContent = "Input validation failed.";
      return false;
    }

    if (data.warnings && data.warnings.length > 0) {
      const ok = confirm("Target Strength Warning\n\n" + data.warnings.join("\n") + "\n\nContinue with this target?");
      if (!ok) {
        statusEl.textContent = "Generation cancelled by user.";
        return false;
      }
    }

    alert(`Validation\n\nAll inputs are valid for ${data.count} depth(s).`);
    statusEl.textContent = `Valid input: ${data.count} depth(s).`;
    return true;
  } catch (e) {
    alert("Input Error\n\n" + e.message);
    return false;
  }
}

async function generateAll() {
  const payload = getPayload();
  const statusEl = document.getElementById('statusVar');

  try {
    const valRes = await fetch('/api/dst/validate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const valData = await valRes.json().catch(() => ({}));
    if (!valRes.ok || !valData.valid) {
      const errMsg = (valData.errors && valData.errors.length > 0) ? valData.errors.join("\n") : (valData.error || "Validation failed.");
      alert("Input Error\n\n" + errMsg);
      return;
    }

    if (valData.warnings && valData.warnings.length > 0) {
      const ok = confirm("Target Strength Warning\n\n" + valData.warnings.join("\n") + "\n\nContinue with this target?");
      if (!ok) return;
    }

    const genRes = await fetch('/api/dst/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const genData = await genRes.json().catch(() => ({}));

    if (!genRes.ok || !genData.success) {
      alert("Generation Error\n\n" + (genData.error || "Server generation failed."));
      return;
    }

    if (genData.results) {
      results = genData.results;
      refreshResults();
      if (results.length > 0) {
        selectResult(0);
      }
      statusEl.textContent = `Generated ${results.length} depth(s) × 3 trials × 61 observations.`;
    }
  } catch (e) {
    alert("Generation Error\n\n" + e.message);
  }
}

function refreshResults() {
  const tbody = document.getElementById('resultsTableBody');
  tbody.innerHTML = '';

  results.forEach((r, idx) => {
    const s = r.summary;
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${Number(r.depth).toFixed(2)}</td>
      <td>${r.soil}</td>
      <td>${r.sample_type}</td>
      <td>${Number(r.c).toFixed(3)}</td>
      <td>${Number(r.phi).toFixed(2)}</td>
      <td>${Number(s.calculated_c).toFixed(3)}</td>
      <td>${Number(s.calculated_phi).toFixed(2)}</td>
      <td>${Number(s.r_squared).toFixed(3)}</td>
      <td>3</td>
    `;
    tr.addEventListener('click', () => selectResult(idx));
    tbody.appendChild(tr);
  });
}

function selectResult(index) {
  if (!results || !results[index]) return;
  selectedResult = results[index];

  const rows = document.querySelectorAll('#resultsTableBody tr');
  rows.forEach((r, i) => {
    if (i === index) r.classList.add('selected');
    else r.classList.remove('selected');
  });

  redraw();
}

// -------------------------------------------------------------
// EXACT CANVAS REDRAW MATCHING DST_MultiDepth_Excel_Import.py
// -------------------------------------------------------------
function redraw() {
  const canvas = document.getElementById('dstCanvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const width = canvas.width;
  const height = canvas.height;

  ctx.clearRect(0, 0, width, height);

  const result = selectedResult;
  if (!result || !result.trials) {
    ctx.fillStyle = "#000000";
    ctx.font = "bold 14pt 'Segoe UI', Tahoma, sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";
    ctx.fillText("Generate data to display the DST Failure Envelope.", width / 2, height / 2);
    return;
  }

  const bh = document.getElementById('bh_id').value.trim();
  const trials = result.trials;
  const summary = result.summary;
  const c_calc = Number(summary.calculated_c);
  const phi_calc = Number(summary.calculated_phi);
  const r2 = Number(summary.r_squared);
  const slope = Math.tan((phi_calc * Math.PI) / 180.0);

  const margin = 8;
  const x0 = margin;
  const x1 = width - margin;
  let y = 6;
  const title_h = 58;
  const info_h = 70;
  const graph_title_h = 35;
  const graph_h = 410;
  const result_title_h = 35;
  const table_header_h = 50;
  const table_row_h = 38;

  ctx.lineWidth = 1;
  ctx.strokeStyle = "#222222";

  // 1. Outer title box
  ctx.lineWidth = 2;
  ctx.strokeRect(x0, y, x1 - x0, title_h);
  ctx.fillStyle = "#172554";
  ctx.font = "bold 20pt 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("DIRECT SHEAR TEST – FAILURE ENVELOPE", (x0 + x1) / 2, y + 22);

  ctx.font = "bold 10.5pt 'Segoe UI', sans-serif";
  ctx.fillText("IS 2720 (Part 13) – 1986", (x0 + x1) / 2, y + 47);
  y += title_h;

  // 2. Top information table
  ctx.lineWidth = 1;
  ctx.strokeRect(x0, y, x1 - x0, info_h);

  const c1 = x0 + (x1 - x0) * 0.28;
  const c2 = x0 + (x1 - x0) * 0.52;
  const c3 = x0 + (x1 - x0) * 0.78;

  ctx.beginPath();
  [c1, c2, c3].forEach(xx => {
    ctx.moveTo(xx, y);
    ctx.lineTo(xx, y + info_h);
  });
  ctx.stroke();

  const sample_display = result.sample_type === "Undisturbed Sample" ? "UDS" : "Remoulded";
  const depth_text = `${Number(result.depth).toFixed(2)} m`;
  const density_text = `${Number(result.density).toFixed(2)} gm/cc`;
  const weight_text = `${Number(result.weight).toFixed(2)} g`;
  const normal_text = trials.map(t => Number(t.normal_stress).toFixed(2)).join('/') + " kg/cm²";

  function reportPair(px, py, label, value) {
    ctx.textAlign = "left";
    ctx.textBaseline = "middle";
    ctx.fillStyle = "#000000";
    ctx.font = "bold 9pt 'Segoe UI', sans-serif";
    ctx.fillText(label, px, py);
    ctx.fillText(":", px + 115, py);
    ctx.font = "9pt 'Segoe UI', sans-serif";
    ctx.fillText(String(value), px + 130, py);
  }

  reportPair(x0 + 14, y + 22, "Borehole ID", bh);
  reportPair(x0 + 14, y + 49, "Depth", depth_text);

  reportPair(c1 + 14, y + 22, "Sample Type", sample_display);
  reportPair(c1 + 14, y + 49, "Density", density_text);

  reportPair(c2 + 14, y + 22, "Shear Box Size", "60 mm × 60 mm");
  reportPair(c2 + 14, y + 49, "Sample Weight", weight_text);

  reportPair(c3 + 14, y + 22, "Condition", sample_display);
  reportPair(c3 + 14, y + 49, "Normal Stress", normal_text);
  y += info_h;

  // 3. Graph heading
  ctx.strokeRect(x0, y, x1 - x0, graph_title_h);
  ctx.fillStyle = "#000000";
  ctx.font = "bold 11pt 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("NORMAL STRESS vs PEAK SHEAR STRESS (FAILURE ENVELOPE)", (x0 + x1) / 2, y + graph_title_h / 2);
  y += graph_title_h;

  // 4. Graph area
  const graph_top = y;
  const graph_bottom = y + graph_h;
  ctx.strokeRect(x0, graph_top, x1 - x0, graph_h);

  const left = x0 + 90;
  const right = x1 - 40;
  const top = graph_top + 40;
  const bottom = graph_bottom - 60;
  const pw = right - left;
  const ph = bottom - top;

  const normal = trials.map(t => Number(t.normal_stress));
  const peak = trials.map(t => Number(t.peak_shear_stress));
  const max_sigma = Math.max(...normal);
  const max_tau = Math.max(...peak);
  let xlim = Math.max(1.0, Math.ceil((max_sigma * 1.18) / 0.1) * 0.1);
  if (max_sigma <= 1.0) xlim = 1.2;
  let ylim = Math.max(0.4, Math.ceil((max_tau * 1.20) / 0.1) * 0.1);
  if (max_tau <= 1.0) ylim = 1.2;

  function xp(v) { return left + (v / xlim) * pw; }
  function yp(v) { return bottom - (v / ylim) * ph; }

  // Grid and ticks
  const x_step = xlim <= 1.2 ? 0.2 : 0.5;
  const y_step = ylim <= 1.2 ? 0.2 : 0.5;

  ctx.strokeStyle = "#d7d7d7";
  ctx.setLineDash([3, 3]);

  for (let xv = 0.0; xv <= xlim + 1e-9; xv += x_step) {
    const xx = xp(xv);
    ctx.beginPath();
    ctx.moveTo(xx, top);
    ctx.lineTo(xx, bottom);
    ctx.stroke();

    ctx.fillStyle = "#000000";
    ctx.font = "9pt 'Segoe UI', sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "top";
    ctx.fillText(xv.toFixed(2), xx, bottom + 8);
  }

  for (let yv = 0.0; yv <= ylim + 1e-9; yv += y_step) {
    const yy = yp(yv);
    ctx.beginPath();
    ctx.moveTo(left, yy);
    ctx.lineTo(right, yy);
    ctx.stroke();

    ctx.fillStyle = "#000000";
    ctx.font = "9pt 'Segoe UI', sans-serif";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    ctx.fillText(yv.toFixed(2), left - 8, yy);
  }

  ctx.setLineDash([]);
  ctx.strokeStyle = "#000000";
  ctx.lineWidth = 1.8;
  ctx.beginPath();
  ctx.moveTo(left, top);
  ctx.lineTo(left, bottom);
  ctx.lineTo(right, bottom);
  ctx.stroke();

  // Axis labels
  ctx.fillStyle = "#000000";
  ctx.font = "bold 10pt 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("Normal Stress, σₙ (kg/cm²)", (left + right) / 2, graph_bottom - 20);

  ctx.save();
  ctx.translate(x0 + 25, (top + bottom) / 2);
  ctx.rotate(-Math.PI / 2);
  ctx.fillText("Peak Shear Stress, τ (kg/cm²)", 0, 0);
  ctx.restore();

  // Failure Envelope line (Dashed start + Solid line)
  const first_sigma = Math.min(...normal);
  const first_y = c_calc + slope * first_sigma;
  const env_end_y = c_calc + slope * xlim;

  ctx.strokeStyle = "#e00000";
  ctx.lineWidth = 2.0;

  if (c_calc >= 0 && first_sigma > 0) {
    ctx.setLineDash([7, 5]);
    ctx.beginPath();
    ctx.moveTo(xp(0), yp(c_calc));
    ctx.lineTo(xp(first_sigma), yp(first_y));
    ctx.stroke();
  }

  ctx.setLineDash([]);
  ctx.beginPath();
  ctx.moveTo(xp(first_sigma), yp(first_y));
  ctx.lineTo(xp(xlim), yp(env_end_y));
  ctx.stroke();

  // Peak points + coordinate labels
  trials.forEach(t => {
    const xx = xp(Number(t.normal_stress));
    const yy = yp(Number(t.peak_shear_stress));

    ctx.fillStyle = "#ffffff";
    ctx.strokeStyle = "#111111";
    ctx.lineWidth = 2.0;
    ctx.beginPath();
    ctx.arc(xx, yy, 6, 0, 2 * Math.PI);
    ctx.fill();
    ctx.stroke();

    ctx.fillStyle = "#172bb5";
    ctx.font = "bold 9pt 'Segoe UI', sans-serif";
    ctx.textAlign = "center";
    ctx.textBaseline = "bottom";
    ctx.fillText(`(${Number(t.normal_stress).toFixed(2)}, ${Number(t.peak_shear_stress).toFixed(2)})`, xx, yy - 8);
  });

  // Legend box (Top-left)
  const lx = left + 15;
  const ly = top + 15;
  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#222222";
  ctx.lineWidth = 1;
  ctx.fillRect(lx, ly, 195, 65);
  ctx.strokeRect(lx, ly, 195, 65);

  ctx.beginPath();
  ctx.arc(lx + 24, ly + 20, 5, 0, 2 * Math.PI);
  ctx.fillStyle = "#ffffff";
  ctx.fill();
  ctx.strokeStyle = "#111111";
  ctx.stroke();

  ctx.fillStyle = "#000000";
  ctx.font = "9pt 'Segoe UI', sans-serif";
  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillText("Test Results (Peak)", lx + 42, ly + 20);

  ctx.strokeStyle = "#e00000";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(lx + 14, ly + 46);
  ctx.lineTo(lx + 34, ly + 46);
  ctx.stroke();

  ctx.fillStyle = "#000000";
  ctx.fillText("Failure Envelope", lx + 42, ly + 46);

  // Equation box (Bottom-right)
  const bx1 = right - 260;
  const by1 = bottom - 150;
  const bw = 250;
  const bh_h = 140;

  ctx.fillStyle = "#ffffff";
  ctx.strokeStyle = "#222222";
  ctx.lineWidth = 1;
  ctx.fillRect(bx1, by1, bw, bh_h);
  ctx.strokeRect(bx1, by1, bw, bh_h);

  ctx.textAlign = "left";
  ctx.textBaseline = "middle";
  ctx.fillStyle = "#000000";
  ctx.font = "bold 10pt 'Segoe UI', sans-serif";
  ctx.fillText("τ = c + σₙ tan φ", bx1 + 14, by1 + 20);

  ctx.font = "bold 9pt 'Segoe UI', sans-serif";
  ctx.fillText(`= ${c_calc.toFixed(2)} + σₙ tan ${phi_calc.toFixed(1)}°`, bx1 + 14, by1 + 44);

  ctx.fillStyle = "#e00000";
  ctx.font = "bold 10pt 'Segoe UI', sans-serif";
  ctx.fillText(`c = ${c_calc.toFixed(2)} kg/cm²`, bx1 + 14, by1 + 75);
  ctx.fillText(`φ = ${phi_calc.toFixed(1)}°`, bx1 + 14, by1 + 99);

  ctx.fillStyle = "#000000";
  ctx.font = "bold 10pt 'Segoe UI', sans-serif";
  ctx.fillText(`R² = ${r2.toFixed(3)}`, bx1 + 14, by1 + 123);

  y = graph_bottom;

  // 5. Bottom results title
  ctx.strokeStyle = "#222222";
  ctx.lineWidth = 1;
  ctx.strokeRect(x0, y, x1 - x0, result_title_h);
  ctx.fillStyle = "#172bb5";
  ctx.font = "bold 11.5pt 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.fillText("DIRECT SHEAR TEST RESULTS", (x0 + x1) / 2, y + result_title_h / 2);
  y += result_title_h;

  // 6. Bottom results table
  const rel = [0.07, 0.13, 0.12, 0.11, 0.15, 0.16, 0.11, 0.15];
  const colx = [x0];
  rel.forEach(frac => colx.push(colx[colx.length - 1] + (x1 - x0) * frac));

  const headers = [
    "Sr. No.", "Sample Depth\n(m)", "Sample Type", "Density\n(gm/cc)",
    "Normal Stress,\nσₙ (kg/cm²)", "Shear Stress at Failure\n(Peak), τ (kg/cm²)",
    "Cohesion,\nc (kg/cm²)", "Angle of Internal Friction,\nφ (°)"
  ];

  ctx.fillStyle = "#f1f4fb";
  ctx.fillRect(x0, y, x1 - x0, table_header_h);
  ctx.strokeRect(x0, y, x1 - x0, table_header_h);

  for (let i = 1; i < colx.length - 1; i++) {
    ctx.beginPath();
    ctx.moveTo(colx[i], y);
    ctx.lineTo(colx[i], y + table_header_h + 3 * table_row_h);
    ctx.stroke();
  }

  ctx.fillStyle = "#000000";
  ctx.font = "bold 8pt 'Segoe UI', sans-serif";
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";

  headers.forEach((h, i) => {
    const lines = h.split('\n');
    const cx = (colx[i] + colx[i + 1]) / 2;
    if (lines.length === 1) {
      ctx.fillText(lines[0], cx, y + table_header_h / 2);
    } else {
      ctx.fillText(lines[0], cx, y + table_header_h / 2 - 7);
      ctx.fillText(lines[1], cx, y + table_header_h / 2 + 7);
    }
  });

  y += table_header_h;

  // 3 Trial rows
  const table_y0 = y;
  for (let i = 0; i < trials.length; i++) {
    const t = trials[i];
    ctx.strokeRect(x0, y, x1 - x0, table_row_h);

    const vals = [
      String(i + 1),
      i === 1 ? depth_text : "",
      i === 1 ? sample_display : "",
      i === 1 ? Number(result.density).toFixed(2) : "",
      Number(t.normal_stress).toFixed(2),
      Number(t.peak_shear_stress).toFixed(2),
      i === 1 ? c_calc.toFixed(2) : "",
      i === 1 ? phi_calc.toFixed(1) : ""
    ];

    ctx.font = "9pt 'Segoe UI', sans-serif";
    ctx.fillStyle = "#000000";
    vals.forEach((val, j) => {
      if (val) {
        const cx = (colx[j] + colx[j + 1]) / 2;
        const cy = [1, 2, 3, 6, 7].includes(j) ? table_y0 + 1.5 * table_row_h : y + table_row_h / 2;
        ctx.fillText(val, cx, cy);
      }
    });

    y += table_row_h;
  }
}

function triggerBrowserFileDownload(data, fallbackFilename, mimeType = "application/zip") {
  const filename = data.zip_filename || data.excel_filename || data.filename || fallbackFilename;

  // 1. Direct Blob Download via Base64 (Reliable client-side download across any remote machine)
  if (data.file_base64) {
    try {
      const byteCharacters = atob(data.file_base64);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: mimeType });
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = blobUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(blobUrl);
      }, 300);
      return;
    } catch (err) {
      console.warn("Base64 blob conversion fallback", err);
    }
  }

  // 2. URL Fallback
  const url = data.download_zip_url || data.download_excel_url || data.download_url || data.url;
  if (url) {
    const a = document.createElement('a');
    a.href = url.includes("?") ? `${url}&download=1` : `${url}?download=1`;
    a.download = filename;
    a.target = '_blank';
    document.body.appendChild(a);
    a.click();
    setTimeout(() => {
      document.body.removeChild(a);
    }, 300);
  }
}

async function exportAllExcel() {
  if (!results || results.length === 0) {
    alert("Export\n\nPlease generate all depth data first.");
    return;
  }

  const bhId = document.getElementById('bh_id').value.trim() || 'DST_Result';
  const payload = {
    bh_id: bhId,
    pr_constant: parseFloat(document.getElementById('pr_constant').value),
    normal_1: parseFloat(document.getElementById('normal_1').value),
    normal_2: parseFloat(document.getElementById('normal_2').value),
    normal_3: parseFloat(document.getElementById('normal_3').value),
    results: results
  };

  document.getElementById('statusVar').textContent = "Generating DST Excel & Failure Plots Package... Please wait.";

  try {
    const res = await fetch('/api/dst/export_excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.success) {
      throw new Error(data.error || "Export failed on server.");
    }

    const fallbackZip = `DST_Package_${bhId}.zip`;
    triggerBrowserFileDownload(data, fallbackZip, "application/zip");

    alert(
      `Generation Completed\n\n` +
      `DST Multi-Depth generation completed successfully in output folder.\n\n` +
      `Depths generated: ${data.count}\n\n` +
      `ZIP Package (Downloaded):\n${data.zip_filename || fallbackZip}\n\n` +
      `Excel output:\n${data.excel_path}\n\n` +
      `Failure plots folder:\n${data.plot_dir}`
    );
    document.getElementById('statusVar').textContent = `Package downloaded to your computer: ${data.zip_filename || fallbackZip}`;
  } catch (e) {
    alert("DST Export Error\n\n" + e.message);
    document.getElementById('statusVar').textContent = "Export failed.";
  }
}

async function exportPlots() {
  if (!results || results.length === 0) {
    alert("Plot Export\n\nPlease generate all depth test results first.");
    return;
  }

  const bhId = document.getElementById('bh_id').value.trim() || 'DST_Result';
  const payload = {
    bh_id: bhId,
    results: results
  };

  document.getElementById('statusVar').textContent = "Exporting Failure Envelope Plots... Please wait.";

  try {
    const res = await fetch('/api/dst/export_plots', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await res.json().catch(() => ({}));
    if (!res.ok || !data.success) {
      throw new Error(data.error || "Plot export failed on server.");
    }

    const fallbackZip = `${bhId}_Failure_Plots.zip`;
    triggerBrowserFileDownload(data, fallbackZip, "application/zip");

    alert(
      `Plot Export Complete\n\n` +
      `Generated ${data.count} separate failure-envelope plot(s).\n\n` +
      `ZIP Archive (${data.zip_filename || fallbackZip}) has been downloaded directly to your computer's Downloads folder.`
    );
    document.getElementById('statusVar').textContent = `Plots ZIP downloaded to your computer: ${data.zip_filename || fallbackZip}`;
  } catch (e) {
    alert("Plot Export Error\n\n" + e.message);
    document.getElementById('statusVar').textContent = "Plot export failed.";
  }
}

async function importFromExcel(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/dst/import_excel', {
      method: 'POST',
      body: formData
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Excel Import Error\n\n" + (data.error || "Failed to parse Excel input."));
      return;
    }

    if (data.bh_id !== undefined) document.getElementById('bh_id').value = data.bh_id;
    if (data.pr_constant !== undefined) document.getElementById('pr_constant').value = data.pr_constant;
    if (data.normal_1 !== undefined) document.getElementById('normal_1').value = data.normal_1;
    if (data.normal_2 !== undefined) document.getElementById('normal_2').value = data.normal_2;
    if (data.normal_3 !== undefined) document.getElementById('normal_3').value = data.normal_3;
    if (data.random_seed !== undefined) document.getElementById('random_seed').value = data.random_seed;

    if (data.rows && data.rows.length > 0) {
      depthRows = data.rows.map((r, i) => ({
        id: 'row_' + Date.now() + '_' + i,
        depth: r.depth.toString(),
        soil: r.soil,
        sample_type: r.sample_type,
        density: r.density.toString(),
        weight: r.weight.toString(),
        c: r.c.toString(),
        phi: r.phi.toString(),
        behaviour: r.behaviour
      }));
      renderDepthTable();
    }

    results = [];
    selectedResult = null;
    refreshResults();
    redraw();
    document.getElementById('statusVar').textContent = `Imported ${depthRows.length} depth(s) from Excel. Review/validate, then generate.`;

    alert(
      `Excel Import Complete\n\n` +
      `Imported ${depthRows.length} depth(s) successfully.\n\n` +
      `Review the imported values and click GENERATE ALL DEPTHS.`
    );
  } catch (e) {
    alert("Excel Import Error\n\n" + e.message);
  } finally {
    e.target.value = '';
  }
}

function resetForm() {
  document.getElementById('bh_id').value = "";
  document.getElementById('pr_constant').value = "0.325";
  document.getElementById('normal_1').value = "0.50";
  document.getElementById('normal_2').value = "1.00";
  document.getElementById('normal_3').value = "1.50";
  document.getElementById('random_seed').value = "AUTO";

  depthRows = [];
  results = [];
  selectedResult = null;

  addDepthRow();
  document.getElementById('resultsTableBody').innerHTML = '';
  document.getElementById('statusVar').textContent = "Enter one or more depths, then generate.";

  redraw();
}
