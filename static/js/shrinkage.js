// Shrinkage Limit Test – Observation Sheet Generator V5.2
// 100% faithful to shrinkage_limit_generator_v5.py

const DISHES_DATA = {
  "SL_01": { weight: 37.10, diameter: 44.34, height: 14.20 },
  "SL_02": { weight: 36.38, diameter: 44.36, height: 14.18 },
  "SL_03": { weight: 37.29, diameter: 44.34, height: 15.08 },
  "SL_04": { weight: 41.98, diameter: 45.62, height: 16.34 },
  "SL_05": { weight: 47.97, diameter: 45.50, height: 16.40 },
  "SL_06": { weight: 51.18, diameter: 45.21, height: 16.24 },
  "SL_07": { weight: 52.38, diameter: 45.22, height: 16.53 },
  "SL_08": { weight: 47.07, diameter: 44.41, height: 16.08 },
  "SL_09": { weight: 57.57, diameter: 44.89, height: 16.36 },
  "SL_10": { weight: 33.27, diameter: 44.07, height: 14.98 },
};

const CONTAINERS = Object.keys(DISHES_DATA);
const SOIL_CLASSES = ["CL", "CI", "CH"];
const SAMPLE_TYPES = ["UDS", "SPT", "DS", "Block Sample", "Other"];

let dishesList = [];
let inputRows = [];
let generatedObservations = [];
let selectedRowIndices = new Set();

const STORAGE_KEY_INPUTS = "goma_shrinkage_inputs";
const STORAGE_KEY_OBS = "goma_shrinkage_observations";
const STORAGE_KEY_BH = "goma_shrinkage_bh_id";

document.addEventListener('DOMContentLoaded', async () => {
  await loadDishDetails();
  populateDishSelect();

  // Restore saved BH ID or default
  const savedBh = localStorage.getItem(STORAGE_KEY_BH);
  if (savedBh !== null) {
    document.getElementById('bh_id').value = savedBh;
  }

  // Restore saved input rows or start clean as in Tkinter script
  const savedInputs = localStorage.getItem(STORAGE_KEY_INPUTS);
  if (savedInputs) {
    try {
      inputRows = JSON.parse(savedInputs) || [];
    } catch (e) {
      inputRows = [];
    }
  }

  renderInputDepthsTable();

  // Restore saved observations if available
  const savedObs = localStorage.getItem(STORAGE_KEY_OBS);
  if (savedObs) {
    try {
      generatedObservations = JSON.parse(savedObs) || [];
      if (generatedObservations.length > 0) {
        renderObservationTable();
      }
    } catch (e) {
      generatedObservations = [];
    }
  }

  // Set default initial state for input row matching Tkinter script
  document.getElementById('inpSampleType').value = "UDS";
  document.getElementById('inpSoilClass').value = "CL";
  
  // Set next available dish
  const assigned = new Set(inputRows.map(r => r.dish));
  const nextDish = CONTAINERS.find(c => !assigned.has(c)) || "SL_01";
  document.getElementById('inpDish').value = nextDish;

  updateDishInfoLine();
  bindEvents();
});

function calculateDishVolume(diameterMm, heightMm) {
  const dCm = diameterMm / 10.0;
  const hCm = heightMm / 10.0;
  return Math.PI * Math.pow(dCm, 2) / 4.0 * hCm;
}

async function loadDishDetails() {
  try {
    const res = await fetch('/api/shrinkage/dishes');
    const data = await res.json();
    if (data.dishes && data.dishes.length > 0) {
      dishesList = data.dishes;
    } else {
      dishesList = CONTAINERS.map((c, i) => {
        const d = DISHES_DATA[c];
        return {
          sr: i + 1,
          dish: c,
          weight: d.weight,
          diameter: d.diameter,
          height: d.height,
          volume: calculateDishVolume(d.diameter, d.height)
        };
      });
    }
  } catch (e) {
    dishesList = CONTAINERS.map((c, i) => {
      const d = DISHES_DATA[c];
      return {
        sr: i + 1,
        dish: c,
        weight: d.weight,
        diameter: d.diameter,
        height: d.height,
        volume: calculateDishVolume(d.diameter, d.height)
      };
    });
  }
  renderDishDetailsTable();
}

function renderDishDetailsTable() {
  const tbody = document.getElementById('dishDetailsTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  dishesList.forEach(d => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">${d.sr}</td>
      <td style="text-align: center; font-weight: 500;">${d.dish}</td>
      <td style="text-align: center;">${d.weight.toFixed(2)}</td>
      <td style="text-align: center;">${d.diameter.toFixed(2)}</td>
      <td style="text-align: center;">${d.height.toFixed(2)}</td>
      <td style="text-align: center; font-weight: 500;">${d.volume.toFixed(2)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function populateDishSelect() {
  const sel = document.getElementById('inpDish');
  if (!sel) return;
  sel.innerHTML = '';
  CONTAINERS.forEach(c => {
    const opt = document.createElement('option');
    opt.value = c;
    opt.textContent = c;
    sel.appendChild(opt);
  });
}

function updateDishInfoLine() {
  const selDish = document.getElementById('inpDish').value;
  const line = document.getElementById('dishInfoLine');
  if (!line) return;

  const d = DISHES_DATA[selDish];
  if (d) {
    const vol = calculateDishVolume(d.diameter, d.height);
    line.textContent = `${selDish}: Weight = ${d.weight.toFixed(2)} g | Diameter = ${d.diameter.toFixed(2)} mm | Height = ${d.height.toFixed(2)} mm | Calculated cylindrical volume = ${vol.toFixed(2)} cm³`;
  }
}

function classifyFromLL(ll) {
  if (ll < 35.0) return "CL";
  if (ll <= 50.0) return "CI";
  return "CH";
}

function plasticityDescription(ll) {
  if (ll < 35.0) return "Low plasticity";
  if (ll <= 50.0) return "Intermediate plasticity";
  return "High plasticity";
}

function validateSoilClass(ll, soilClass) {
  const expected = classifyFromLL(ll);
  if (soilClass !== expected) {
    return {
      valid: false,
      error: `LL = ${ll}% corresponds to ${expected} (${plasticityDescription(ll)}), not ${soilClass}.`
    };
  }
  return { valid: true, error: "" };
}

function autoSaveState() {
  localStorage.setItem(STORAGE_KEY_INPUTS, JSON.stringify(inputRows));
  localStorage.setItem(STORAGE_KEY_OBS, JSON.stringify(generatedObservations));
  const bh = document.getElementById('bh_id').value;
  localStorage.setItem(STORAGE_KEY_BH, bh);
}

function bindEvents() {
  document.getElementById('bh_id').addEventListener('input', autoSaveState);
  document.getElementById('inpDish').addEventListener('change', updateDishInfoLine);

  document.getElementById('btnAddDepth').addEventListener('click', handleAddDepth);
  document.getElementById('btnRemoveSelected').addEventListener('click', handleRemoveSelected);
  document.getElementById('btnClearInputs').addEventListener('click', handleClearInputs);
  document.getElementById('btnGenerate').addEventListener('click', handleGenerate);
  document.getElementById('btnExportExcel').addEventListener('click', handleExportExcel);
  document.getElementById('btnClearResults').addEventListener('click', handleClearResults);
}

function handleAddDepth() {
  const depthStr = document.getElementById('inpDepth').value.trim();
  const sample = document.getElementById('inpSampleType').value.trim();
  const soil = document.getElementById('inpSoilClass').value.trim();
  const dish = document.getElementById('inpDish').value.trim();
  const llStr = document.getElementById('inpLL').value.trim();
  const plStr = document.getElementById('inpPL').value.trim();

  if (!depthStr) {
    alert("Missing Depth\n\nEnter the depth.");
    return;
  }

  const depth = parseFloat(depthStr);
  if (isNaN(depth) || depth < 0) {
    alert("Invalid Depth\n\nEnter a valid positive number for Depth.");
    return;
  }

  if (!sample) {
    alert("Missing Sample Type\n\nSelect Sample Type.");
    return;
  }

  if (!SOIL_CLASSES.includes(soil)) {
    alert("Missing Soil Class\n\nSelect CL, CI or CH.");
    return;
  }

  if (!CONTAINERS.includes(dish)) {
    alert("Missing Dish\n\nSelect SL_01 to SL_10.");
    return;
  }

  const ll = parseFloat(llStr);
  const pl = parseFloat(plStr);
  if (isNaN(ll) || isNaN(pl)) {
    alert("Invalid LL/PL\n\nEnter a valid positive number for LL/PL.");
    return;
  }

  if (ll <= 0 || pl <= 0) {
    alert("Invalid LL/PL\n\nLL and PL must be greater than zero.");
    return;
  }

  if (pl >= ll) {
    alert("Invalid LL/PL\n\nPL must be less than LL.");
    return;
  }

  const check = validateSoilClass(ll, soil);
  if (!check.valid) {
    alert(`Soil Classification Mismatch\n\n${check.error}`);
    return;
  }

  for (const r of inputRows) {
    if (r.dish === dish) {
      alert(`Duplicate Dish\n\n${dish} is already assigned to another depth.`);
      return;
    }
  }

  const pi = parseFloat((ll - pl).toFixed(2));
  addDepthRow({
    depth: depthStr,
    sample_type: sample,
    soil_class: soil,
    dish: dish,
    ll: ll,
    pl: pl,
    pi: pi
  });

  // Clear inputs matching Tkinter script behavior
  document.getElementById('inpDepth').value = "";
  document.getElementById('inpLL').value = "";
  document.getElementById('inpPL').value = "";

  // Advance dish to next unassigned container
  const assigned = new Set(inputRows.map(r => r.dish));
  const nextDish = CONTAINERS.find(c => !assigned.has(c));
  if (nextDish) {
    document.getElementById('inpDish').value = nextDish;
    updateDishInfoLine();
  }

  autoSaveState();
}

function addDepthRow(row) {
  inputRows.push(row);
  renderInputDepthsTable();
  autoSaveState();
}

function renderInputDepthsTable() {
  const tbody = document.getElementById('inputDepthsTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (inputRows.length === 0) {
    tbody.innerHTML = `
      <tr id="emptyInputRow">
        <td colspan="7" style="text-align: center; color: #888; font-style: italic; padding: 12px;">
          No depth rows added yet. Enter values above and click "Add Depth".
        </td>
      </tr>
    `;
    selectedRowIndices.clear();
    return;
  }

  inputRows.forEach((r, idx) => {
    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';
    if (selectedRowIndices.has(idx)) {
      tr.className = 'tk-tree-selected';
    }

    tr.innerHTML = `
      <td style="text-align: center;">${r.depth}</td>
      <td style="text-align: center;">${r.sample_type}</td>
      <td style="text-align: center; font-weight: 500;">${r.soil_class}</td>
      <td style="text-align: center;">${r.dish}</td>
      <td style="text-align: center;">${parseFloat(r.ll).toFixed(2)}</td>
      <td style="text-align: center;">${parseFloat(r.pl).toFixed(2)}</td>
      <td style="text-align: center; font-weight: 500;">${parseFloat(r.pi).toFixed(2)}</td>
    `;

    tr.addEventListener('click', () => {
      if (selectedRowIndices.has(idx)) {
        selectedRowIndices.delete(idx);
      } else {
        selectedRowIndices.add(idx);
      }
      renderInputDepthsTable();
    });

    tbody.appendChild(tr);
  });
}

function handleRemoveSelected() {
  if (selectedRowIndices.size === 0) {
    if (inputRows.length > 0) {
      inputRows.pop();
      renderInputDepthsTable();
      autoSaveState();
    }
    return;
  }

  inputRows = inputRows.filter((_, idx) => !selectedRowIndices.has(idx));
  selectedRowIndices.clear();
  renderInputDepthsTable();
  autoSaveState();
}

function handleClearInputs() {
  inputRows = [];
  selectedRowIndices.clear();
  renderInputDepthsTable();
  autoSaveState();
}

function handleClearResults() {
  generatedObservations = [];
  const tbody = document.getElementById('observationTableBody');
  if (tbody) {
    tbody.innerHTML = `
      <tr id="emptyResultRow">
        <td colspan="14" style="text-align: center; color: #888; font-style: italic; padding: 20px;">
          No observation data generated yet. Add depths above and click "Generate Observation Data".
        </td>
      </tr>
    `;
  }
  autoSaveState();
}

async function handleGenerate() {
  const bhId = document.getElementById('bh_id').value.trim();
  if (!bhId) {
    alert("Missing BH ID\n\nEnter BH ID.");
    return;
  }

  if (inputRows.length === 0) {
    alert("No Depths\n\nAdd at least one depth.");
    return;
  }

  handleClearResults();

  try {
    const payload = {
      bh_id: bhId,
      rows: inputRows
    };

    const res = await fetch('/api/shrinkage/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });

    const data = await res.json();
    if (!res.ok || data.error) {
      alert(`Generation Error\n\n${data.error || "Failed to generate observations."}`);
      return;
    }

    generatedObservations = data.results;
    renderObservationTable();
    autoSaveState();

    alert(`Completed\n\nGenerated observations for ${generatedObservations.length} depth(s).`);
  } catch (e) {
    alert(`Generation Error\n\n${e.message}`);
  }
}

function renderObservationTable() {
  const tbody = document.getElementById('observationTableBody');
  if (!tbody) return;
  tbody.innerHTML = '';

  if (generatedObservations.length === 0) {
    tbody.innerHTML = `
      <tr id="emptyResultRow">
        <td colspan="14" style="text-align: center; color: #888; font-style: italic; padding: 20px;">
          No observation data generated yet. Add depths above and click "Generate Observation Data".
        </td>
      </tr>
    `;
    return;
  }

  generatedObservations.forEach(o => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">${o.bh_id}</td>
      <td style="text-align: center;">${parseFloat(o.depth).toFixed(2)}</td>
      <td style="text-align: center;">${o.sample_type}</td>
      <td style="text-align: center; font-weight: 500;">${o.soil_class}</td>
      <td style="text-align: center;">${o.dish}</td>
      <td style="text-align: center;">${o.w1.toFixed(2)}</td>
      <td style="text-align: center;">${o.w2.toFixed(2)}</td>
      <td style="text-align: center;">${o.w3.toFixed(2)}</td>
      <td style="text-align: center;">${o.m1.toFixed(2)}</td>
      <td style="text-align: center;">${o.m2.toFixed(2)}</td>
      <td style="text-align: center;">${o.v1.toFixed(2)}</td>
      <td style="text-align: center;">${o.mv.toFixed(2)}</td>
      <td style="text-align: center;">${o.v2.toFixed(2)}</td>
      <td style="text-align: center; font-weight: bold; color: #b91c1c;">${o.sl.toFixed(2)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function triggerSafeDownload(url, filename) {
  const a = document.createElement('a');
  a.href = url;
  if (filename) a.setAttribute('download', filename);
  a.target = '_blank';
  document.body.appendChild(a);
  a.click();
  a.remove();
}

async function handleExportExcel() {
  if (!generatedObservations || generatedObservations.length === 0) {
    alert("No Data\n\nGenerate observation data first.");
    return;
  }

  const bhId = document.getElementById('bh_id').value.trim() || "RBH-12";
  try {
    const res = await fetch('/api/shrinkage/export_excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        bh_id: bhId,
        input_rows: inputRows,
        results: generatedObservations
      })
    });

    const data = await res.json();
    if (!res.ok || data.error) {
      throw new Error(data.error || "Excel export failed.");
    }

    if (data.excel_url) {
      triggerSafeDownload(data.excel_url, data.filename);
    }

    alert(`Export Complete\n\nObservation sheet saved to:\n${data.filename || data.excel_path}`);
  } catch (e) {
    alert(`Export Error\n\n${e.message}`);
  }
}
