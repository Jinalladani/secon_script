// Soil & Groundwater Chemical Test Generator Controller v2
// 100% Faithful to chemical_test_generator_v13_input_controlled_groundwater.py

const SOIL_TYPES = ["CL", "CI", "CH", "ML", "SM", "SC", "SP", "SW", "GP", "GW", "GM", "GC", "SILT", "CLAY"];
const CONDITIONS = ["Dry", "Slightly Moist", "Moist", "Wet", "Saturated", "Saline", "Highly Saline"];
const SALINITY_CONTROLS = ["Low", "Normal", "High", "Very High"];
const GW_CONDITIONS = ["Normal", "Slightly Saline", "Saline", "Highly Saline"];
const GW_SALINITY_CONTROLS = ["Low", "Normal", "High", "Very High"];
const RESIDUE_OPTIONS = ["None", "Slight", "Visible", "Strong"];

let generatedData = null;

document.addEventListener('DOMContentLoaded', () => {
  initDefaults();
  bindEvents();
  buildInitialRows(5);
});

function initDefaults() {
  const seedInput = document.getElementById('seedVar');
  if (!seedInput.value) {
    seedInput.value = Math.floor(100000 + Math.random() * 900000);
  }
}

function bindEvents() {
  // Tabs Navigation
  document.querySelectorAll('.chem-tab-btn').forEach(btn => {
    btn.addEventListener('click', () => {
      const tabId = btn.getAttribute('data-tab');
      switchTab(tabId);
    });
  });

  // Randomize Seed
  document.getElementById('btnRandomSeed').addEventListener('click', () => {
    document.getElementById('seedVar').value = Math.floor(100000 + Math.random() * 900000);
  });

  // Count Spinner change
  document.getElementById('countVar').addEventListener('change', (e) => {
    let n = parseInt(e.target.value);
    if (isNaN(n) || n < 1) n = 1;
    if (n > 500) n = 500;
    e.target.value = n;
    adjustRowsCount(n);
  });

  // Row Manipulation
  document.getElementById('btnAddRow').addEventListener('click', addSingleRow);
  document.getElementById('btnRemoveRow').addEventListener('click', removeLastRow);
  document.getElementById('btnResetRows').addEventListener('click', () => {
    const count = parseInt(document.getElementById('countVar').value) || 5;
    buildInitialRows(count);
  });

  // Actions
  document.getElementById('btnGenerate').addEventListener('click', generateChemicalData);
  document.getElementById('btnExportExcel').addEventListener('click', exportToExcel);

  // Import Excel
  const fileInput = document.getElementById('excelFileInput');
  document.getElementById('btnImportExcel').addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });
  fileInput.addEventListener('change', handleExcelImport);
}

function switchTab(tabId) {
  document.querySelectorAll('.chem-tab-btn').forEach(b => b.classList.remove('active'));
  document.querySelectorAll('.chem-tab-content').forEach(c => c.classList.remove('active'));

  const activeBtn = document.querySelector(`.chem-tab-btn[data-tab="${tabId}"]`);
  const activeContent = document.getElementById(tabId);

  if (activeBtn) activeBtn.classList.add('active');
  if (activeContent) activeContent.classList.add('active');
}

function setStatus(text) {
  const footer = document.getElementById('statusFooter');
  if (footer) footer.textContent = text;
}

function makeOptions(optionsList, selectedVal) {
  return optionsList.map(opt => `<option value="${opt}" ${opt === selectedVal ? 'selected' : ''}>${opt}</option>`).join('');
}

function buildInitialRows(count) {
  const tbody = document.getElementById('inputTableBody');
  tbody.innerHTML = '';

  for (let i = 1; i <= count; i++) {
    const tr = createRowElement({
      bh_id: `BH-${String(i).padStart(2, '0')}`,
      s1_depth: "1.5",
      s1_type: "CL",
      s1_condition: "Moist",
      s1_salinity: "Normal",
      s2_depth: "4.5",
      s2_type: "CL",
      s2_condition: "Moist",
      s2_salinity: "Normal",
      gw_depth: "8.0",
      gw_condition: "Normal",
      gw_salinity: "Normal",
      white_residue: "None"
    }, i);
    tbody.appendChild(tr);
  }
  document.getElementById('countVar').value = count;
}

function adjustRowsCount(targetCount) {
  const tbody = document.getElementById('inputTableBody');
  const currentCount = tbody.children.length;

  if (targetCount > currentCount) {
    for (let i = currentCount + 1; i <= targetCount; i++) {
      const tr = createRowElement({
        bh_id: `BH-${String(i).padStart(2, '0')}`,
        s1_depth: "1.5",
        s1_type: "CL",
        s1_condition: "Moist",
        s1_salinity: "Normal",
        s2_depth: "4.5",
        s2_type: "CL",
        s2_condition: "Moist",
        s2_salinity: "Normal",
        gw_depth: "8.0",
        gw_condition: "Normal",
        gw_salinity: "Normal",
        white_residue: "None"
      }, i);
      tbody.appendChild(tr);
    }
  } else if (targetCount < currentCount) {
    while (tbody.children.length > targetCount) {
      tbody.removeChild(tbody.lastChild);
    }
  }
}

function addSingleRow() {
  const tbody = document.getElementById('inputTableBody');
  const nextIdx = tbody.children.length + 1;
  const tr = createRowElement({
    bh_id: `BH-${String(nextIdx).padStart(2, '0')}`,
    s1_depth: "1.5",
    s1_type: "CL",
    s1_condition: "Moist",
    s1_salinity: "Normal",
    s2_depth: "4.5",
    s2_type: "CL",
    s2_condition: "Moist",
    s2_salinity: "Normal",
    gw_depth: "8.0",
    gw_condition: "Normal",
    gw_salinity: "Normal",
    white_residue: "None"
  }, nextIdx);
  tbody.appendChild(tr);
  document.getElementById('countVar').value = nextIdx;
}

function removeLastRow() {
  const tbody = document.getElementById('inputTableBody');
  if (tbody.children.length > 1) {
    tbody.removeChild(tbody.lastChild);
    document.getElementById('countVar').value = tbody.children.length;
  }
}

function createRowElement(data, index) {
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="tk-grid-input bh-id" value="${data.bh_id || `BH-${String(index).padStart(2, '0')}`}" style="text-align: center; font-weight: bold;"></td>
    <td><input type="text" class="tk-grid-input s1-depth" value="${data.s1_depth || '1.5'}" style="text-align: center;"></td>
    <td><select class="tk-grid-select s1-type">${makeOptions(SOIL_TYPES, data.s1_type || 'CL')}</select></td>
    <td><select class="tk-grid-select s1-condition">${makeOptions(CONDITIONS, data.s1_condition || 'Moist')}</select></td>
    <td><select class="tk-grid-select s1-salinity">${makeOptions(SALINITY_CONTROLS, data.s1_salinity || 'Normal')}</select></td>
    <td><input type="text" class="tk-grid-input s2-depth" value="${data.s2_depth || '4.5'}" style="text-align: center;"></td>
    <td><select class="tk-grid-select s2-type">${makeOptions(SOIL_TYPES, data.s2_type || 'CL')}</select></td>
    <td><select class="tk-grid-select s2-condition">${makeOptions(CONDITIONS, data.s2_condition || 'Moist')}</select></td>
    <td><select class="tk-grid-select s2-salinity">${makeOptions(SALINITY_CONTROLS, data.s2_salinity || 'Normal')}</select></td>
    <td><input type="text" class="tk-grid-input gw-depth" value="${data.gw_depth || '8.0'}" style="text-align: center;"></td>
    <td><select class="tk-grid-select gw-condition">${makeOptions(GW_CONDITIONS, data.gw_condition || 'Normal')}</select></td>
    <td><select class="tk-grid-select gw-salinity">${makeOptions(GW_SALINITY_CONTROLS, data.gw_salinity || 'Normal')}</select></td>
    <td><select class="tk-grid-select white-residue">${makeOptions(RESIDUE_OPTIONS, data.white_residue || 'None')}</select></td>
  `;
  return tr;
}

function collectBoreholesInput() {
  const rows = [];
  const tbody = document.getElementById('inputTableBody');

  for (let i = 0; i < tbody.children.length; i++) {
    const tr = tbody.children[i];
    rows.push({
      bh_id: tr.querySelector('.bh-id').value.trim(),
      s1_depth: tr.querySelector('.s1-depth').value.trim(),
      s1_type: tr.querySelector('.s1-type').value,
      s1_condition: tr.querySelector('.s1-condition').value,
      s1_salinity: tr.querySelector('.s1-salinity').value,
      s2_depth: tr.querySelector('.s2-depth').value.trim(),
      s2_type: tr.querySelector('.s2-type').value,
      s2_condition: tr.querySelector('.s2-condition').value,
      s2_salinity: tr.querySelector('.s2-salinity').value,
      gw_depth: tr.querySelector('.gw-depth').value.trim(),
      gw_condition: tr.querySelector('.gw-condition').value,
      gw_salinity: tr.querySelector('.gw-salinity').value,
      white_residue: tr.querySelector('.white-residue').value
    });
  }
  return rows;
}

async function generateChemicalData() {
  const boreholes = collectBoreholesInput();
  if (boreholes.length === 0) {
    alert("Input Required\n\nPlease add at least one borehole.");
    return;
  }

  const project = document.getElementById('projectVar').value.trim();
  const location = document.getElementById('locationVar').value.trim();
  const environment = document.getElementById('envVar').value;
  const seed = document.getElementById('seedVar').value.trim();

  const btnGen = document.getElementById('btnGenerate');
  btnGen.disabled = true;
  setStatus("Generating synthetic chemical test data...");

  try {
    const res = await fetch('/api/chemical/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        project,
        location,
        environment,
        seed,
        boreholes
      })
    });

    const data = await res.json();
    btnGen.disabled = false;

    if (!res.ok || !data.success) {
      alert("Validation Error:\n\n" + (data.error || "Failed to generate chemical test results."));
      setStatus("Generation error: " + (data.error || "Failed"));
      return;
    }

    generatedData = data;
    renderSoilResults(data.soil_results || []);
    renderGwResults(data.gw_results || []);
    renderSummaryResults(data.summary_rows || []);

    setStatus(data.status_text || `Generated ${boreholes.length} boreholes successfully.`);
    switchTab('tab-soil');

  } catch (err) {
    btnGen.disabled = false;
    alert("Generation Error:\n\n" + err.message);
    setStatus("Error: " + err.message);
  }
}

function renderSoilResults(rows) {
  const tbody = document.getElementById('soilTableBody');
  tbody.innerHTML = '';

  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="15" style="text-align: center; color: #888; padding: 20px;">No soil results.</td></tr>';
    return;
  }

  rows.forEach(r => {
    const tr = document.createElement('tr');
    const isValid = r.validation === "PASS";
    tr.innerHTML = `
      <td style="text-align: center; font-weight: bold;">${r.bh_id}</td>
      <td style="text-align: center; font-weight: 500; color: #1e40af;">${r.sample}</td>
      <td style="text-align: center;">${r.depth}</td>
      <td style="text-align: center;">${r.zone}</td>
      <td style="text-align: center;">${r.soil_type}</td>
      <td style="text-align: center;">${r.condition}</td>
      <td style="text-align: center;">${r.salinity_control}</td>
      <td style="text-align: center;">${r.environment}</td>
      <td style="text-align: center; font-weight: bold;">${r.ph}</td>
      <td style="text-align: center; font-weight: bold;">${r.chloride}</td>
      <td style="text-align: center; font-weight: bold;">${r.sulphate}</td>
      <td style="text-align: center;">${r.organic_matter}</td>
      <td style="text-align: center;">${r.white_residue}</td>
      <td style="text-align: center;">${r.chemical_dominance}</td>
      <td style="text-align: center;"><span class="${isValid ? 'badge-pass' : 'badge-check'}">${r.validation}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function renderGwResults(rows) {
  const tbody = document.getElementById('gwTableBody');
  tbody.innerHTML = '';

  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align: center; color: #888; padding: 20px;">No groundwater results.</td></tr>';
    return;
  }

  rows.forEach(r => {
    const tr = document.createElement('tr');
    const isValid = r.validation === "PASS";
    tr.innerHTML = `
      <td style="text-align: center; font-weight: bold;">${r.bh_id}</td>
      <td style="text-align: center; font-weight: 500; color: #0284c7;">${r.sample}</td>
      <td style="text-align: center;">${r.depth}</td>
      <td style="text-align: center;">${r.environment}</td>
      <td style="text-align: center;">${r.condition}</td>
      <td style="text-align: center; font-weight: bold;">${r.ph}</td>
      <td style="text-align: center; font-weight: bold;">${r.chloride}</td>
      <td style="text-align: center; font-weight: bold;">${r.sulphate}</td>
      <td style="text-align: center;">${r.salinity_condition}</td>
      <td style="text-align: center;">${r.chemical_dominance}</td>
      <td style="text-align: center;"><span class="${isValid ? 'badge-pass' : 'badge-check'}">${r.validation}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

function renderSummaryResults(rows) {
  const tbody = document.getElementById('summaryTableBody');
  tbody.innerHTML = '';

  if (!rows || rows.length === 0) {
    tbody.innerHTML = '<tr><td colspan="24" style="text-align: center; color: #888; padding: 20px;">No summary results.</td></tr>';
    return;
  }

  rows.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center; font-weight: bold;">${r.bh_id}</td>
      <td style="text-align: center;">${r.environment}</td>
      <td style="text-align: center;">${r.s1_depth}</td>
      <td style="text-align: center;">${r.s1_zone}</td>
      <td style="text-align: center;">${r.s1_type}</td>
      <td style="text-align: center;">${r.s1_condition}</td>
      <td style="text-align: center;">${r.s1_salinity}</td>
      <td style="text-align: center;">${r.s2_depth}</td>
      <td style="text-align: center;">${r.s2_zone}</td>
      <td style="text-align: center;">${r.s2_type}</td>
      <td style="text-align: center;">${r.s2_condition}</td>
      <td style="text-align: center;">${r.s2_salinity}</td>
      <td style="text-align: center;">${r.gw_depth}</td>
      <td style="text-align: center;">${r.gw_condition}</td>
      <td style="text-align: center;">${r.gw_salinity}</td>
      <td style="text-align: center;">${r.white_residue}</td>
      <td style="text-align: center;">${r.chemical_dominance}</td>
      <td style="text-align: center; font-weight: bold;">${r.s1_cl}</td>
      <td style="text-align: center; font-weight: bold;">${r.s2_cl}</td>
      <td style="text-align: center; font-weight: bold; color: #0284c7;">${r.gw_cl}</td>
      <td style="text-align: center; font-weight: bold;">${r.s1_so4}</td>
      <td style="text-align: center; font-weight: bold;">${r.s2_so4}</td>
      <td style="text-align: center; font-weight: bold; color: #0284c7;">${r.gw_so4}</td>
      <td style="text-align: center;"><span class="badge-pass">${r.overall_check || 'PASS'}</span></td>
    `;
    tbody.appendChild(tr);
  });
}

async function exportToExcel() {
  if (!generatedData || !generatedData.soil_results || generatedData.soil_results.length === 0) {
    alert("No Data\n\nPlease generate data first before exporting.");
    return;
  }

  const project = document.getElementById('projectVar').value.trim() || "Chemical_Report";
  const location = document.getElementById('locationVar').value.trim();
  const environment = document.getElementById('envVar').value;
  const seed = document.getElementById('seedVar').value.trim();

  try {
    const res = await fetch('/api/chemical/export_excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        project,
        location,
        environment,
        seed,
        soil_results: generatedData.soil_results,
        gw_results: generatedData.gw_results,
        summary_rows: generatedData.summary_rows
      })
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      alert("Export Error:\n\n" + (data.error || "Failed to export Excel."));
      return;
    }

    // Trigger download
    const dlLink = document.createElement('a');
    dlLink.href = data.download_url;
    dlLink.download = data.filename;
    document.body.appendChild(dlLink);
    dlLink.click();
    document.body.removeChild(dlLink);

    setStatus("Workbook saved successfully: " + data.filename);
    alert(`Export Complete\n\nWorkbook saved successfully:\n${data.filename}`);

  } catch (err) {
    alert("Export Error:\n\n" + err.message);
  }
}

async function handleExcelImport(e) {
  const file = e.target.files[0];
  if (!file) return;

  const formData = new FormData();
  formData.append('file', file);

  setStatus("Importing borehole data from Excel...");

  try {
    const res = await fetch('/api/chemical/import_excel', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      alert("Excel Import Error:\n\n" + (data.error || "Failed to import Excel file."));
      setStatus("Import failed.");
      return;
    }

    // Populate Settings if present
    if (data.project) document.getElementById('projectVar').value = data.project;
    if (data.location) document.getElementById('locationVar').value = data.location;
    if (data.environment) document.getElementById('envVar').value = data.environment;
    if (data.seed) document.getElementById('seedVar').value = data.seed;

    // Populate Table
    const tbody = document.getElementById('inputTableBody');
    tbody.innerHTML = '';

    (data.boreholes || []).forEach((b, idx) => {
      const tr = createRowElement(b, idx + 1);
      tbody.appendChild(tr);
    });

    document.getElementById('countVar').value = (data.boreholes || []).length;

    setStatus(`Imported ${data.count} borehole(s) from Excel. Review inputs, then click Generate.`);
    switchTab('tab-input');

    alert(`Import Complete\n\nSuccessfully imported ${data.count} borehole(s).\n\nPlease review the Borehole Input tab and click Generate.`);

  } catch (err) {
    alert("Excel Import Error:\n\n" + err.message);
    setStatus("Import error: " + err.message);
  }
}
