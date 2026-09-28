// Engineering Strength Parameter Generator Controller v11 (Formula-Safe All-Sheets)
// 100% Faithful to engineering_strength_parameter_generator_v11_formula_safe_all_sheets.py

let selectedFile = null;
let currentPreviewData = null;
let activeSheetFilter = "all";

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
});

function bindEvents() {
  const fileInput = document.getElementById('paramFileInput');
  const btnBrowse = document.getElementById('btnBrowseInput');
  const btnSaveAs = document.getElementById('btnSaveAs');
  const btnGen = document.getElementById('btnGenerate');
  const btnClr = document.getElementById('btnClear');

  btnBrowse.addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  fileInput.addEventListener('change', handleFileSelected);

  btnSaveAs.addEventListener('click', () => {
    const currentName = document.getElementById('outputFileName').value || "Synthetic_Strength_Parameters_v4.xlsx";
    const newName = prompt("Save Output Excel as:", currentName);
    if (newName && newName.trim()) {
      document.getElementById('outputFileName').value = newName.trim();
    }
  });

  btnGen.addEventListener('click', generateAndExport);
  btnClr.addEventListener('click', clearAll);
}

function log(text) {
  const logBox = document.getElementById('statusLog');
  logBox.textContent += "\n" + text;
  logBox.scrollTop = logBox.scrollHeight;
}

async function handleFileSelected(e) {
  const file = e.target.files[0];
  if (!file) return;

  selectedFile = file;
  document.getElementById('inputFilePath').value = file.name;

  const formData = new FormData();
  formData.append('file', file);

  try {
    const res = await fetch('/api/parameter/inspect_file', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      alert("Excel Error:\n\n" + (data.error || "Failed to inspect Excel file."));
      return;
    }

    if (data.default_output) {
      document.getElementById('outputFileName').value = data.default_output;
    }

    log("Input loaded: " + file.name);
    log(`Worksheets found (${(data.sheets || []).length}): ${(data.sheets || []).join(', ')}`);
    log("All worksheets will be processed automatically.");

  } catch (err) {
    alert("Excel Error:\n\n" + err.message);
  }
}

async function generateAndExport() {
  if (!selectedFile) {
    alert("Input Required\n\nPlease select an input Excel file.");
    return;
  }

  const outputName = document.getElementById('outputFileName').value.trim();
  if (!outputName) {
    alert("Output Required\n\nPlease select an output Excel file.");
    return;
  }

  const variabilityStr = document.getElementById('variabilityVar').value.trim();
  const variability = parseFloat(variabilityStr);
  if (isNaN(variability)) {
    alert("Invalid Variability\n\nVariability must be a number, for example 10.");
    return;
  }
  if (variability < 0 || variability > 30) {
    alert("Invalid Variability\n\nPlease enter a value between 0 and 30%.");
    return;
  }

  const seedStr = document.getElementById('seedVar').value.trim();
  if (seedStr && isNaN(parseInt(seedStr))) {
    alert("Invalid Seed\n\nRandom Seed must be an integer or left blank.");
    return;
  }

  const btnGen = document.getElementById('btnGenerate');
  btnGen.disabled = true;

  log("Generating synthetic strength parameters for ALL worksheets...");
  log(`Variability: ${variability.toFixed(1)}%`);
  log(seedStr ? `Random seed: ${seedStr}` : "Random seed: automatic");

  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('variability', variabilityStr);
  formData.append('seed', seedStr);
  formData.append('output_name', outputName);

  try {
    const res = await fetch('/api/parameter/generate', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    btnGen.disabled = false;

    if (!res.ok || !data.success) {
      log("ERROR: " + (data.error || "Generation error"));
      alert("Generation Error\n\n" + (data.error || "Failed to generate parameters."));
      return;
    }

    log(`Total rows processed: ${data.count}`);

    if (data.processed_sheets && data.processed_sheets.length > 0) {
      log("Processed worksheets:");
      data.processed_sheets.forEach(([sh, count]) => {
        log(`  • ${sh}: ${count} rows`);
      });
    }

    if (data.skipped_sheets && data.skipped_sheets.length > 0) {
      log("Skipped worksheets (no usable Depth + SPT N data): " + data.skipped_sheets.join(', '));
    }

    log("Output saved: " + data.excel_filename);

    showResults(data);

    const procCount = data.processed_sheets ? data.processed_sheets.length : 0;
    const skipCount = data.skipped_sheets ? data.skipped_sheets.length : 0;
    alert(`Completed\n\nAll worksheet processing completed.\n\nWorksheets processed: ${procCount}\nTotal rows processed: ${data.count}\n\nSheets without usable numeric Depth + SPT N: ${skipCount}`);

  } catch (err) {
    btnGen.disabled = false;
    log("ERROR: " + err.message);
    alert("Generation Error\n\n" + err.message);
  }
}

function showResults(data) {
  currentPreviewData = data;
  const resSec = document.getElementById('resultsSection');
  resSec.style.display = 'block';

  const procCount = data.processed_sheets ? data.processed_sheets.length : 0;
  document.getElementById('resultStats').textContent = 
    `Generated ${data.count} rows of synthetic strength parameters across ${procCount} worksheets into ${data.excel_filename}.`;

  const btnDl = document.getElementById('btnDownloadResult');
  btnDl.href = data.download_url;
  btnDl.setAttribute('download', data.excel_filename);

  renderSheetTabs(data);
  renderTableRows("all");
}

function renderSheetTabs(data) {
  const container = document.getElementById('sheetTabsContainer');
  container.innerHTML = '';

  const allBtn = document.createElement('button');
  allBtn.type = 'button';
  allBtn.className = 'sheet-tab-btn active';
  allBtn.textContent = `All Sheets (${data.count})`;
  allBtn.addEventListener('click', () => {
    setActiveTab(allBtn, "all");
  });
  container.appendChild(allBtn);

  if (data.processed_sheets && data.processed_sheets.length > 1) {
    data.processed_sheets.forEach(([sheetName, count]) => {
      const tabBtn = document.createElement('button');
      tabBtn.type = 'button';
      tabBtn.className = 'sheet-tab-btn';
      tabBtn.textContent = `${sheetName} (${count})`;
      tabBtn.addEventListener('click', () => {
        setActiveTab(tabBtn, sheetName);
      });
      container.appendChild(tabBtn);
    });
  }
}

function setActiveTab(btn, sheetKey) {
  document.querySelectorAll('.sheet-tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  activeSheetFilter = sheetKey;
  document.getElementById('currentSheetLabel').textContent = sheetKey === "all" ? "All Worksheets" : sheetKey;
  renderTableRows(sheetKey);
}

function renderTableRows(sheetKey) {
  if (!currentPreviewData) return;

  const tbody = document.getElementById('previewTableBody');
  tbody.innerHTML = '';

  let rowsToRender = [];
  if (sheetKey === "all") {
    rowsToRender = currentPreviewData.rows || [];
  } else if (currentPreviewData.preview_by_sheet && currentPreviewData.preview_by_sheet[sheetKey]) {
    rowsToRender = currentPreviewData.preview_by_sheet[sheetKey];
  }

  rowsToRender.forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;"><span class="sheet-badge">${r.sheet || ''}</span></td>
      <td style="text-align: center;">${r.row}</td>
      <td style="text-align: center;">${r.depth}</td>
      <td style="text-align: center; font-weight: bold;">${r.field_n}</td>
      <td style="text-align: center;">${r.density}</td>
      <td style="text-align: center;">${r.obd}</td>
      <td style="text-align: center;">${r.pa}</td>
      <td style="text-align: center;">${r.cn_raw}</td>
      <td style="text-align: center;">${r.cn}</td>
      <td style="text-align: center;">${r.n1}</td>
      <td style="text-align: center; font-weight: bold; color: #1e40af;">${r.n1_prime}</td>
      <td style="text-align: center;">${r.pi}</td>
      <td style="text-align: center;">${r.li}</td>
      <td style="text-align: center;">${r.cf}</td>
      <td style="text-align: center;">${r.soil}</td>
      <td style="text-align: center;">${r.k_pi}</td>
      <td style="text-align: center;">${r.f_li}</td>
      <td style="text-align: center;">${r.f_c}</td>
      <td style="text-align: center; font-weight: bold;">${r.su}</td>
      <td style="text-align: center;">${r.ucs}</td>
      <td style="text-align: center;">${r.c_dr}</td>
      <td style="text-align: center;">${r.phi_dr}</td>
      <td style="text-align: center;">${r.cu_uu}</td>
      <td style="text-align: center;">${r.phi_uu}</td>
    `;
    tbody.appendChild(tr);
  });
}

function clearAll() {
  selectedFile = null;
  currentPreviewData = null;
  activeSheetFilter = "all";

  document.getElementById('paramFileInput').value = '';
  document.getElementById('inputFilePath').value = '';
  document.getElementById('outputFileName').value = '';
  document.getElementById('variabilityVar').value = '10';
  document.getElementById('seedVar').value = '';
  document.getElementById('resultsSection').style.display = 'none';

  const logBox = document.getElementById('statusLog');
  logBox.textContent = 'Ready. Select an Excel file. All worksheets will be processed automatically. Layout on every sheet: A=Depth, G=SPT N, I=Gravel, J=Sand, K=Silt, L=Clay, M=LL, N=PL, S=FMC. Input starts at row 10; output starts at AU.';
}
