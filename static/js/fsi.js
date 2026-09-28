// Free Swell Index (FSI) Desktop-Matched Controller
// 100% faithful to FSI_Data_Generator.py

let soilOptions = [
  "Auto / Random Soil Type",
  "Sand / sandy soil",
  "Silty soil",
  "CL / low plasticity clay",
  "CI / medium plasticity clay",
  "CH / high plasticity clay",
  "Expansive clay",
  "Black cotton soil",
  "Highly expansive / bentonitic soil"
];

let sampleTypes = ["UDS", "SPT"];
let depthRows = [];
let generatedRows = [];

document.addEventListener('DOMContentLoaded', async () => {
  try {
    const res = await fetch('/api/fsi/options');
    const data = await res.json();
    if (data.soil_options) soilOptions = data.soil_options;
    if (data.sample_types) sampleTypes = data.sample_types;
  } catch (e) {
    console.warn("Using default FSI options");
  }

  // Initial single row matching original FSI desktop app
  addDepthRow();
  bindEvents();
});

function bindEvents() {
  document.getElementById('btnAddDepth').addEventListener('click', addDepthRow);
  document.getElementById('btnClearDepths').addEventListener('click', clearDepths);
  document.getElementById('btnGenerate').addEventListener('click', generateFSI);
  document.getElementById('btnClearGenerated').addEventListener('click', clearGenerated);
  document.getElementById('btnExportExcel').addEventListener('click', exportExcel);
}

function addDepthRow() {
  const num = depthRows.length + 1;
  const defaultDepth = (1.50 + (num - 1) * 1.50).toFixed(2);
  const rowId = 'row_' + Date.now() + '_' + Math.random().toString(36).substr(2, 5);

  const row = {
    id: rowId,
    depth: defaultDepth,
    sample_type: "UDS",
    soil: "Auto / Random Soil Type"
  };

  depthRows.push(row);
  renderDepthRows();
}

function renderDepthRows() {
  const tbody = document.getElementById('depthTableBody');
  tbody.innerHTML = '';

  depthRows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.id = row.id;

    let soilOpts = '';
    soilOptions.forEach(s => {
      soilOpts += `<option value="${s}" ${s === row.soil ? 'selected' : ''}>${s}</option>`;
    });

    let sampleOpts = '';
    sampleTypes.forEach(t => {
      sampleOpts += `<option value="${t}" ${t === row.sample_type ? 'selected' : ''}>${t}</option>`;
    });

    tr.innerHTML = `
      <td style="font-weight: normal; text-align: center;">${idx + 1}</td>
      <td style="text-align: center;">
        <input type="text" class="tk-table-entry" style="width: 120px;" value="${row.depth}" onchange="updateRow('${row.id}', 'depth', this.value)">
      </td>
      <td style="text-align: center;">
        <select class="tk-table-select" style="width: 110px;" onchange="updateRow('${row.id}', 'sample_type', this.value)">
          ${sampleOpts}
        </select>
      </td>
      <td style="text-align: left; padding-left: 10px;">
        <select class="fsi-soil-select" onchange="updateRow('${row.id}', 'soil', this.value)">
          ${soilOpts}
        </select>
      </td>
      <td style="text-align: center;">
        <button class="tk-btn-remove" onclick="removeRow('${row.id}')">Remove</button>
      </td>
    `;
    tbody.appendChild(tr);
  });
}

function updateRow(id, field, val) {
  const row = depthRows.find(r => r.id === id);
  if (row) row[field] = val;
}

function removeRow(id) {
  if (depthRows.length <= 1) {
    depthRows = [];
    addDepthRow();
    return;
  }
  depthRows = depthRows.filter(r => r.id !== id);
  renderDepthRows();
}

function clearDepths() {
  depthRows = [];
  addDepthRow();
  document.getElementById('statusVar').textContent = "Depths cleared.";
}

async function generateFSI() {
  const bhId = document.getElementById('bh_id').value.trim();
  if (!bhId) {
    alert("Input Error\n\nBH ID cannot be empty.");
    return;
  }

  if (!depthRows || depthRows.length === 0) {
    alert("Input Error\n\nAdd at least one sample depth.");
    return;
  }

  const usedDepths = new Set();
  for (let i = 0; i < depthRows.length; i++) {
    const r = depthRows[i];
    const dVal = parseFloat(r.depth);
    if (isNaN(dVal)) {
      alert(`Input Error\n\nInvalid Sample Depth in row ${i + 1}.`);
      return;
    }
    if (dVal < 0) {
      alert(`Input Error\n\nSample Depth in row ${i + 1} cannot be negative.`);
      return;
    }
    if (usedDepths.has(dVal)) {
      alert(`Input Error\n\nDuplicate Sample Depth ${dVal} m.`);
      return;
    }
    usedDepths.add(dVal);
  }

  const payload = {
    bh_id: bhId,
    rows: depthRows
  };

  const statusEl = document.getElementById('statusVar');
  statusEl.textContent = "Generating...";

  try {
    const res = await fetch('/api/fsi/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      alert("Generation Error\n\n" + (data.error || "Generation failed"));
      statusEl.textContent = "Ready";
      return;
    }

    generatedRows = data.results;
    renderGeneratedTable();
    statusEl.textContent = `Generated ${generatedRows.length} sample depth(s).`;
  } catch (e) {
    alert("Generation Error\n\n" + e.message);
    statusEl.textContent = "Ready";
  }
}

function renderGeneratedTable() {
  const tbody = document.getElementById('resultsTableBody');
  tbody.innerHTML = '';

  generatedRows.forEach(item => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">${item["SR No"]}</td>
      <td style="text-align: center;">${item["BH ID"]}</td>
      <td style="text-align: center;">${item["Sample ID"]}</td>
      <td style="text-align: center;">${item["Sample Type"]}</td>
      <td style="text-align: center;">${item["Depth"]}</td>
      <td style="text-align: left; padding-left: 8px;">${item["Soil Description"]}</td>
      <td style="text-align: center;">${item["Vk"]}</td>
      <td style="text-align: center;">${item["Vd"]}</td>
      <td style="text-align: center;">${item["Difference"]}</td>
      <td style="text-align: center; font-weight: bold;">${item["FSI"].toFixed(2)}</td>
    `;
    tbody.appendChild(tr);
  });
}

function clearGenerated() {
  generatedRows = [];
  document.getElementById('resultsTableBody').innerHTML = '';
  document.getElementById('statusVar').textContent = "Generated data cleared.";
}

async function triggerBrowserFileDownload(data, fallbackFilename, mimeType = "application/zip") {
  const filename = data.filename || data.zip_filename || data.excel_filename || fallbackFilename;
  const b64Data = data.file_base64 || data.zip_base64 || data.excel_base64;
  
  if (b64Data) {
    try {
      const byteCharacters = atob(b64Data);
      const byteNumbers = new Array(byteCharacters.length);
      for (let i = 0; i < byteCharacters.length; i++) {
        byteNumbers[i] = byteCharacters.charCodeAt(i);
      }
      const byteArray = new Uint8Array(byteNumbers);
      const blob = new Blob([byteArray], { type: mimeType });

      // 1. If modern browser supports showSaveFilePicker (Chrome/Edge), open native "Save As" Dialog so user can choose ANY folder!
      if (window.showSaveFilePicker) {
        try {
          const ext = filename.includes('.') ? '.' + filename.split('.').pop() : (mimeType.includes('zip') ? '.zip' : '.xlsx');
          const handle = await window.showSaveFilePicker({
            suggestedName: filename,
            types: [{
              description: ext.toUpperCase() + ' File',
              accept: { [mimeType]: [ext] }
            }]
          });
          const writable = await handle.createWritable();
          await writable.write(blob);
          await writable.close();
          return true;
        } catch (pickerErr) {
          if (pickerErr.name === 'AbortError') {
            // User cancelled folder selection
            return false;
          }
          console.warn("SaveFilePicker fallback", pickerErr);
        }
      }

      // 2. Standard direct browser download fallback
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
      return true;
    } catch (err) {
      console.warn("Base64 blob conversion fallback", err);
    }
  }

  // URL Fallback
  const url = data.download_zip_url || data.download_url || data.url;
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
  return true;
}

async function exportExcel() {
  const bhId = document.getElementById('bh_id').value.trim() || 'RBH-12';
  const statusEl = document.getElementById('statusVar');
  statusEl.textContent = "Exporting FSI Excel ZIP Package... Please choose destination folder.";

  // If no generated data yet, generate first
  if (!generatedRows || generatedRows.length === 0) {
    await generateFSI();
    if (!generatedRows || generatedRows.length === 0) {
      statusEl.textContent = "Ready";
      return;
    }
  }

  try {
    const payload = { bh_id: bhId, results: generatedRows, rows: depthRows };
    const res = await fetch('/api/fsi/export_excel', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(payload)
    });
    const data = await res.json();

    if (!res.ok || data.error) {
      throw new Error(data.error || "Excel export failed on server.");
    }

    const fallbackZip = `FSI_Package_${bhId}.zip`;
    const saved = await triggerBrowserFileDownload(
      {
        filename: data.zip_filename || fallbackZip,
        file_base64: data.zip_base64 || data.file_base64,
        download_zip_url: data.download_zip_url
      },
      fallbackZip,
      "application/zip"
    );

    if (saved !== false) {
      alert(
        `Generation Completed\n\n` +
        `FSI Report generation completed successfully.\n\n` +
        `Samples exported: ${data.count}\n` +
        `ZIP Package: ${data.zip_filename || fallbackZip}\n` +
        `Excel File: ${data.excel_filename}\n\n` +
        `Downloaded to your computer's selected folder.`
      );
      statusEl.textContent = `ZIP Package saved: ${data.zip_filename || fallbackZip}`;
    } else {
      statusEl.textContent = "Export cancelled by user.";
    }
  } catch (e) {
    alert("Excel Export Error\n\n" + e.message);
    statusEl.textContent = "Ready";
  }
}
