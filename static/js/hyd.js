// Hydrometer Synthetic Data Generator Controller v1.9
// 100% Faithful to hydrometer_synthetic_generator_v1_9_auto_clay_PSD_v2_9_plot_GSD_Fines_fixed.py

const PRESETS = [
  "1 - Coarse Silt Dominant",
  "2 - Normal Silt",
  "3 - Fine Silt Dominant",
  "4 - Clayey Silt",
  "5 - Silty Clay",
  "6 - Broad Natural"
];

let selectedFile = null;
let currentBhId = "BH-01";
let sampleRows = [];

document.addEventListener('DOMContentLoaded', () => {
  bindEvents();
});

function bindEvents() {
  const fileInput = document.getElementById('templateFileInput');
  const btnOpen = document.getElementById('btnOpenTemplate');
  const templateTypeSelect = document.getElementById('templateType');
  const btnGen = document.getElementById('btnGenerate');

  const btnNewTest = document.getElementById('btnNewTest');

  btnOpen.addEventListener('click', () => {
    fileInput.value = '';
    fileInput.click();
  });

  if (btnNewTest) {
    btnNewTest.addEventListener('click', () => {
      window.location.href = '/tests/hyd';
    });
  }

  fileInput.addEventListener('change', (e) => {
    if (e.target.files && e.target.files.length > 0) {
      selectedFile = e.target.files[0];
      document.getElementById('fileLabel').textContent = selectedFile.name;
      loadTemplateData();
    }
  });

  templateTypeSelect.addEventListener('change', () => {
    if (selectedFile) {
      loadTemplateData();
    }
  });

  btnGen.addEventListener('click', generateHYD);
}

function resetAllData() {
  selectedFile = null;
  currentBhId = "BH-01";
  sampleRows = [];
  document.getElementById('templateFileInput').value = '';
  document.getElementById('fileLabel').textContent = 'No template selected';
  document.getElementById('statusText').textContent = 'Open a GSD & HYD template. Fixed settings: T=27°C, Cm=0.0005, Mt=0, X=0.3.';
  renderTable();
}

async function loadTemplateData() {
  if (!selectedFile) return;

  const templateType = document.getElementById('templateType').value;
  const statusEl = document.getElementById('statusText');
  statusEl.textContent = "Loading and reading GSD & HYD template...";

  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('template_type', templateType);

  try {
    const res = await fetch('/api/hyd/parse_template', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    if (!res.ok || !data.success) {
      alert("Template Error:\n" + (data.error || "Failed to parse template"));
      statusEl.textContent = "Failed to load template.";
      return;
    }

    currentBhId = data.bh_id || "BH-01";
    sampleRows = data.rows || [];

    renderTable();
    statusEl.textContent = `${sampleRows.length} sample row(s) loaded. BH ID: ${currentBhId}. Fixed settings: T=27°C, Cm=0.0005, Mt=0, X=0.3.`;

  } catch (err) {
    alert("Template Error:\n" + err.message);
    statusEl.textContent = "Error reading template file.";
  }
}

function renderTable() {
  const tbody = document.getElementById('hydTableBody');
  tbody.innerHTML = '';

  if (sampleRows.length === 0) {
    tbody.innerHTML = `
      <tr id="emptyPlaceholderRow">
        <td colspan="12" style="text-align: center; padding: 40px 10px; color: #666666; font-style: italic;">
          <i class="fas fa-file-excel" style="font-size: 24pt; margin-bottom: 8px; color: #107c41; display: block;"></i>
          No GSD sample rows found in this template.
        </td>
      </tr>
    `;
    return;
  }

  sampleRows.forEach((row, idx) => {
    const tr = document.createElement('tr');
    tr.id = `row_${idx}`;

    let presetOptionsHtml = '';
    PRESETS.forEach(p => {
      const isSel = p === (row.preset || "2 - Normal Silt") ? "selected" : "";
      presetOptionsHtml += `<option value="${p}" ${isSel}>${p}</option>`;
    });

    tr.innerHTML = `
      <td class="hyd-cell-ridge">${row.depth}</td>
      <td class="hyd-cell-ridge">${row.sample_type}</td>
      <td class="hyd-cell-ridge">${row.gravel}</td>
      <td class="hyd-cell-ridge">${row.sand}</td>
      <td class="hyd-cell-ridge">${row.fines}</td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="text" class="tk-table-entry" id="silt_${idx}" value="${row.silt || ''}"
          style="width: 85px;" oninput="onSiltChanged(${idx}, this.value)">
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="text" class="tk-table-entry hyd-cell-readonly" id="clay_${idx}" value="${row.clay || ''}"
          style="width: 85px;" readonly>
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <select class="tk-table-select" id="preset_${idx}" style="width: 175px;" onchange="sampleRows[${idx}].preset = this.value">
          ${presetOptionsHtml}
        </select>
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="checkbox" id="gen_${idx}" ${row.generate ? 'checked' : ''} onchange="toggleHyd(${idx}, this.checked)">
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="text" class="tk-table-entry ${!row.generate ? 'hyd-cell-disabled' : ''}" id="spgr_${idx}"
          value="${row.spgr || ''}" style="width: 95px;" ${!row.generate ? 'disabled' : ''}
          oninput="sampleRows[${idx}].spgr = this.value">
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="text" class="tk-table-entry ${!row.generate ? 'hyd-cell-disabled' : ''}" id="hydro_${idx}"
          value="${row.hydrometer || ''}" style="width: 95px;" ${!row.generate ? 'disabled' : ''}
          oninput="sampleRows[${idx}].hydrometer = this.value">
      </td>
      <td style="padding: 2px 3px; text-align: center;">
        <input type="text" class="tk-table-entry ${!row.generate ? 'hyd-cell-disabled' : ''}" id="cylinder_${idx}"
          value="${row.cylinder || ''}" style="width: 95px;" ${!row.generate ? 'disabled' : ''}
          oninput="sampleRows[${idx}].cylinder = this.value">
      </td>
    `;

    tbody.appendChild(tr);
  });
}

function onSiltChanged(idx, val) {
  sampleRows[idx].silt = val;
  const clayEl = document.getElementById(`clay_${idx}`);
  try {
    const fines = parseFloat(sampleRows[idx].fines);
    const siltText = (val || '').trim();
    if (!siltText || isNaN(parseFloat(siltText))) {
      sampleRows[idx].clay = "";
      clayEl.value = "";
      return;
    }
    const silt = parseFloat(siltText);
    const clay = fines - silt;
    if (clay < 0) {
      sampleRows[idx].clay = "";
      clayEl.value = "";
    } else {
      sampleRows[idx].clay = clay.toFixed(2);
      clayEl.value = clay.toFixed(2);
    }
  } catch (e) {
    sampleRows[idx].clay = "";
    clayEl.value = "";
  }
}

function toggleHyd(idx, isChecked) {
  sampleRows[idx].generate = isChecked;
  const spgrEl = document.getElementById(`spgr_${idx}`);
  const hydroEl = document.getElementById(`hydro_${idx}`);
  const cylEl = document.getElementById(`cylinder_${idx}`);

  if (isChecked) {
    spgrEl.disabled = false;
    spgrEl.classList.remove('hyd-cell-disabled');

    hydroEl.disabled = false;
    hydroEl.classList.remove('hyd-cell-disabled');

    cylEl.disabled = false;
    cylEl.classList.remove('hyd-cell-disabled');
  } else {
    spgrEl.disabled = true;
    spgrEl.classList.add('hyd-cell-disabled');

    hydroEl.disabled = true;
    hydroEl.classList.add('hyd-cell-disabled');

    cylEl.disabled = true;
    cylEl.classList.add('hyd-cell-disabled');
  }
}

async function generateHYD() {
  if (!selectedFile) {
    alert("Generation Error:\nOpen a GSD & HYD template first.");
    return;
  }

  if (sampleRows.length === 0) {
    alert("Generation:\nNo GSD sample rows are loaded.");
    return;
  }

  const wbVal = document.getElementById('wbVar').value.trim();
  const tolVal = document.getElementById('toleranceVar').value.trim();
  const seedVal = document.getElementById('seedVar').value.trim();
  const templateType = document.getElementById('templateType').value;

  if (isNaN(parseFloat(wbVal)) || parseFloat(wbVal) <= 0) {
    alert("Generation Error:\nWb (g) must be a valid positive number.");
    return;
  }
  if (isNaN(parseFloat(tolVal)) || parseFloat(tolVal) < 0) {
    alert("Generation Error:\nSilt/Clay allowed error must be a valid number.");
    return;
  }

  // Pre-validate any rows marked for HYD generation
  const tolerance = parseFloat(tolVal);
  for (let i = 0; i < sampleRows.length; i++) {
    const r = sampleRows[i];
    if (r.generate) {
      const gsdRow = r.gsd_row;
      const fines = parseFloat(r.fines);
      const silt = parseFloat(r.silt);

      if (isNaN(silt)) {
        alert(`Generation Error:\nGSD row ${gsdRow} (Depth ${r.depth} m): Target Silt % is blank or non-numeric.`);
        return;
      }
      if (silt < 0) {
        alert(`Generation Error:\nGSD row ${gsdRow}: Target Silt % cannot be negative.`);
        return;
      }
      if (silt > fines + tolerance) {
        alert(`Generation Error:\nGSD row ${gsdRow}: Target Silt = ${silt.toFixed(2)}% exceeds GSD Fines = ${fines.toFixed(2)}%.`);
        return;
      }
      const clay = fines - silt;
      if (clay < 0) {
        alert(`Generation Error:\nGSD row ${gsdRow}: Calculated Clay (${clay.toFixed(2)}%) is negative.`);
        return;
      }
      const spgr = parseFloat(r.spgr);
      if (isNaN(spgr) || spgr <= 1.0) {
        alert(`Generation Error:\nGSD row ${gsdRow}: SPGR must be greater than 1.`);
        return;
      }
    }
  }

  const statusEl = document.getElementById('statusText');
  const btnGen = document.getElementById('btnGenerate');
  btnGen.disabled = true;
  statusEl.textContent = "Processing and generating HYD data & report PSD plots... Please wait.";

  const config = {
    template_type: templateType,
    wb: wbVal,
    tolerance: tolVal,
    seed: seedVal,
    bh_id: currentBhId,
    rows: sampleRows
  };

  const formData = new FormData();
  formData.append('file', selectedFile);
  formData.append('config', JSON.stringify(config));

  try {
    const res = await fetch('/api/hyd/generate', {
      method: 'POST',
      body: formData
    });

    const data = await res.json();
    btnGen.disabled = false;

    if (!res.ok || !data.success) {
      alert("Generation Error:\n\n" + (data.error || "Generation stopped because of an error."));
      statusEl.textContent = "Generation stopped because of an error.";
      return;
    }

    statusEl.textContent = `Generation completed successfully for ${data.hyd_generated_samples} HYD sample(s) and ${data.total_samples} PSD plot(s).`;

    // Auto-trigger single ZIP download containing Excel, PSD Summary CSV, and all PSD Plots
    if (data.zip_url) {
      const autoLink = document.createElement('a');
      autoLink.href = data.zip_url + '?download=1';
      autoLink.setAttribute('download', data.zip_filename);
      document.body.appendChild(autoLink);
      autoLink.click();
      document.body.removeChild(autoLink);
    }

    // Show completion alert matching other test modules
    alert(
      `Generation Complete\n\n` +
      `HYD data generated successfully.\n\n` +
      `Borehole ID: ${data.bh_id}\n` +
      `Total GSD Samples: ${data.total_samples}\n` +
      `HYD Samples Generated: ${data.hyd_generated_samples}\n\n` +
      `Excel Workbook:\n${data.excel_filename}\n\n` +
      `PSD Summary:\n${data.summary_filename}\n\n` +
      `PSD Plots:\n${(data.plots || []).length} plot image(s) created\n\n` +
      `Package Download:\n${data.zip_filename} (Download started)`
    );

    showResultsModal(data);

  } catch (err) {
    btnGen.disabled = false;
    alert("Generation Error:\n\n" + err.message);
    statusEl.textContent = "Generation failed due to network or server error.";
  }
}

function showResultsModal(data) {
  document.getElementById('modalSummaryStats').innerHTML = `
    <strong>Borehole ID:</strong> ${data.bh_id} &nbsp;|&nbsp;
    <strong>Total GSD Depths:</strong> ${data.total_samples} &nbsp;|&nbsp;
    <strong>HYD Samples Generated:</strong> ${data.hyd_generated_samples}
  `;

  document.getElementById('btnDownloadExcel').href = data.excel_url;
  document.getElementById('btnDownloadExcel').setAttribute('download', data.excel_filename);

  document.getElementById('btnDownloadZip').href = data.zip_url;
  document.getElementById('btnDownloadZip').setAttribute('download', data.zip_filename);

  document.getElementById('btnDownloadSummary').href = data.summary_url;
  document.getElementById('btnDownloadSummary').setAttribute('download', data.summary_filename);

  document.getElementById('plotCountSpan').textContent = data.plots.length;

  // Render summary rows table
  const summaryTbody = document.getElementById('modalSummaryTableBody');
  summaryTbody.innerHTML = '';
  (data.summary_rows || []).forEach(r => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td style="text-align: center;">${r.bh_id}</td>
      <td style="text-align: center;">${r.depth}</td>
      <td style="text-align: center;">${r.sample_type}</td>
      <td style="text-align: center;">${r.sieve_set}</td>
      <td style="text-align: center;">${r.gravel !== null && r.gravel !== undefined ? parseFloat(r.gravel).toFixed(2) : '—'}</td>
      <td style="text-align: center;">${r.sand !== null && r.sand !== undefined ? parseFloat(r.sand).toFixed(2) : '—'}</td>
      <td style="text-align: center;">${r.silt !== null && r.silt !== undefined ? parseFloat(r.silt).toFixed(2) : '—'}</td>
      <td style="text-align: center;">${r.clay !== null && r.clay !== undefined ? parseFloat(r.clay).toFixed(2) : '—'}</td>
    `;
    summaryTbody.appendChild(tr);
  });

  // Render plot cards
  const plotsContainer = document.getElementById('modalPlotsContainer');
  plotsContainer.innerHTML = '';
  (data.plots || []).forEach((p, idx) => {
    const card = document.createElement('div');
    card.className = 'plot-card';
    card.innerHTML = `
      <div style="font-weight: bold; font-size: 9.5pt; margin-bottom: 6px; color: #1e3a8a;">
        Sample ${idx + 1}: Depth ${p.depth} m (${p.sample_type}) — ${p.filename}
      </div>
      <img src="${p.url}" alt="${p.filename}" loading="lazy">
      <div style="margin-top: 6px;">
        <a href="${p.url}" download="${p.filename}" class="tk-btn" style="padding: 2px 10px; font-size: 8.5pt; text-decoration: none;">
          <i class="fas fa-download" style="margin-right: 4px;"></i> Download PNG
        </a>
      </div>
    `;
    plotsContainer.appendChild(card);
  });

  document.getElementById('resultsModal').classList.add('show');
}

function closeResultsModal() {
  document.getElementById('resultsModal').classList.remove('show');
}
