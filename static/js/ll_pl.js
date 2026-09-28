// LL & PL Synthetic Data Generator Frontend Logic
// Exact behavior, folder structure, plot image generation & validations matching LL_PL_Generator_15_Row_GUI_SemiLog_FlowCurve_FINAL_v3.py

const MAX_SAMPLES = 15;
const SAMPLE_TYPES = ["DS", "UDS", "SPT", "SHELBY", "OTHER"];

document.addEventListener("DOMContentLoaded", () => {
  const rowsContainer = document.getElementById("rowsContainer");
  const bhInput = document.getElementById("bh_var");
  const chainageInput = document.getElementById("chainage_var");
  const statusLabel = document.getElementById("statusLabel");
  const btnClearAll = document.getElementById("btnClearAll");
  const btnGenerateExcel = document.getElementById("btnGenerateExcel");

  // Build the 15 input rows exactly matching desktop Tkinter table
  for (let i = 0; i < MAX_SAMPLES; i++) {
    const rowDiv = document.createElement("div");
    rowDiv.className = "tk-row-grid";

    // 1. Sr. No.
    const srLabel = document.createElement("div");
    srLabel.className = "tk-row-label";
    srLabel.textContent = (i + 1).toString();
    rowDiv.appendChild(srLabel);

    // 2. Depth
    const depthInput = document.createElement("input");
    depthInput.type = "text";
    depthInput.className = "tk-row-input";
    depthInput.id = `depth_${i}`;
    rowDiv.appendChild(depthInput);

    // 3. Sample Type
    const typeSelect = document.createElement("select");
    typeSelect.className = "tk-row-select";
    typeSelect.id = `type_${i}`;
    SAMPLE_TYPES.forEach(t => {
      const opt = document.createElement("option");
      opt.value = t;
      opt.textContent = t;
      if (t === "DS") opt.selected = true;
      typeSelect.appendChild(opt);
    });
    rowDiv.appendChild(typeSelect);

    // 4. Liquid Limit (LL)
    const llInput = document.createElement("input");
    llInput.type = "text";
    llInput.className = "tk-row-input";
    llInput.id = `ll_${i}`;
    rowDiv.appendChild(llInput);

    // 5. Plastic Limit (PL)
    const plInput = document.createElement("input");
    plInput.type = "text";
    plInput.className = "tk-row-input";
    plInput.id = `pl_${i}`;
    rowDiv.appendChild(plInput);

    rowsContainer.appendChild(rowDiv);
  }

  // Clear All matching desktop script messagebox prompt
  btnClearAll.addEventListener("click", () => {
    if (!confirm("Clear Borehole, Chainage and all 15 sample rows?")) {
      return;
    }

    bhInput.value = "";
    chainageInput.value = "";

    for (let i = 0; i < MAX_SAMPLES; i++) {
      document.getElementById(`depth_${i}`).value = "";
      document.getElementById(`type_${i}`).value = "DS";
      document.getElementById(`ll_${i}`).value = "";
      document.getElementById(`pl_${i}`).value = "";
    }

    statusLabel.textContent = "All input fields cleared.";
  });

  // Validation function strictly matching collect_samples from the master script
  function collectSamples() {
    const bh = bhInput.value.trim();
    const chainage = chainageInput.value.trim();

    if (!bh) {
      throw new Error("Borehole No. is required.");
    }
    if (!chainage) {
      throw new Error("Chainage is required.");
    }

    const samples = [];

    for (let i = 0; i < MAX_SAMPLES; i++) {
      const depthText = document.getElementById(`depth_${i}`).value.trim();
      const typeText = document.getElementById(`type_${i}`).value.trim();
      const llText = document.getElementById(`ll_${i}`).value.trim();
      const plText = document.getElementById(`pl_${i}`).value.trim();

      // Completely blank row = unused row
      if (!depthText && !llText && !plText) {
        continue;
      }

      // A partially filled row is an error
      if (!depthText) {
        throw new Error(`Row ${i + 1}: Depth is required.`);
      }
      if (!llText) {
        throw new Error(`Row ${i + 1}: Liquid Limit (LL) is required.`);
      }

      const depth = Number(depthText);
      if (isNaN(depth) || depthText === "") {
        throw new Error(`Row ${i + 1}: Depth must be a number.`);
      }
      if (depth < 0) {
        throw new Error(`Row ${i + 1}: Depth cannot be negative.`);
      }

      const ll = Number(llText);
      if (isNaN(ll) || llText === "") {
        throw new Error(`Row ${i + 1}: Liquid Limit (LL) must be a number.`);
      }
      if (ll <= 0 || ll >= 100) {
        throw new Error(`Row ${i + 1}: Liquid Limit (LL) must be between 0 and 100.`);
      }

      let pl = null;
      if (plText) {
        pl = Number(plText);
        if (isNaN(pl) || plText === "") {
          throw new Error(`Row ${i + 1}: Plastic Limit (PL) must be a number.`);
        }
        if (pl <= 0 || pl >= 100) {
          throw new Error(`Row ${i + 1}: Plastic Limit (PL) must be between 0 and 100.`);
        }
        if (pl >= ll) {
          throw new Error(`Row ${i + 1}: Plastic Limit (PL) must be less than Liquid Limit (LL).`);
        }
      }

      samples.push({
        sr: samples.length + 1,
        bh: bh,
        chainage: chainage,
        depth: depth,
        sample_type: typeText || "DS",
        ll: ll,
        pl: pl
      });
    }

    if (samples.length === 0) {
      throw new Error("Please enter at least one sample.");
    }

    return { bh, chainage, samples };
  }

  async function saveFileWithPicker(url, filename, base64Data, mimeType) {
    let blob = null;

    if (base64Data) {
      try {
        const byteCharacters = atob(base64Data);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i++) {
          byteNumbers[i] = byteCharacters.charCodeAt(i);
        }
        const byteArray = new Uint8Array(byteNumbers);
        blob = new Blob([byteArray], { type: mimeType || "application/zip" });
      } catch (e) {
        console.warn("Base64 parsing failed", e);
      }
    }

    if (!blob && url) {
      try {
        const res = await fetch(url);
        blob = await res.blob();
      } catch (e) {
        console.warn("Blob fetch failed", e);
      }
    }

    // 1. If Browser supports native Save File Location Picker dialog
    if (window.showSaveFilePicker && blob) {
      try {
        const handle = await window.showSaveFilePicker({
          suggestedName: filename,
          types: [{
            description: "ZIP Archive",
            accept: { "application/zip": [".zip"] }
          }]
        });
        const writable = await handle.createWritable();
        await writable.write(blob);
        await writable.close();
        return true;
      } catch (err) {
        if (err.name === "AbortError") {
          // User clicked Cancel in the Save Dialog
          return false;
        }
        console.warn("showSaveFilePicker failed, fallback to standard download", err);
      }
    }

    // 2. Standard direct download fallback
    if (blob) {
      const blobUrl = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = blobUrl;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
        URL.revokeObjectURL(blobUrl);
      }, 300);
      return true;
    } else if (url) {
      const a = document.createElement("a");
      a.href = url.includes("?") ? `${url}&download=1` : `${url}?download=1`;
      a.download = filename;
      document.body.appendChild(a);
      a.click();
      setTimeout(() => {
        document.body.removeChild(a);
      }, 300);
      return true;
    }
    return false;
  }

  // Generate Excel & Flow Plots matching exact backend script logic
  btnGenerateExcel.addEventListener("click", async () => {
    let collected;
    try {
      collected = collectSamples();
    } catch (err) {
      alert("Invalid Input\n\n" + err.message);
      return;
    }

    const { bh, chainage, samples } = collected;

    statusLabel.textContent = `Generating data for ${samples.length} sample(s)... Please wait.`;

    try {
      const response = await fetch("/api/ll_pl/export_excel", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          bh_id: bh,
          chainage: chainage,
          rows: samples
        })
      });

      const resData = await response.json().catch(() => ({}));
      if (!response.ok || !resData.success) {
        throw new Error(resData.error || "Generation failed on server.");
      }

      statusLabel.textContent = `Generation completed successfully. ${resData.count} sample(s) generated.`;

      // Prompt user to choose save folder/location for the complete ZIP package
      const zipName = resData.zip_filename || `LL_PL_Package_${bh}.zip`;
      const saved = await saveFileWithPicker(
        resData.download_zip_url,
        zipName,
        resData.file_base64,
        "application/zip"
      );

      if (saved !== false) {
        // Completion alert matching desktop script
        alert(
          `Generation Completed\n\n` +
          `LL & PL generation completed successfully.\n\n` +
          `Samples generated: ${resData.count}\n\n` +
          `ZIP Package (Excel + Plots + PDF):\n${zipName}\n\n` +
          `Excel output:\n${resData.excel_filename}\n\n` +
          `Flow plots folder:\n${resData.plot_dir}\n\n` +
          `Combined PDF:\n${resData.plot_pdf}`
        );
      } else {
        statusLabel.textContent = "Save cancelled by user.";
      }
    } catch (err) {
      statusLabel.textContent = "Generation failed.";
      alert("Generation Failed\n\n" + err.message);
    }
  });
});

