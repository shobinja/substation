let rawWorkbookData = null;

const dropzone = document.getElementById('dropzone');
const fileInput = document.getElementById('fileInput');
const fileStatus = document.getElementById('fileStatus');
const btnSubmit = document.getElementById('btnSubmit');
const noticeBox = document.getElementById('noticeBox');
const progressLog = document.getElementById('progressLog');

const REPORT_CONFIG = {
  RMU: {
    title: 'RMU',
    sheetCandidates: ['RMU', 'rmu'],
    templateCandidates: ['Template/RMU.docx', 'template/RMU.docx', './Template/RMU.docx'],
    outputSuffix: 'RMU.docx',
    reportCategory: 'single'
  },
  HV_CABLE: {
    title: 'HV Cable',
    sheetCandidates: ['HV_CABLE', 'HV_Cable', 'hv_cable', 'HV Cable'],
    templateCandidates: ['Template/HV_CABLE.docx', 'template/HV_CABLE.docx', './Template/HV_CABLE.docx'],
    reportLabel: 'HV_Cable',
    reportCategory: 'cable'
  },
  LV_CABLE: {
    title: 'LV Cable',
    sheetCandidates: ['LV_CABLE', 'LV_Cable', 'lv_cable', 'LV Cable'],
    templateCandidates: ['Template/LV_CABLE.docx', 'template/LV_CABLE.docx', './Template/LV_CABLE.docx'],
    reportLabel: 'LV_Cable',
    reportCategory: 'cable'
  },
  TX_IR: {
    title: 'Tx IR test',
    sheetCandidates: ['Tx_IR', 'Tx_ir', 'TX_IR', 'Tx IR'],
    templateCandidates: ['Template/TX_IR.docx', 'template/TX_IR.docx', './Template/TX_IR.docx'],
    outputSuffix: 'TX_IR.docx',
    reportCategory: 'transformer'
  },
  TX_RATIO: {
    title: 'Tx Ratio test',
    sheetCandidates: ['Tx_Ratio', 'Tx_ratio', 'TX_RATIO', 'Tx Ratio'],
    templateCandidates: ['Template/TX_RATIO.docx', 'template/TX_RATIO.docx', './Template/TX_RATIO.docx'],
    outputSuffix: 'TX_RATIO.docx',
    reportCategory: 'transformer'
  },
  TX_VECTOR: {
    title: 'Tx Vector test',
    sheetCandidates: ['Tx_Vector', 'Tx_vector', 'TX_VECTOR', 'Tx Vector'],
    templateCandidates: ['Template/TX_VECTOR.docx', 'template/TX_VECTOR.docx', './Template/TX_VECTOR.docx'],
    outputSuffix: 'TX_VECTOR.docx',
    reportCategory: 'transformer'
  },
  TX_BDV: {
    title: 'Tx BDV test',
    sheetCandidates: ['Tx_BDV', 'Tx BDV', 'Tx_bdv', 'TX_BDV'],
    templateCandidates: ['Template/TX_BDV.docx', 'template/TX_BDV.docx', './Template/TX_BDV.docx'],
    outputSuffix: 'TX_BDV.docx',
    reportCategory: 'transformer'
  },
  ACB: {
    title: 'ACB Test report',
    sheetCandidates: ['ACB', 'acb', 'ACB_Test', 'ACB Test'],
    templateCandidates: ['Template/ACB.docx', 'template/ACB.docx', './Template/ACB.docx'],
    outputSuffix: 'ACB.docx',
    reportCategory: 'single'
  }
};

function logStep(msg, type = 'info') {
  progressLog.style.display = 'block';
  const item = document.createElement('div');
  item.className = `log-item ${type}`;
  const time = new Date().toLocaleTimeString();
  item.textContent = `[${time}] ${msg}`;
  progressLog.appendChild(item);
  progressLog.scrollTop = progressLog.scrollHeight;
}

function clearLog() {
  progressLog.innerHTML = '';
  progressLog.style.display = 'none';
}

function showNotice(text, type = 'error') {
  noticeBox.textContent = text;
  noticeBox.className = `notification ${type}`;
}

function clearNotice() {
  noticeBox.textContent = '';
  noticeBox.className = 'notification';
}

function sanitizeForFilename(str) {
  if (!str) return 'Report';
  return String(str)
    .replace(/[^a-zA-Z0-9]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '');
}

function formatValue(tag, val) {
  if (val === undefined || val === null) return "";

  if (typeof val === 'boolean') {
    return val ? 'OK' : 'N/A';
  }

  const twoDecimalRegex = /^(calRuv[1-5]|calRvw[1-5]|calRwu[1-5]|avgt[1-5])$/i;
  if (typeof tag === 'string' && twoDecimalRegex.test(tag.trim())) {
    const num = parseFloat(val);
    if (!isNaN(num)) {
      return num.toFixed(2);
    }
  }

  if (val instanceof Date) {
    const dd = String(val.getDate()).padStart(2, '0');
    const mm = String(val.getMonth() + 1).padStart(2, '0');
    const yyyy = val.getFullYear();
    return `${dd}-${mm}-${yyyy}`;
  }

  return String(val).trim();
}

function extractAllRows(sheet) {
  const data = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, cellDates: true });
  if (data.length < 2) return [];

  const tagsRow = data[1] || []; // Row 2: Tag names
  const result = [];

  for (let r = 2; r < data.length; r++) {
    const dataRow = data[r] || [];
    if (!dataRow.some(v => v !== null && v !== undefined && v !== '')) continue;

    const rowMap = {};
    tagsRow.forEach((tag, idx) => {
      if (tag) {
        const key = String(tag).trim();
        const formatted = formatValue(key, dataRow[idx]);
        rowMap[key] = formatted;
        rowMap[key.toUpperCase()] = formatted;
        rowMap[key.toLowerCase()] = formatted;
      }
    });
    result.push(rowMap);
  }
  return result;
}

function findSheet(workbook, candidateNames) {
  for (const name of candidateNames) {
    if (workbook.Sheets[name]) return workbook.Sheets[name];
  }
  const sheetNames = workbook.SheetNames;
  for (const candidate of candidateNames) {
    const target = candidate.toLowerCase().replace(/[\s_]/g, '');
    const match = sheetNames.find(s => s.toLowerCase().replace(/[\s_]/g, '') === target);
    if (match) return workbook.Sheets[match];
  }
  return null;
}

async function loadTemplate(candidatePaths) {
  for (const path of candidatePaths) {
    try {
      const res = await fetch(path);
      if (res.ok) return await res.arrayBuffer();
    } catch (e) {
      // try next path candidate
    }
  }
  throw new Error(`Template not found: ${candidatePaths[0]}`);
}

function handleFile(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext !== 'xlsx' && ext !== 'xlsm') {
    showNotice("Please upload a valid .xlsx or .xlsm file.");
    fileInput.value = '';
    rawWorkbookData = null;
    fileStatus.textContent = '';
    return;
  }

  clearNotice();
  clearLog();
  fileStatus.textContent = `Selected: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;

  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      rawWorkbookData = new Uint8Array(e.target.result);
    } catch (err) {
      showNotice("Failed to read file: " + err.message);
    }
  };
  reader.readAsArrayBuffer(file);
}

fileInput.addEventListener('change', (e) => handleFile(e.target.files[0]));
dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
dropzone.addEventListener('drop', (e) => {
  e.preventDefault();
  dropzone.classList.remove('dragover');
  if (e.dataTransfer.files.length > 0) handleFile(e.dataTransfer.files[0]);
});

btnSubmit.addEventListener('click', async () => {
  clearNotice();
  clearLog();

  if (!rawWorkbookData) {
    showNotice("Please upload an Excel workbook first.");
    return;
  }

  const selectedCheckboxes = Array.from(document.querySelectorAll('input[name="reportType"]:checked'));
  if (selectedCheckboxes.length === 0) {
    showNotice("Please select at least one report to generate.");
    return;
  }

  const PizZipConstructor = window.PizZip;
  const DocxtemplaterConstructor = window.docxtemplater;
  const JSZipConstructor = window.JSZip;

  if (!PizZipConstructor || !DocxtemplaterConstructor || !JSZipConstructor) {
    showNotice("Document libraries failed to load. Check your internet connection.");
    return;
  }

  btnSubmit.disabled = true;
  btnSubmit.textContent = "Processing...";

  try {
    logStep("Step 1/5: Parsing uploaded Excel workbook...");
    const workbook = XLSX.read(rawWorkbookData, { type: 'array', cellDates: true });

    // Validate General Sheet
    const generalSheet = findSheet(workbook, ['General', 'general']);
    if (!generalSheet) {
      throw new Error("Workbook must contain the 'General' sheet.");
    }
    const generalRows = extractAllRows(generalSheet);
    const generalData = generalRows[0] || {};
    logStep("✓ 'General' sheet loaded successfully.", "success");

    // Format RD & BLK in General
    let formattedRD = "";
    if (generalData["RD"] && generalData["RD"].length > 0) {
      formattedRD = `, RD ${generalData["RD"]}`;
    }
    let formattedBLK = "";
    if (generalData["BLK"] && generalData["BLK"].length > 0) {
      formattedBLK = `, BLK ${generalData["BLK"]}`;
    }

    const baseData = Object.assign({}, generalData, {
      RD: formattedRD,
      rd: formattedRD,
      BLK: formattedBLK,
      blk: formattedBLK
    });

    if (baseData["date"] && !baseData["DATE"]) {
      baseData["DATE"] = baseData["date"];
    }

    // Check if Transformer reports are selected
    const needsTx = selectedCheckboxes.some(cb => {
      const conf = REPORT_CONFIG[cb.value];
      return conf && conf.reportCategory === 'transformer';
    });

    let txDtlsRows = [];
    if (needsTx) {
      logStep("Step 2/5: Reading 'TxDtls' transformer specifications...");
      const txSheet = findSheet(workbook, ['TxDtls', 'txdtls', 'Tx_Dtls', 'Tx Details']);
      if (!txSheet) {
        throw new Error("Transformer reports require the 'TxDtls' sheet, which was not found.");
      }
      txDtlsRows = extractAllRows(txSheet);
      if (txDtlsRows.length === 0) {
        throw new Error("No data rows found in 'TxDtls' sheet.");
      }
      logStep(`✓ Found ${txDtlsRows.length} transformer(s) in 'TxDtls'.`, "success");
    }

    const outputZip = new JSZipConstructor();
    let commonSS = baseData["SS"] || baseData["ss"] || "Report";
    const sanitizedSS = sanitizeForFilename(commonSS);
    let totalGenerated = 0;

    logStep("Step 3/5: Generating selected report documents...");

    for (const cb of selectedCheckboxes) {
      const reportKey = cb.value;
      const config = REPORT_CONFIG[reportKey];
      if (!config) continue;

      const reportSheet = findSheet(workbook, config.sheetCandidates);
      if (!reportSheet) {
        logStep(`⚠ Skipped '${config.title}': Sheet not in workbook.`, "warning");
        continue;
      }

      const sheetRows = extractAllRows(reportSheet);
      if (sheetRows.length === 0) {
        logStep(`⚠ Skipped '${config.title}': No data rows in sheet.`, "warning");
        continue;
      }

      logStep(`Loading template for '${config.title}'...`);
      let templateBuffer;
      try {
        templateBuffer = await loadTemplate(config.templateCandidates);
      } catch (err) {
        logStep(`⚠ Error loading template for '${config.title}': ${err.message}`, "error");
        continue;
      }

      // --- Category 1: Cable Reports (HV Cable & LV Cable - loop every row) ---
      if (config.reportCategory === 'cable') {
        logStep(`Processing ${sheetRows.length} cable circuit(s) for '${config.title}'...`);

        sheetRows.forEach((cableRow, rowIdx) => {
          const combinedData = Object.assign({}, baseData, cableRow);

          // Get circuit name for filename (CBL_CKT)
          const circuitRaw = cableRow["CBL_CKT"] || cableRow["cbl_ckt"] || `Circuit_${rowIdx + 1}`;
          const sanitizedCircuit = sanitizeForFilename(circuitRaw);

          // Naming format: [SS]_[ReportLabel]_[CBL_CKT].docx
          const fileName = `${sanitizedSS}_${config.reportLabel}_${sanitizedCircuit}.docx`;

          const zip = new PizZipConstructor(templateBuffer);
          const doc = new DocxtemplaterConstructor(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{', end: '}' },
            nullGetter: () => ""
          });

          doc.setData(combinedData);
          doc.render();

          const docxArrayBuffer = doc.getZip().generate({
            type: "arraybuffer",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          });

          outputZip.file(fileName, docxArrayBuffer);
          totalGenerated++;
          logStep(`✓ Created: ${fileName}`, "success");
        });

      // --- Category 2: Transformer Reports ---
      } else if (config.reportCategory === 'transformer') {
        for (const txRow of txDtlsRows) {
          const currentTxId = txRow["Tx_ID"] || txRow["tx_id"] || "";
          const currentTxNo = txRow["TxNo"] || txRow["txno"] || currentTxId;

          const matchedTestRow = sheetRows.find(r => {
            const ref = r["Tx_Ref"] || r["tx_ref"] || "";
            return String(ref).trim() === String(currentTxId).trim();
          }) || sheetRows[0];

          const combinedData = Object.assign({}, baseData, txRow, matchedTestRow);

          combinedData["Tx_ID"] = currentTxNo;
          combinedData["tx_id"] = currentTxNo;
          combinedData["TxNo"] = currentTxNo;
          combinedData["txno"] = currentTxNo;

          const zip = new PizZipConstructor(templateBuffer);
          const doc = new DocxtemplaterConstructor(zip, {
            paragraphLoop: true,
            linebreaks: true,
            delimiters: { start: '{', end: '}' },
            nullGetter: () => ""
          });

          doc.setData(combinedData);
          doc.render();

          const docxArrayBuffer = doc.getZip().generate({
            type: "arraybuffer",
            mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
          });

          const txSuffix = txDtlsRows.length > 1 ? `_TX${currentTxNo}_` : '_';
          const fileName = `${sanitizedSS}${txSuffix}${config.outputSuffix}`;
          outputZip.file(fileName, docxArrayBuffer);
          totalGenerated++;

          logStep(`✓ Created: ${fileName}`, "success");
        }

      // --- Category 3: Single-Item Reports (RMU, ACB) ---
      } else {
        const combinedData = Object.assign({}, baseData, sheetRows[0]);

        if (combinedData["ACB_I"] !== undefined) {
          combinedData[" ACB_I"] = combinedData["ACB_I"];
        }

        const zip = new PizZipConstructor(templateBuffer);
        const doc = new DocxtemplaterConstructor(zip, {
          paragraphLoop: true,
          linebreaks: true,
          delimiters: { start: '{', end: '}' },
          nullGetter: () => ""
        });

        doc.setData(combinedData);
        doc.render();

        const docxArrayBuffer = doc.getZip().generate({
          type: "arraybuffer",
          mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        });

        const fileName = `${sanitizedSS}_${config.outputSuffix}`;
        outputZip.file(fileName, docxArrayBuffer);
        totalGenerated++;

        logStep(`✓ Created: ${fileName}`, "success");
      }
    }

    if (totalGenerated === 0) {
      throw new Error("No reports could be generated. Please verify sheet names and selected options.");
    }

    logStep(`Step 4/5: Compressing ${totalGenerated} document(s) into zip archive...`);
    const zipArchiveName = `${sanitizedSS}_Reports.zip`;
    const zipBlob = await outputZip.generateAsync({ type: "blob" });

    logStep(`Step 5/5: Initiating download for ${zipArchiveName}...`, "success");
    saveAs(zipBlob, zipArchiveName);

    showNotice(`Successfully generated and downloaded ${zipArchiveName} with ${totalGenerated} report(s)!`, "success");

  } catch (err) {
    logStep(`❌ Failed: ${err.message}`, "error");
    showNotice(err.message || "An error occurred while generating reports.");
  } finally {
    btnSubmit.disabled = false;
    btnSubmit.textContent = "Generate Selected Reports";
  }
});