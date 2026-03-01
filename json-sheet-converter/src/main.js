import "./style.css";
import * as XLSX from "xlsx";

// ── Mode tabs ──────────────────────────────────────────────────────────────
const tabButtons = document.querySelectorAll(".tab-btn");
const sectionJsonToExcel = document.querySelector("#section-json-to-excel");
const sectionExcelToJson = document.querySelector("#section-excel-to-json");
const previewJsonToExcel = document.querySelector("#preview-json-to-excel");
const previewExcelToJson = document.querySelector("#preview-excel-to-json");

const switchMode = (mode) => {
  tabButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.mode === mode));
  sectionJsonToExcel.hidden = mode !== "json-to-excel";
  sectionExcelToJson.hidden = mode !== "excel-to-json";
  previewJsonToExcel.hidden = mode !== "json-to-excel";
  previewExcelToJson.hidden = mode !== "excel-to-json";
};

tabButtons.forEach((btn) => {
  btn.addEventListener("click", () => switchMode(btn.dataset.mode));
});

// ── Shared utilities ───────────────────────────────────────────────────────
const sanitizeSheetName = (name, fallback) => {
  const cleaned = String(name || fallback)
    .replace(/[\\/*?:[\]]/g, "_")
    .slice(0, 31)
    .trim();
  return cleaned || fallback;
};

const toCellValue = (value) => {
  if (value === null || value === undefined) {
    return "";
  }

  if (typeof value === "object") {
    return JSON.stringify(value);
  }

  return value;
};

const renderSheets = (sheets, rootEl) => {
  rootEl.innerHTML = "";

  if (sheets.length === 0) {
    rootEl.innerHTML = '<p class="empty">No sheets to display.</p>';
    return;
  }

  sheets.forEach((sheet) => {
    const section = document.createElement("section");
    section.className = "sheet-card";

    const title = document.createElement("h3");
    title.textContent = `${sheet.name} (${sheet.rows.length} rows)`;
    section.appendChild(title);

    const columnList = document.createElement("p");
    columnList.className = "columns";
    columnList.textContent = sheet.columns.length > 0 ? sheet.columns.join(", ") : "No columns";
    section.appendChild(columnList);

    if (sheet.rows.length > 0) {
      const table = document.createElement("table");
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      sheet.columns.slice(0, 6).forEach((col) => {
        const th = document.createElement("th");
        th.textContent = col;
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      sheet.rows.slice(0, 5).forEach((row) => {
        const tr = document.createElement("tr");
        sheet.columns.slice(0, 6).forEach((col) => {
          const td = document.createElement("td");
          td.textContent = String(row[col] ?? "");
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      section.appendChild(table);
    }

    rootEl.appendChild(section);
  });
};

// ── JSON → Excel ────────────────────────────────────────────────────────────
const fileInput = document.querySelector("#file-input");
const dropzone = document.querySelector("#dropzone");
const convertButton = document.querySelector("#convert-button");
const downloadButton = document.querySelector("#download-button");
const workbookNameInput = document.querySelector("#workbook-name");
const statusEl = document.querySelector("#status");
const selectedFileEl = document.querySelector("#selected-file");
const sheetCountEl = document.querySelector("#sheet-count");
const previewRoot = document.querySelector("#preview-root");

let selectedFile = null;
let workbookData = null;

const setStatus = (message) => {
  statusEl.textContent = message;
};

const resetPreview = () => {
  workbookData = null;
  previewRoot.innerHTML = "";
  sheetCountEl.textContent = "0";
  downloadButton.disabled = true;
};

const flattenValue = (value, prefix = "", target = {}) => {
  if (Array.isArray(value)) {
    if (value.length === 0) {
      target[prefix || "value"] = "";
      return target;
    }

    value.forEach((item, index) => {
      const nextPrefix = prefix ? `${prefix}.${index}` : String(index);
      flattenValue(item, nextPrefix, target);
    });
    return target;
  }

  if (value && typeof value === "object") {
    const entries = Object.entries(value);
    if (entries.length === 0) {
      target[prefix || "value"] = "";
      return target;
    }

    entries.forEach(([key, nestedValue]) => {
      const nextPrefix = prefix ? `${prefix}.${key}` : key;
      flattenValue(nestedValue, nextPrefix, target);
    });
    return target;
  }

  target[prefix || "value"] = value;
  return target;
};

const normalizeRows = (value) => {
  if (Array.isArray(value)) {
    return value.map((item, index) => {
      const row = flattenValue(item);
      if (!("row_index" in row)) {
        row.row_index = index + 1;
      }
      return row;
    });
  }

  if (value && typeof value === "object") {
    return [flattenValue(value)];
  }

  return [{ value: toCellValue(value) }];
};

const buildSheetFromRows = (name, rows) => {
  const columnSet = new Set();
  rows.forEach((row) => {
    Object.keys(row).forEach((key) => columnSet.add(key));
  });

  const columns = Array.from(columnSet);
  const normalizedRows = rows.map((row) =>
    Object.fromEntries(columns.map((column) => [column, toCellValue(row[column])])),
  );

  return {
    name: sanitizeSheetName(name, "Sheet"),
    columns,
    rows: normalizedRows,
  };
};

const buildWordContentSheet = (content) => {
  const rows = content.map((block, index) => ({
    block_index: index + 1,
    type: block.type,
    level: block.level ?? "",
    ordered: block.ordered ?? "",
    text: block.text ?? "",
    rows_json: block.rows ? JSON.stringify(block.rows) : "",
    runs_json: block.runs ? JSON.stringify(block.runs) : "",
  }));

  return buildSheetFromRows("content_blocks", rows);
};

const buildWordTableSheet = (content) => {
  const rows = [];

  content.forEach((block, index) => {
    if (block.type !== "table" || !Array.isArray(block.rows)) {
      return;
    }

    block.rows.forEach((tableRow, rowIndex) => {
      const row = {
        block_index: index + 1,
        table_row_index: rowIndex + 1,
      };

      tableRow.forEach((cell, cellIndex) => {
        row[`column_${cellIndex + 1}`] = cell;
      });

      rows.push(row);
    });
  });

  return rows.length > 0 ? buildSheetFromRows("tables", rows) : null;
};

const buildWorkbookData = (json) => {
  const sheets = [];

  if (Array.isArray(json)) {
    sheets.push(buildSheetFromRows("data", normalizeRows(json)));
  } else if (json && typeof json === "object") {
    const scalarSummary = {};

    Object.entries(json).forEach(([key, value]) => {
      if (Array.isArray(value)) {
        sheets.push(buildSheetFromRows(key, normalizeRows(value)));
        return;
      }

      if (value && typeof value === "object") {
        sheets.push(buildSheetFromRows(key, normalizeRows(value)));
        return;
      }

      scalarSummary[key] = value;
    });

    if (Object.keys(scalarSummary).length > 0 || sheets.length === 0) {
      sheets.unshift(buildSheetFromRows("summary", normalizeRows(scalarSummary)));
    }
  } else {
    sheets.push(buildSheetFromRows("data", normalizeRows(json)));
  }

  if (json && Array.isArray(json.content)) {
    sheets.push(buildWordContentSheet(json.content));
    const tableSheet = buildWordTableSheet(json.content);
    if (tableSheet) {
      sheets.push(tableSheet);
    }
  }

  const uniqueSheets = [];
  const usedNames = new Set();

  sheets.forEach((sheet, index) => {
    let nextName = sheet.name || `Sheet${index + 1}`;
    let suffix = 2;
    while (usedNames.has(nextName)) {
      nextName = sanitizeSheetName(`${sheet.name}_${suffix}`, `Sheet${index + 1}`);
      suffix += 1;
    }
    usedNames.add(nextName);
    uniqueSheets.push({ ...sheet, name: nextName });
  });

  return uniqueSheets;
};

const convertSelectedFile = async () => {
  try {
    if (!selectedFile) {
      throw new Error("No JSON file selected.");
    }

    setStatus("Parsing JSON...");
    resetPreview();

    const text = await selectedFile.text();
    const json = JSON.parse(text);
    const sheets = buildWorkbookData(json);

    workbookData = sheets;
    sheetCountEl.textContent = String(sheets.length);
    renderSheets(sheets, previewRoot);
    downloadButton.disabled = false;
    setStatus("Ready to export spreadsheet.");
  } catch (error) {
    resetPreview();
    setStatus(error instanceof Error ? error.message : "Conversion failed.");
  }
};

const downloadWorkbook = () => {
  if (!workbookData) {
    return;
  }

  const workbook = XLSX.utils.book_new();

  workbookData.forEach((sheet) => {
    const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  });

  const rawName = workbookNameInput.value.trim() || "converted-workbook";
  const fileName = rawName.toLowerCase().endsWith(".xlsx") ? rawName : `${rawName}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

const handleFileSelection = (file) => {
  selectedFile = file;
  selectedFileEl.textContent = file ? file.name : "None";
  resetPreview();
  setStatus(file ? "Click Preview to continue." : "Select a JSON file");
};

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files ?? [];
  handleFileSelection(file ?? null);
});

convertButton.addEventListener("click", () => {
  void convertSelectedFile();
});

downloadButton.addEventListener("click", downloadWorkbook);

dropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropzone.classList.add("dragging");
});

dropzone.addEventListener("dragleave", () => {
  dropzone.classList.remove("dragging");
});

dropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropzone.classList.remove("dragging");
  const [file] = event.dataTransfer?.files ?? [];
  fileInput.files = event.dataTransfer?.files ?? null;
  handleFileSelection(file ?? null);
});

// ── Excel → JSON ────────────────────────────────────────────────────────────
const xlFileInput = document.querySelector("#xl-file-input");
const xlDropzone = document.querySelector("#xl-dropzone");
const xlPreviewButton = document.querySelector("#xl-preview-button");
const xlDownloadButton = document.querySelector("#xl-download-button");
const jsonNameInput = document.querySelector("#json-name");
const xlStatusEl = document.querySelector("#xl-status");
const xlSelectedFileEl = document.querySelector("#xl-selected-file");
const xlSheetCountEl = document.querySelector("#xl-sheet-count");
const xlPreviewRoot = document.querySelector("#xl-preview-root");

let xlSelectedFile = null;
let xlSheetsData = null;

const setXlStatus = (message) => {
  xlStatusEl.textContent = message;
};

const resetXlPreview = () => {
  xlSheetsData = null;
  xlPreviewRoot.innerHTML = "";
  xlSheetCountEl.textContent = "0";
  xlDownloadButton.disabled = true;
};

const buildSheetsFromExcel = (arrayBuffer) => {
  const workbook = XLSX.read(arrayBuffer, { type: "array" });

  return workbook.SheetNames.map((name) => {
    const ws = workbook.Sheets[name];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "" });
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
    return { name, columns, rows };
  });
};

const previewExcelFile = async () => {
  try {
    if (!xlSelectedFile) throw new Error("No spreadsheet selected.");

    setXlStatus("Reading spreadsheet...");
    resetXlPreview();

    const arrayBuffer = await xlSelectedFile.arrayBuffer();
    const sheets = buildSheetsFromExcel(arrayBuffer);

    xlSheetsData = sheets;
    xlSheetCountEl.textContent = String(sheets.length);
    renderSheets(sheets, xlPreviewRoot);
    xlDownloadButton.disabled = false;
    setXlStatus("Ready to save JSON.");
  } catch (err) {
    resetXlPreview();
    setXlStatus(err instanceof Error ? err.message : "Failed to read file.");
  }
};

const downloadJsonFromSheets = () => {
  if (!xlSheetsData) return;

  const payload =
    xlSheetsData.length === 1
      ? xlSheetsData[0].rows
      : Object.fromEntries(xlSheetsData.map((s) => [s.name, s.rows]));

  const rawName = jsonNameInput.value.trim() || "converted-data";
  const fileName = rawName.toLowerCase().endsWith(".json") ? rawName : `${rawName}.json`;

  const blob = new Blob([JSON.stringify(payload, null, 2)], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = fileName;
  anchor.click();
  URL.revokeObjectURL(url);
};

const handleXlFileSelection = (file) => {
  xlSelectedFile = file;
  xlSelectedFileEl.textContent = file ? file.name : "None";
  if (file) {
    jsonNameInput.value = file.name.replace(/\.(xlsx?|xls)$/i, "");
  }
  resetXlPreview();
  setXlStatus(file ? "Click Preview to continue." : "Select a spreadsheet");
};

xlFileInput.addEventListener("change", (event) => {
  const [file] = event.target.files ?? [];
  handleXlFileSelection(file ?? null);
});

xlPreviewButton.addEventListener("click", () => {
  void previewExcelFile();
});

xlDownloadButton.addEventListener("click", downloadJsonFromSheets);

xlDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  xlDropzone.classList.add("dragging");
});

xlDropzone.addEventListener("dragleave", () => {
  xlDropzone.classList.remove("dragging");
});

xlDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  xlDropzone.classList.remove("dragging");
  const [file] = event.dataTransfer?.files ?? [];
  if (file) {
    xlFileInput.files = event.dataTransfer?.files ?? null;
    handleXlFileSelection(file);
  }
});
