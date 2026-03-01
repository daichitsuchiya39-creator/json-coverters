import "./style.css";
import * as XLSX from "xlsx";

const fileInput = document.querySelector("#file-input");
const dropzone = document.querySelector("#dropzone");
const convertButton = document.querySelector("#convert-button");
const downloadButton = document.querySelector("#download-button");
const workbookNameInput = document.querySelector("#workbook-name");
const statusElement = document.querySelector("#status");
const selectedFileElement = document.querySelector("#selected-file");
const sheetCountElement = document.querySelector("#sheet-count");
const previewRoot = document.querySelector("#preview-root");

let selectedFile = null;
let workbookData = null;

const setStatus = (message) => {
  statusElement.textContent = message;
};

const setSelectedFile = (file) => {
  selectedFile = file;
  selectedFileElement.textContent = file ? file.name : "None";
};

const resetPreview = () => {
  workbookData = null;
  previewRoot.innerHTML = "";
  sheetCountElement.textContent = "0";
  downloadButton.disabled = true;
};

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

const buildGenericSheets = (json) => {
  if (Array.isArray(json)) {
    return [buildSheetFromRows("data", normalizeRows(json))];
  }

  if (json && typeof json === "object") {
    const sheets = [];
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

    return sheets;
  }

  return [buildSheetFromRows("data", normalizeRows(json))];
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
  const sheets = buildGenericSheets(json);

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

const renderPreview = (sheets) => {
  previewRoot.innerHTML = "";

  if (sheets.length === 0) {
    previewRoot.innerHTML = "<p class=\"empty\">No sheets to output.</p>";
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
    columnList.textContent = sheet.columns.length > 0 ? sheet.columns.join(", ") : "列なし";
    section.appendChild(columnList);

    if (sheet.rows.length > 0) {
      const table = document.createElement("table");
      const thead = document.createElement("thead");
      const headRow = document.createElement("tr");
      sheet.columns.slice(0, 6).forEach((column) => {
        const th = document.createElement("th");
        th.textContent = column;
        headRow.appendChild(th);
      });
      thead.appendChild(headRow);
      table.appendChild(thead);

      const tbody = document.createElement("tbody");
      sheet.rows.slice(0, 5).forEach((row) => {
        const tr = document.createElement("tr");
        sheet.columns.slice(0, 6).forEach((column) => {
          const td = document.createElement("td");
          td.textContent = String(row[column] ?? "");
          tr.appendChild(td);
        });
        tbody.appendChild(tr);
      });
      table.appendChild(tbody);
      section.appendChild(table);
    }

    previewRoot.appendChild(section);
  });
};

const parseJsonFile = async (file) => {
  const text = await file.text();
  return JSON.parse(text);
};

const createWorkbook = (sheets) => {
  const workbook = XLSX.utils.book_new();

  sheets.forEach((sheet) => {
    const worksheet = XLSX.utils.json_to_sheet(sheet.rows);
    XLSX.utils.book_append_sheet(workbook, worksheet, sheet.name);
  });

  return workbook;
};

const convertSelectedFile = async () => {
  try {
    if (!selectedFile) {
      throw new Error("No JSON file selected.");
    }

    setStatus("Parsing JSON...");
    resetPreview();

    const json = await parseJsonFile(selectedFile);
    const sheets = buildWorkbookData(json);

    workbookData = sheets;
    sheetCountElement.textContent = String(sheets.length);
    renderPreview(sheets);
    downloadButton.disabled = false;
    setStatus("Ready to export Excel.");
  } catch (error) {
    resetPreview();
    setStatus(error instanceof Error ? error.message : "Conversion failed.");
  }
};

const downloadWorkbook = () => {
  if (!workbookData) {
    return;
  }

  const workbook = createWorkbook(workbookData);
  const rawName = workbookNameInput.value.trim() || "converted-workbook";
  const fileName = rawName.toLowerCase().endsWith(".xlsx") ? rawName : `${rawName}.xlsx`;
  XLSX.writeFile(workbook, fileName);
};

const handleFileSelection = (file) => {
  setSelectedFile(file);
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
