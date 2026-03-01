import "./style.css";
import * as mammoth from "mammoth/mammoth.browser";
import { AlignmentType, Document, HeadingLevel, Packer, Paragraph, Table, TableCell, TableRow, TextRun } from "docx";

// ── Mode tabs ──────────────────────────────────────────────────────────────
const tabButtons = document.querySelectorAll(".tab-btn");
const sectionWordToJson = document.querySelector("#section-word-to-json");
const sectionJsonToWord = document.querySelector("#section-json-to-word");
const outputWordToJson = document.querySelector("#output-word-to-json");
const outputJsonToWord = document.querySelector("#output-json-to-word");

const switchMode = (mode) => {
  tabButtons.forEach((btn) => btn.classList.toggle("active", btn.dataset.mode === mode));
  sectionWordToJson.hidden = mode !== "word-to-json";
  sectionJsonToWord.hidden = mode !== "json-to-word";
  outputWordToJson.hidden = mode !== "word-to-json";
  outputJsonToWord.hidden = mode !== "json-to-word";
};

tabButtons.forEach((btn) => {
  btn.addEventListener("click", () => switchMode(btn.dataset.mode));
});

// ── Word → JSON ────────────────────────────────────────────────────────────
const fileInput = document.querySelector("#file-input");
const convertButton = document.querySelector("#convert-button");
const downloadButton = document.querySelector("#download-button");
const jsonOutput = document.querySelector("#json-output");
const statusEl = document.querySelector("#status");
const selectedFileEl = document.querySelector("#selected-file");
const blockCountEl = document.querySelector("#block-count");
const dropzone = document.querySelector("#dropzone");

let selectedFile = null;
let convertedJson = null;

const setStatus = (message) => {
  statusEl.textContent = message;
};

const updateOutput = (payload) => {
  convertedJson = payload;
  jsonOutput.textContent = JSON.stringify(payload, null, 2);
  blockCountEl.textContent = String(payload.content.length);
  downloadButton.disabled = false;
};

const resetOutput = () => {
  convertedJson = null;
  jsonOutput.textContent = "{}";
  blockCountEl.textContent = "0";
  downloadButton.disabled = true;
};

const createTextRuns = (node) =>
  Array.from(node.childNodes)
    .filter((child) => child.nodeType === Node.TEXT_NODE || child.nodeType === Node.ELEMENT_NODE)
    .flatMap((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent?.trim();
        return text ? [{ text }] : [];
      }

      const text = child.textContent?.trim();
      if (!text) {
        return [];
      }

      return [
        {
          text,
          bold: child.tagName === "STRONG",
          italic: child.tagName === "EM",
          underline: child.tagName === "U",
        },
      ];
    });

const parseList = (listNode, ordered = false, level = 1) =>
  Array.from(listNode.children).flatMap((item) => {
    if (item.tagName !== "LI") {
      return [];
    }

    const blocks = [];
    const textParts = [];

    Array.from(item.childNodes).forEach((child) => {
      if (child.nodeType === Node.TEXT_NODE) {
        const text = child.textContent?.trim();
        if (text) {
          textParts.push(text);
        }
        return;
      }

      if (child.tagName === "UL" || child.tagName === "OL") {
        blocks.push(...parseList(child, child.tagName === "OL", level + 1));
        return;
      }

      const text = child.textContent?.trim();
      if (text) {
        textParts.push(text);
      }
    });

    const text = textParts.join(" ").trim();
    const itemBlock = text
      ? [
          {
            type: "list-item",
            ordered,
            level,
            text,
          },
        ]
      : [];

    return [...itemBlock, ...blocks];
  });

const parseTable = (tableNode) => ({
  type: "table",
  rows: Array.from(tableNode.querySelectorAll("tr")).map((row) =>
    Array.from(row.querySelectorAll("th, td")).map((cell) => cell.textContent?.trim() ?? ""),
  ),
});

const parseBlock = (node) => {
  const text = node.textContent?.trim() ?? "";
  if (!text && !["TABLE", "UL", "OL"].includes(node.tagName)) {
    return [];
  }

  if (/^H[1-6]$/.test(node.tagName)) {
    return [
      {
        type: "heading",
        level: Number(node.tagName.slice(1)),
        text,
      },
    ];
  }

  if (node.tagName === "P") {
    return [
      {
        type: "paragraph",
        text,
        runs: createTextRuns(node),
      },
    ];
  }

  if (node.tagName === "UL" || node.tagName === "OL") {
    return parseList(node, node.tagName === "OL");
  }

  if (node.tagName === "TABLE") {
    return [parseTable(node)];
  }

  return [
    {
      type: "text",
      text,
    },
  ];
};

const htmlToJson = (html, fileName, messages) => {
  const parser = new DOMParser();
  const documentFragment = parser.parseFromString(html, "text/html");
  const content = Array.from(documentFragment.body.children).flatMap(parseBlock);

  return {
    fileName,
    convertedAt: new Date().toISOString(),
    messages: messages.map((message) => ({
      type: message.type,
      message: message.message,
    })),
    content,
  };
};

const convertFile = async () => {
  try {
    if (!selectedFile) throw new Error("No file selected.");
    if (!selectedFile.name.toLowerCase().endsWith(".docx")) throw new Error("Please select a .docx file.");

    setStatus("Converting...");
    resetOutput();

    const arrayBuffer = await selectedFile.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const payload = htmlToJson(result.value, selectedFile.name, result.messages);

    updateOutput(payload);
    setStatus("Conversion complete.");
  } catch (error) {
    resetOutput();
    setStatus(error instanceof Error ? error.message : "Conversion failed.");
  }
};

const downloadJson = () => {
  if (!convertedJson) return;

  const blob = new Blob([JSON.stringify(convertedJson, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${selectedFile?.name.replace(/\.docx$/i, "") ?? "document"}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
};

const handleFileSelection = (file) => {
  selectedFile = file;
  selectedFileEl.textContent = file ? file.name : "None";
  setStatus(file ? "Ready to convert." : "Select a file");
  resetOutput();
};

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files ?? [];
  handleFileSelection(file ?? null);
});

convertButton.addEventListener("click", () => {
  void convertFile();
});

downloadButton.addEventListener("click", downloadJson);

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

// ── JSON → Word ────────────────────────────────────────────────────────────
const jsonFileInput = document.querySelector("#json-file-input");
const generateButton = document.querySelector("#generate-button");
const jsonInputPreview = document.querySelector("#json-input-preview");
const jsonStatusEl = document.querySelector("#json-status");
const jsonSelectedFileEl = document.querySelector("#json-selected-file");
const jsonBlockCountEl = document.querySelector("#json-block-count");
const jsonDropzone = document.querySelector("#json-dropzone");

let loadedJson = null;
let jsonInputFileName = null;

const HEADING_MAP = {
  1: HeadingLevel.HEADING_1,
  2: HeadingLevel.HEADING_2,
  3: HeadingLevel.HEADING_3,
  4: HeadingLevel.HEADING_4,
  5: HeadingLevel.HEADING_5,
  6: HeadingLevel.HEADING_6,
};

const blockToDocxElement = (block) => {
  if (block.type === "heading") {
    return new Paragraph({
      heading: HEADING_MAP[block.level] ?? HeadingLevel.HEADING_1,
      children: [new TextRun(block.text ?? "")],
    });
  }

  if (block.type === "paragraph") {
    const runs = block.runs?.length
      ? block.runs.map(
          (run) =>
            new TextRun({
              text: run.text,
              bold: run.bold ?? false,
              italics: run.italic ?? false,
              underline: run.underline ? {} : undefined,
            }),
        )
      : [new TextRun(block.text ?? "")];
    return new Paragraph({ children: runs });
  }

  if (block.type === "list-item") {
    const level = (block.level ?? 1) - 1;
    if (block.ordered) {
      return new Paragraph({
        children: [new TextRun(block.text ?? "")],
        numbering: { reference: "ordered-list", level },
      });
    }
    return new Paragraph({
      children: [new TextRun(block.text ?? "")],
      bullet: { level },
    });
  }

  if (block.type === "table") {
    return new Table({
      rows: (block.rows ?? []).map(
        (row) =>
          new TableRow({
            children: row.map(
              (cell) =>
                new TableCell({
                  children: [new Paragraph(String(cell))],
                }),
            ),
          }),
      ),
    });
  }

  return new Paragraph({ children: [new TextRun(block.text ?? "")] });
};

const loadJsonData = async (file) => {
  if (!file || !file.name.toLowerCase().endsWith(".json")) {
    jsonStatusEl.textContent = "Please select a .json file.";
    return;
  }

  try {
    const text = await file.text();
    const parsed = JSON.parse(text);

    if (!Array.isArray(parsed.content)) {
      throw new Error('Invalid format: "content" array not found.');
    }

    loadedJson = parsed;
    jsonInputFileName = file.name;
    jsonSelectedFileEl.textContent = file.name;
    jsonBlockCountEl.textContent = String(parsed.content.length);
    jsonStatusEl.textContent = "Ready to generate.";
    jsonInputPreview.textContent = JSON.stringify(parsed, null, 2);
    generateButton.disabled = false;
  } catch (err) {
    loadedJson = null;
    generateButton.disabled = true;
    jsonStatusEl.textContent = err instanceof Error ? err.message : "Failed to load JSON.";
    jsonInputPreview.textContent = "{}";
  }
};

const generateDocx = async () => {
  if (!loadedJson) return;

  jsonStatusEl.textContent = "Generating...";
  generateButton.disabled = true;

  try {
    const children = loadedJson.content.map(blockToDocxElement);

    const doc = new Document({
      numbering: {
        config: [
          {
            reference: "ordered-list",
            levels: Array.from({ length: 9 }, (_, i) => ({
              level: i,
              format: "decimal",
              text: `%${i + 1}.`,
              alignment: AlignmentType.LEFT,
              style: {
                paragraph: {
                  indent: { left: 720 * (i + 1), hanging: 260 },
                },
              },
            })),
          },
        ],
      },
      sections: [{ children }],
    });

    const blob = await Packer.toBlob(doc);
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = (jsonInputFileName?.replace(/\.json$/i, "") ?? "document") + ".docx";
    anchor.click();
    URL.revokeObjectURL(url);

    jsonStatusEl.textContent = "Generated successfully.";
  } catch (err) {
    jsonStatusEl.textContent = err instanceof Error ? err.message : "Generation failed.";
  } finally {
    generateButton.disabled = false;
  }
};

jsonFileInput.addEventListener("change", (event) => {
  const [file] = event.target.files ?? [];
  if (file) void loadJsonData(file);
});

generateButton.addEventListener("click", () => {
  void generateDocx();
});

jsonDropzone.addEventListener("dragover", (event) => {
  event.preventDefault();
  jsonDropzone.classList.add("dragging");
});

jsonDropzone.addEventListener("dragleave", () => {
  jsonDropzone.classList.remove("dragging");
});

jsonDropzone.addEventListener("drop", (event) => {
  event.preventDefault();
  jsonDropzone.classList.remove("dragging");
  const [file] = event.dataTransfer?.files ?? [];
  if (file) {
    jsonFileInput.files = event.dataTransfer?.files ?? null;
    void loadJsonData(file);
  }
});
