import "./style.css";
import * as mammoth from "mammoth/mammoth.browser";

const fileInput = document.querySelector("#file-input");
const convertButton = document.querySelector("#convert-button");
const downloadButton = document.querySelector("#download-button");
const jsonOutput = document.querySelector("#json-output");
const statusElement = document.querySelector("#status");
const selectedFileElement = document.querySelector("#selected-file");
const blockCountElement = document.querySelector("#block-count");
const dropzone = document.querySelector("#dropzone");

let selectedFile = null;
let convertedJson = null;

const setStatus = (message) => {
  statusElement.textContent = message;
};

const setSelectedFile = (file) => {
  selectedFile = file;
  selectedFileElement.textContent = file ? file.name : "未選択";
};

const updateOutput = (payload) => {
  convertedJson = payload;
  jsonOutput.textContent = JSON.stringify(payload, null, 2);
  blockCountElement.textContent = String(payload.content.length);
  downloadButton.disabled = false;
};

const resetOutput = () => {
  convertedJson = null;
  jsonOutput.textContent = "{}";
  blockCountElement.textContent = "0";
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

const ensureDocxFile = (file) => {
  if (!file) {
    throw new Error("ファイルが選択されていません。");
  }

  if (!file.name.toLowerCase().endsWith(".docx")) {
    throw new Error(".docx ファイルを選択してください。");
  }
};

const convertFile = async () => {
  try {
    ensureDocxFile(selectedFile);
    setStatus("変換中です...");
    resetOutput();

    const arrayBuffer = await selectedFile.arrayBuffer();
    const result = await mammoth.convertToHtml({ arrayBuffer });
    const payload = htmlToJson(result.value, selectedFile.name, result.messages);

    updateOutput(payload);
    setStatus("変換が完了しました。");
  } catch (error) {
    resetOutput();
    setStatus(error instanceof Error ? error.message : "変換に失敗しました。");
  }
};

const downloadJson = () => {
  if (!convertedJson) {
    return;
  }

  const blob = new Blob([JSON.stringify(convertedJson, null, 2)], {
    type: "application/json",
  });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `${selectedFile?.name.replace(/\.docx$/i, "") ?? "word-document"}.json`;
  anchor.click();
  URL.revokeObjectURL(url);
};

const handleFileSelection = (file) => {
  setSelectedFile(file);
  setStatus(file ? "変換の準備ができました。" : "ファイルを選択してください");
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
