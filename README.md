# JSON Converters

A monorepo of browser-based tools for converting Word documents to JSON and JSON to Excel.
All processing happens entirely in the browser — no files are uploaded to any server.

## Tools

| Directory | Tool | Description |
|---|---|---|
| [`doc-to-json/`](./doc-to-json/) | Word to JSON | Convert `.docx` files to JSON |
| [`json-sheet-converter/`](./json-sheet-converter/) | JSON-Sheet Converter | Convert JSON files to Excel (.xlsx) |

## Combined Workflow

The two tools are designed to work together end-to-end.

```
Word (.docx)  →  [doc-to-json]  →  JSON  →  [json-sheet-converter]  →  Excel (.xlsx)
```

When you load the JSON output from `doc-to-json` into `json-sheet-converter`, it automatically generates dedicated `content_blocks` and `tables` sheets.

## Tech Stack

- **Build tool**: [Vite](https://vite.dev/) v7
- **Language**: Vanilla JavaScript (ESM)
- **Word parsing**: [mammoth.js](https://github.com/mwilliamson/mammoth.js) (`doc-to-json`)
- **Excel generation**: [SheetJS (xlsx)](https://sheetjs.com/) (`json-sheet-converter`)

## Development Setup

Navigate into each directory, install dependencies, and start the dev server.

```bash
# doc-to-json
cd doc-to-json
npm install
npm run dev

# json-sheet-converter
cd json-sheet-converter
npm install
npm run dev
```

## Build

```bash
npm run build    # Output production files to dist/
npm run preview  # Preview the production build locally
```
