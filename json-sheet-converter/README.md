# json-sheet-converter — JSON to Excel

A single-page browser app that converts JSON files to Excel (.xlsx).
No files are uploaded to any server — all processing happens in the browser.

## Features

- Select a JSON file by clicking or drag & drop
- Preview the sheet layout before downloading
- Specify a custom Excel file name
- Download the output as an `.xlsx` file

## JSON Structure and Sheet Mapping

### Array JSON

```json
[{ "id": 1, "name": "Alice" }, { "id": 2, "name": "Bob" }]
```

→ Converted into a single sheet named `data`.

### Object JSON

```json
{
  "title": "Report",
  "users": [{ "id": 1, "name": "Alice" }],
  "settings": { "theme": "dark" }
}
```

| Sheet name | Content |
|---|---|
| `summary` | Scalar values (e.g. `title`) |
| `users` | Array key → rows |
| `settings` | Object key → single flattened row |

### Nested Objects and Arrays

Keys are flattened using dot notation.

```json
{ "user": { "address": { "city": "Tokyo" } } }
// → column name: "user.address.city"
```

### Special Support for doc-to-json Output

When loading a `{ content: [...] }` JSON produced by `doc-to-json`, the following sheets are automatically added:

| Sheet name | Content |
|---|---|
| `content_blocks` | type / level / text of each block |
| `tables` | Table data from the document (if any) |

## Tech Stack

- **Build tool**: [Vite](https://vite.dev/) v7
- **Language**: Vanilla JavaScript (ESM)
- **Excel generation**: [SheetJS (xlsx)](https://sheetjs.com/) v0.18

## Setup

```bash
npm install
npm run dev      # Start dev server
npm run build    # Build to dist/
npm run preview  # Preview production build
```

## Integration

This tool accepts the JSON output from **[doc-to-json](../doc-to-json/)** directly.

```
Word (.docx)  →  doc-to-json  →  JSON  →  json-sheet-converter  →  Excel (.xlsx)
```
