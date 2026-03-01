# doc-to-json — Word to JSON

A single-page browser app that converts `.docx` files to JSON.
No files are uploaded to any server — all processing happens in the browser.

## Features

- Select a `.docx` file by clicking or drag & drop
- Extracts headings, paragraphs, lists, and tables as structured JSON
- Preview the conversion result directly in the browser
- Download the output as a JSON file

## Supported Elements

| Word element | JSON `type` |
|---|---|
| Headings (H1–H6) | `"heading"` |
| Paragraphs | `"paragraph"` |
| Unordered list items | `"list-item"` (`ordered: false`) |
| Ordered list items | `"list-item"` (`ordered: true`) |
| Tables | `"table"` |

Inline styles within paragraphs (bold, italic, underline) are recorded in the `runs` array.
Nested lists are represented by the `level` field.

## Output JSON Structure

```jsonc
{
  "fileName": "example.docx",
  "convertedAt": "2025-01-01T00:00:00.000Z",
  "messages": [],          // warnings returned by mammoth
  "content": [
    {
      "type": "heading",
      "level": 1,
      "text": "Title"
    },
    {
      "type": "paragraph",
      "text": "Body text",
      "runs": [
        { "text": "Normal" },
        { "text": "Bold text", "bold": true }
      ]
    },
    {
      "type": "list-item",
      "ordered": false,
      "level": 1,
      "text": "List item"
    },
    {
      "type": "table",
      "rows": [
        ["Header 1", "Header 2"],
        ["Cell 1",   "Cell 2"]
      ]
    }
  ]
}
```

## Tech Stack

- **Build tool**: [Vite](https://vite.dev/) v7
- **Language**: Vanilla JavaScript (ESM)
- **Word parsing**: [mammoth.js](https://github.com/mwilliamson/mammoth.js) v1

## Setup

```bash
npm install
npm run dev      # Start dev server
npm run build    # Build to dist/
npm run preview  # Preview production build
```

## Integration

The JSON output from this tool can be loaded directly into **[json-sheet-converter](../json-sheet-converter/)** to convert it to Excel.
