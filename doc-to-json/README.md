# doc-to-json — Word to JSON

`.docx` ファイルをブラウザ上で JSON に変換するシングルページアプリです。
ファイルはサーバーに送信されず、すべての処理がブラウザ内で完結します。

## 機能

- `.docx` ファイルをクリックまたはドラッグ&ドロップで選択
- 見出し・段落・リスト・表を構造化 JSON として抽出
- 変換結果をブラウザ上でプレビュー
- JSON ファイルとしてダウンロード

## 対応要素

| Word 要素 | JSON の `type` |
|---|---|
| 見出し (H1〜H6) | `"heading"` |
| 段落 | `"paragraph"` |
| 番号なしリスト項目 | `"list-item"` (`ordered: false`) |
| 番号付きリスト項目 | `"list-item"` (`ordered: true`) |
| 表 | `"table"` |

段落内のインラインスタイル（太字・斜体・下線）は `runs` 配列に記録されます。
ネストしたリストは `level` フィールドで階層を表現します。

## 出力 JSON の構造

```jsonc
{
  "fileName": "example.docx",
  "convertedAt": "2025-01-01T00:00:00.000Z",
  "messages": [],          // mammoth が返す警告メッセージ
  "content": [
    {
      "type": "heading",
      "level": 1,
      "text": "タイトル"
    },
    {
      "type": "paragraph",
      "text": "本文テキスト",
      "runs": [
        { "text": "通常" },
        { "text": "太字", "bold": true }
      ]
    },
    {
      "type": "list-item",
      "ordered": false,
      "level": 1,
      "text": "リスト項目"
    },
    {
      "type": "table",
      "rows": [
        ["ヘッダー1", "ヘッダー2"],
        ["セル1",    "セル2"]
      ]
    }
  ]
}
```

## 技術スタック

- **ビルドツール**: [Vite](https://vite.dev/) v7
- **言語**: Vanilla JavaScript (ESM)
- **Word 解析**: [mammoth.js](https://github.com/mwilliamson/mammoth.js) v1

## セットアップ

```bash
npm install
npm run dev      # 開発サーバー起動
npm run build    # dist/ にビルド
npm run preview  # ビルド済みをプレビュー
```

## 連携

このツールが出力する JSON は **[json-sheet-converter](../json-sheet-converter/)** でそのまま Excel に変換できます。
