# JSON Converters

Word ドキュメントから JSON 変換、JSON から Excel 変換を行うブラウザアプリのモノレポです。
サーバーへのアップロードは一切なく、すべての処理をブラウザ内で完結します。

## ツール一覧

| ディレクトリ | ツール名 | 機能 |
|---|---|---|
| [`doc-to-json/`](./doc-to-json/) | Word to JSON | `.docx` ファイルを JSON に変換 |
| [`json-sheet-converter/`](./json-sheet-converter/) | JSON-Sheet Converter | JSON ファイルを Excel (.xlsx) に変換 |

## 連携ワークフロー

2 つのツールはそのまま連携して使えます。

```
Word (.docx)  →  [doc-to-json]  →  JSON  →  [json-sheet-converter]  →  Excel (.xlsx)
```

`doc-to-json` が出力する JSON を `json-sheet-converter` に読み込むと、
`content_blocks` シートと `tables` シートが自動で生成されます。

## 技術スタック

- **ビルドツール**: [Vite](https://vite.dev/) v7
- **言語**: Vanilla JavaScript (ESM)
- **Word 解析**: [mammoth.js](https://github.com/mwilliamson/mammoth.js) (`doc-to-json`)
- **Excel 生成**: [SheetJS (xlsx)](https://sheetjs.com/) (`json-sheet-converter`)

## 開発環境のセットアップ

各ディレクトリに移動して依存パッケージをインストールし、開発サーバーを起動してください。

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

## ビルド

```bash
npm run build   # dist/ に成果物を出力
npm run preview # ビルド済み成果物をローカルでプレビュー
```
