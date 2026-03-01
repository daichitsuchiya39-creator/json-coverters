# json-sheet-converter — JSON to Excel

JSON ファイルをブラウザ上で Excel (.xlsx) に変換するシングルページアプリです。
ファイルはサーバーに送信されず、すべての処理がブラウザ内で完結します。

## 機能

- JSON ファイルをクリックまたはドラッグ&ドロップで選択
- 変換前にシート構成をプレビュー
- Excel ファイル名を自由に指定してダウンロード

## JSON 構造とシートの対応

### 配列 JSON

```json
[{ "id": 1, "name": "Alice" }, { "id": 2, "name": "Bob" }]
```

→ `data` という 1 枚のシートに変換されます。

### オブジェクト JSON

```json
{
  "title": "レポート",
  "users": [{ "id": 1, "name": "Alice" }],
  "settings": { "theme": "dark" }
}
```

| シート名 | 内容 |
|---|---|
| `summary` | スカラー値 (`title` など) |
| `users` | 配列のキー → 行展開 |
| `settings` | オブジェクトのキー → 1 行でフラット化 |

### ネストしたオブジェクト・配列

ドット記法でキーをフラット化します。

```json
{ "user": { "address": { "city": "Tokyo" } } }
// → 列名: "user.address.city"
```

### doc-to-json 出力の特別対応

`doc-to-json` が出力する `{ content: [...] }` 形式を読み込むと、
以下のシートが自動で追加されます。

| シート名 | 内容 |
|---|---|
| `content_blocks` | 各ブロックの type / level / text など |
| `tables` | ドキュメント内の表データ（存在する場合） |

## 技術スタック

- **ビルドツール**: [Vite](https://vite.dev/) v7
- **言語**: Vanilla JavaScript (ESM)
- **Excel 生成**: [SheetJS (xlsx)](https://sheetjs.com/) v0.18

## セットアップ

```bash
npm install
npm run dev      # 開発サーバー起動
npm run build    # dist/ にビルド
npm run preview  # ビルド済みをプレビュー
```

## 連携

**[doc-to-json](../doc-to-json/)** が出力した JSON をそのまま入力として使えます。

```
Word (.docx)  →  doc-to-json  →  JSON  →  json-sheet-converter  →  Excel (.xlsx)
```
