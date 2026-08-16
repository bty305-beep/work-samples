# 設定駆動データ収集ツール

指定したページから必要な項目を集め、CSVと集計レポートを書き出します。

**別サイトを収集したい場合、Pythonのコードは触りません。`config.json` を差し替えるだけです。**

---

## 動かす

```bash
pip install requests beautifulsoup4
python collect.py
```

同梱のサンプルページを収集して、`out/products.csv` と `out/report.md` を書き出します。

```
取扱商品リスト収集: 2ページから 5項目を取得します。
  [1/2] samples/page1.html → 4件
  [2/2] samples/page2.html → 5件
重複を除外: 9件 → 8件
CSVを書き出しました: out/products.csv
レポートを書き出しました: out/report.md
```

---

## 設定の書き方

```json
{
  "job_name": "取扱商品リスト収集",
  "sources": ["https://example.com/list?p=1", "https://example.com/list?p=2"],
  "wait_seconds": 1.0,
  "dedupe_by": "url",
  "rules": {
    "row_selector": "li.item",
    "fields": [
      { "name": "商品名", "selector": "a.name", "clean": "text" },
      { "name": "url",    "selector": "a.name", "attr": "href", "clean": "text" },
      { "name": "価格",   "selector": "span.price", "clean": "number" },
      { "name": "更新日", "selector": "span.upd", "clean": "date" }
    ]
  },
  "report": {
    "breakdown": [{ "field": "分類", "top": 10 }],
    "stats": ["価格"]
  },
  "output": { "csv": "out/products.csv", "report": "out/report.md" }
}
```

| 項目 | 内容 |
|---|---|
| `sources` | 収集するページ。URLでもローカルファイルでも動きます |
| `row_selector` | 1件分をくくるCSSセレクタ |
| `fields[].selector` | その項目のCSSセレクタ |
| `fields[].attr` | 属性から取る場合に指定（`href`、`src`など）。相対URLは絶対URLに直します |
| `fields[].clean` | `text` / `number` / `date` のいずれか |
| `dedupe_by` | この項目が同じ行を1件にまとめます |
| `wait_seconds` | 各ページの間隔。相手サーバーへの負荷を抑えます |

---

## 整形について

- **number**：「3,280円」→ `3280`。通貨記号やカンマを外して数値にします
- **date**：「2026年8月11日」→ `2026-08-11`。表記ゆれを吸収します
- **text**：改行と連続する空白を1つに詰めます

CSVはBOM付きUTF-8なので、Excelでそのまま開いても文字化けしません。

---

## レポートについて

CSVだけを渡すと、受け取った側は結局自分で集計することになります。
そのため、納品時にそのまま読める形の集計を同時に出します。

- 取得件数と収集元ページ数
- 指定した項目の内訳（件数と構成比）
- 数値項目の最小・中央・最大・平均
- **取得できなかった項目の件数**

最後の項目が重要です。空欄が急に増えたときは、対象ページの構造が変わっている合図になります。

---

## 収集にあたって

- `wait_seconds` で必ず間隔を空けます（初期値1秒）
- 対象サイトの利用規約と `robots.txt` を確認したうえで実行してください
- 取得に失敗したページは処理を止めずに記録し、最後にまとめて表示します

---

## ファイル構成

```
collect.py          本体
config.json         設定（ここだけ差し替えます）
make_samples.py     サンプルページを作り直すスクリプト
samples/            動作確認用のサンプルページ
out/                書き出し先
```
