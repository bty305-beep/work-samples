#!/usr/bin/env python3
"""
動作確認用のサンプルページを作ります。

    python make_samples.py

samples/page1.html と samples/page2.html を書き出します。
config.json の row_selector（li.item）と各 selector に合わせた構造です。

収集ツールの挙動が分かるよう、次を仕込んであります。

    価格の表記    「3,280円」  → number で 3280 に揃える
    日付の表記ゆれ 「2026年8月11日」「2026/08/09」「2026.8.3」 → date で 2026-08-11 形式に揃える
    重複掲載      同じ商品が1ページ目と2ページ目の両方にある → dedupe_by: url で1件にまとめる
    項目の欠け    更新日のない行・分類のない行 → レポートの「取得できなかった項目」に出る
"""

from pathlib import Path

SAMPLES = Path(__file__).resolve().parent / "samples"

BASE = "https://example.com/products/"

# 商品名 / スラッグ / 分類 / 価格 / 更新日
# 分類と更新日は None にすると、その span ごと出力しません（取得漏れの再現）。
PAGE1 = [
    ("A4コピー用紙 5000枚",          "a4-copy-paper",   "事務用品", "3,280円",  "2026年8月11日"),
    ("インクジェットカートリッジ 黒", "ink-black",       "消耗品",   "1,980円",  "2026/08/09"),
    ("レーザープリンタ RX-200",      "printer-rx200",   "機器",     "42,800円", "2026年8月7日"),
    ("ラベルシール 24面",            "label-24",        "事務用品", "780円",    "2026.8.3"),
]

PAGE2 = [
    # 1ページ目と同じ商品。URL が同じなので dedupe_by で1件にまとまります。
    ("A4コピー用紙 5000枚",          "a4-copy-paper",   "事務用品", "3,280円",  "2026年8月11日"),
    ("折りたたみ会議テーブル",        "table-fold",      "什器",     "18,600円", "2026年8月12日"),
    ("デスクチェア メッシュ",         "chair-mesh",      "什器",     "12,400円", None),
    ("シュレッダー S-15",            "shredder-s15",    "機器",     "24,000円", "2026年8月5日"),
    ("詰め替えインク 3色セット",      "ink-refill-3",    None,       "2,150円",  "2026年8月14日"),
]

PAGE = """<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="utf-8">
<title>取扱商品リスト（{page}／2ページ）</title>
<style>
body {{ font-family: system-ui, sans-serif; margin: 2rem auto; max-width: 720px; color: #222; }}
ul.list {{ list-style: none; padding: 0; }}
li.item {{ border-bottom: 1px solid #ddd; padding: .8rem 0; }}
a.name {{ font-weight: 600; text-decoration: none; color: #1a4d8f; }}
span {{ display: inline-block; margin-right: 1rem; font-size: .9rem; color: #555; }}
span.price {{ color: #b3261e; }}
</style>
</head>
<body>
<h1>取扱商品リスト</h1>
<p>{page} / 2 ページ</p>
<ul class="list">
{items}
</ul>
<nav>
<a href="page1.html">1</a>
<a href="page2.html">2</a>
</nav>
</body>
</html>
"""

ITEM = """  <li class="item">
    <a class="name" href="{url}">{name}</a>
{spans}
  </li>"""


def render_item(name, slug, category, price, updated):
    spans = []
    if category is not None:
        spans.append(f'    <span class="cat">{category}</span>')
    spans.append(f'    <span class="price">{price}</span>')
    if updated is not None:
        spans.append(f'    <span class="upd">更新: {updated}</span>')
    return ITEM.format(url=BASE + slug, name=name, spans="\n".join(spans))


def write_page(number, items):
    html = PAGE.format(page=number, items="\n".join(render_item(*i) for i in items))
    path = SAMPLES / f"page{number}.html"
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(html, encoding="utf-8")
    print(f"  {path.name}: {len(items)}件")
    return path


def main():
    print(f"サンプルを作成します: {SAMPLES}")
    write_page(1, PAGE1)
    write_page(2, PAGE2)
    print()
    print("できました。続けて次を実行すると、CSVとレポートが出ます。")
    print("  python collect.py")


if __name__ == "__main__":
    main()
