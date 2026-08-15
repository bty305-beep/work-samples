#!/usr/bin/env python3
"""
設定駆動のデータ収集ツール。

収集対象・取得項目・整形ルールを config.json に外出ししてあります。
別サイトを収集したい場合、Python のコードは触らずに config.json だけ差し替えます。

使い方:
    python collect.py                     # config.json を読んで実行
    python collect.py --config other.json # 別の設定で実行
    python collect.py --dry-run           # 取得せず、設定の妥当性だけ確認
"""

import argparse
import csv
import json
import re
import sys
import time
from collections import Counter
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin, urlparse

import requests
from bs4 import BeautifulSoup

# --------------------------------------------------------------------------
# 整形ルール。config.json の "clean" で名前を指定して呼び出します。
# --------------------------------------------------------------------------

def clean_text(value):
    return re.sub(r"\s+", " ", value or "").strip()


def clean_number(value):
    digits = re.sub(r"[^\d.-]", "", value or "")
    if digits in ("", "-", "."):
        return None
    return float(digits) if "." in digits else int(digits)


def clean_date(value):
    text = clean_text(value)
    m = re.search(r"(\d{4})\D+(\d{1,2})\D+(\d{1,2})", text)
    if not m:
        return None
    y, mo, d = (int(g) for g in m.groups())
    return f"{y:04d}-{mo:02d}-{d:02d}"


CLEANERS = {"text": clean_text, "number": clean_number, "date": clean_date}


# --------------------------------------------------------------------------
# 取得
# --------------------------------------------------------------------------

def load_html(source, session, timeout, user_agent):
    """http(s) の URL でもローカルファイルでも読めるようにしてあります。"""
    if urlparse(source).scheme in ("http", "https"):
        res = session.get(source, timeout=timeout, headers={"User-Agent": user_agent})
        res.raise_for_status()
        res.encoding = res.apparent_encoding or res.encoding
        return res.text
    return Path(source).read_text(encoding="utf-8")


def extract_rows(html, base_url, rules):
    soup = BeautifulSoup(html, "html.parser")
    rows = []
    for node in soup.select(rules["row_selector"]):
        row = {}
        for field in rules["fields"]:
            target = node.select_one(field["selector"]) if field.get("selector") else node
            if target is None:
                row[field["name"]] = None
                continue

            attr = field.get("attr")
            raw = target.get(attr, "") if attr else target.get_text()

            if attr in ("href", "src") and raw:
                raw = urljoin(base_url, raw)

            cleaner = CLEANERS.get(field.get("clean", "text"), clean_text)
            row[field["name"]] = cleaner(raw)
        rows.append(row)
    return rows


# --------------------------------------------------------------------------
# 出力
# --------------------------------------------------------------------------

def write_csv(rows, columns, path):
    path.parent.mkdir(parents=True, exist_ok=True)
    # Excel でそのまま開けるよう BOM 付き UTF-8 で書き出します。
    with path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=columns)
        writer.writeheader()
        writer.writerows(rows)


def build_report(rows, config):
    """納品時に添える集計サマリー。数字を見る側が判断に使える形にしています。"""
    lines = []
    lines.append(f"# {config['job_name']} 収集レポート")
    lines.append("")
    lines.append(f"- 実行日時: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    lines.append(f"- 取得件数: {len(rows)}件")
    lines.append(f"- 収集元: {len(config['sources'])}ページ")
    lines.append("")

    for spec in config.get("report", {}).get("breakdown", []):
        name = spec["field"]
        counts = Counter(r.get(name) for r in rows if r.get(name) not in (None, ""))
        if not counts:
            continue
        lines.append(f"## {spec.get('label', name)}の内訳")
        lines.append("")
        for value, count in counts.most_common(spec.get("top", 10)):
            share = count / len(rows) * 100
            lines.append(f"- {value}: {count}件 ({share:.1f}%)")
        lines.append("")

    for name in config.get("report", {}).get("stats", []):
        values = [r[name] for r in rows if isinstance(r.get(name), (int, float))]
        if not values:
            continue
        values.sort()
        mid = values[len(values) // 2]
        lines.append(f"## {name} の分布")
        lines.append("")
        lines.append(f"- 最小: {min(values):,}")
        lines.append(f"- 中央: {mid:,}")
        lines.append(f"- 最大: {max(values):,}")
        lines.append(f"- 平均: {sum(values) / len(values):,.1f}")
        lines.append("")

    missing = {
        f["name"]: sum(1 for r in rows if r.get(f["name"]) in (None, ""))
        for f in config["rules"]["fields"]
    }
    incomplete = {k: v for k, v in missing.items() if v}
    if incomplete:
        lines.append("## 取得できなかった項目")
        lines.append("")
        for name, count in incomplete.items():
            lines.append(f"- {name}: {count}件が空")
        lines.append("")
        lines.append("空欄が多い項目は、対象ページ側の構造が変わっている可能性があります。")
        lines.append("")

    return "\n".join(lines)


# --------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="設定駆動のデータ収集ツール")
    parser.add_argument("--config", default="config.json")
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    config = json.loads(Path(args.config).read_text(encoding="utf-8"))

    required = ["job_name", "sources", "rules", "output"]
    for key in required:
        if key not in config:
            sys.exit(f"設定に {key} がありません。config.json を確認してください。")

    columns = [f["name"] for f in config["rules"]["fields"]]
    print(f"{config['job_name']}: {len(config['sources'])}ページから {len(columns)}項目を取得します。")

    if args.dry_run:
        print("設定は読み込めました。--dry-run のため取得は行いません。")
        return

    session = requests.Session()
    wait = config.get("wait_seconds", 1.0)
    timeout = config.get("timeout_seconds", 20)
    user_agent = config.get("user_agent", "Mozilla/5.0 (compatible; data-collector/1.0)")

    rows, failures = [], []
    for i, source in enumerate(config["sources"], 1):
        try:
            html = load_html(source, session, timeout, user_agent)
            found = extract_rows(html, source, config["rules"])
            rows.extend(found)
            print(f"  [{i}/{len(config['sources'])}] {source} → {len(found)}件")
        except Exception as e:
            failures.append((source, str(e)))
            print(f"  [{i}/{len(config['sources'])}] {source} → 失敗: {e}")
        if i < len(config["sources"]):
            time.sleep(wait)

    if config.get("dedupe_by"):
        key = config["dedupe_by"]
        seen, unique = set(), []
        for r in rows:
            if r.get(key) in seen:
                continue
            seen.add(r.get(key))
            unique.append(r)
        print(f"重複を除外: {len(rows)}件 → {len(unique)}件")
        rows = unique

    csv_path = Path(config["output"]["csv"])
    write_csv(rows, columns, csv_path)
    print(f"CSVを書き出しました: {csv_path}")

    if config["output"].get("report"):
        report_path = Path(config["output"]["report"])
        report_path.parent.mkdir(parents=True, exist_ok=True)
        report_path.write_text(build_report(rows, config), encoding="utf-8")
        print(f"レポートを書き出しました: {report_path}")

    if failures:
        print(f"\n{len(failures)}ページで取得に失敗しています:")
        for source, reason in failures:
            print(f"  - {source}: {reason}")


if __name__ == "__main__":
    main()
