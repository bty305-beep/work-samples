#!/usr/bin/env python3
"""
Excel・スプレッドシート業務の自動化キット。

毎月くり返している3つの作業を、設定ファイルを書くだけで自動化します。

    aggregate  複数のファイルを1つにまとめて、条件別に集計する
    reconcile  2つの表を突き合わせて、合わない箇所だけを洗い出す
    fill       決まった書式に流し込んで帳票を作る

使い方:
    python run.py jobs/aggregate.json
    python run.py jobs/reconcile.json
    python run.py jobs/fill.json
"""

import json
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd


# --------------------------------------------------------------------------
# 読み込み
# --------------------------------------------------------------------------

def read_table(spec):
    """Excel でも CSV でも同じように読みます。"""
    path = Path(spec["path"])
    if not path.exists():
        raise FileNotFoundError(f"ファイルが見つかりません: {path}")

    if path.suffix.lower() in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, sheet_name=spec.get("sheet", 0), header=spec.get("header_row", 0))
    else:
        df = pd.read_csv(path, encoding=spec.get("encoding", "utf-8-sig"), header=spec.get("header_row", 0))

    if spec.get("rename"):
        df = df.rename(columns=spec["rename"])

    df.columns = [str(c).strip() for c in df.columns]
    return df


def read_many(specs):
    frames = []
    for spec in specs:
        df = read_table(spec)
        if spec.get("label_column") and spec.get("label"):
            df[spec["label_column"]] = spec["label"]
        frames.append(df)
        print(f"  読み込み: {spec['path']} → {len(df)}行")
    return pd.concat(frames, ignore_index=True)


def normalize_key(series, rules):
    s = series.astype(str).str.strip()
    if rules.get("upper"):
        s = s.str.upper()
    if rules.get("strip_spaces"):
        s = s.str.replace(r"\s+", "", regex=True)
    if rules.get("digits_only"):
        s = s.str.replace(r"\D", "", regex=True)
    return s


# --------------------------------------------------------------------------
# 集計
# --------------------------------------------------------------------------

def job_aggregate(config):
    df = read_many(config["inputs"])
    group = config["group_by"]
    value = config["value_column"]

    df[value] = pd.to_numeric(df[value], errors="coerce")
    dropped = int(df[value].isna().sum())
    if dropped:
        print(f"  数値として読めない{value}が{dropped}行ありました。集計から除外します。")
    df = df.dropna(subset=[value])

    summary = (
        df.groupby(group, dropna=False)[value]
        .agg(件数="count", 合計="sum", 平均="mean", 最大="max")
        .reset_index()
        .sort_values("合計", ascending=False)
    )
    summary["平均"] = summary["平均"].round(1)

    out = Path(config["output"])
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        summary.to_excel(writer, sheet_name="集計", index=False)
        df.to_excel(writer, sheet_name="明細", index=False)

    print(f"  合計 {df[value].sum():,.0f} ／ {len(summary)}区分")
    return out, {"件数": len(df), "区分数": len(summary), "合計": float(df[value].sum())}


# --------------------------------------------------------------------------
# 突合
# --------------------------------------------------------------------------

def job_reconcile(config):
    left = read_table(config["left"])
    right = read_table(config["right"])

    lk, rk = config["left"]["key"], config["right"]["key"]
    lv, rv = config["left"]["value"], config["right"]["value"]
    rules = config.get("key_rules", {})
    tolerance = config.get("tolerance", 0)

    left["_key"] = normalize_key(left[lk], rules)
    right["_key"] = normalize_key(right[rk], rules)
    left["_val"] = pd.to_numeric(left[lv], errors="coerce")
    right["_val"] = pd.to_numeric(right[rv], errors="coerce")

    dup_l = left["_key"].duplicated().sum()
    dup_r = right["_key"].duplicated().sum()

    merged = left.merge(right, on="_key", how="outer", suffixes=("_左", "_右"), indicator=True)

    only_left = merged[merged["_merge"] == "left_only"].drop(columns=["_merge"])
    only_right = merged[merged["_merge"] == "right_only"].drop(columns=["_merge"])
    both = merged[merged["_merge"] == "both"].drop(columns=["_merge"]).copy()

    both["差額"] = (both["_val_左"].fillna(0) - both["_val_右"].fillna(0)).round(2)
    mismatch = both[both["差額"].abs() > tolerance].sort_values("差額", key=abs, ascending=False)
    matched = both[both["差額"].abs() <= tolerance]

    drop_cols = ["_key", "_val_左", "_val_右"]
    tidy = lambda d: d.drop(columns=[c for c in drop_cols if c in d.columns])

    out = Path(config["output"])
    out.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        tidy(mismatch).to_excel(writer, sheet_name="金額不一致", index=False)
        tidy(only_left).to_excel(writer, sheet_name=f"{config['left']['label']}のみ", index=False)
        tidy(only_right).to_excel(writer, sheet_name=f"{config['right']['label']}のみ", index=False)
        tidy(matched).to_excel(writer, sheet_name="一致", index=False)

    stats = {
        "左の件数": len(left),
        "右の件数": len(right),
        "一致": len(matched),
        "金額不一致": len(mismatch),
        f"{config['left']['label']}のみ": len(only_left),
        f"{config['right']['label']}のみ": len(only_right),
        "差額合計": float(mismatch["差額"].sum()) if len(mismatch) else 0.0,
        "左の重複キー": int(dup_l),
        "右の重複キー": int(dup_r),
    }
    for k, v in stats.items():
        print(f"  {k}: {v:,}" if isinstance(v, (int, float)) else f"  {k}: {v}")
    return out, stats


# --------------------------------------------------------------------------
# 転記
# --------------------------------------------------------------------------

def job_fill(config):
    df = read_table(config["input"])
    template = config["template"]

    out_dir = Path(config["output_dir"])
    out_dir.mkdir(parents=True, exist_ok=True)

    made = []
    for _, row in df.iterrows():
        text = template
        for col in df.columns:
            value = row[col]
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            if isinstance(value, (int, float)):
                value = f"{value:,}"
            text = text.replace("{" + str(col) + "}", str(value))

        name = str(config["filename"])
        for col in df.columns:
            name = name.replace("{" + str(col) + "}", str(row[col]))
        path = out_dir / name
        path.write_text(text, encoding="utf-8")
        made.append(path.name)

    print(f"  {len(made)}件を書き出しました（例: {made[0] if made else '-'}）")
    return out_dir, {"作成件数": len(made)}


# --------------------------------------------------------------------------

JOBS = {"aggregate": job_aggregate, "reconcile": job_reconcile, "fill": job_fill}


def main():
    if len(sys.argv) < 2:
        sys.exit("設定ファイルを指定してください。例: python run.py jobs/reconcile.json")

    config_path = Path(sys.argv[1])
    config = json.loads(config_path.read_text(encoding="utf-8"))

    job = config.get("job")
    if job not in JOBS:
        sys.exit(f"job には {' / '.join(JOBS)} のいずれかを指定してください。")

    print(f"{config.get('job_name', job)} を実行します。")
    started = datetime.now()
    out, stats = JOBS[job](config)
    elapsed = (datetime.now() - started).total_seconds()

    print(f"書き出しました: {out}")
    print(f"所要 {elapsed:.2f} 秒")

    if config.get("summary"):
        lines = [f"# {config.get('job_name', job)}", "", f"実行日時: {started.strftime('%Y-%m-%d %H:%M')}", ""]
        for k, v in stats.items():
            lines.append(f"- {k}: {v:,}" if isinstance(v, (int, float)) else f"- {k}: {v}")
        lines.append("")
        lines.append(f"出力: `{out}`")
        Path(config["summary"]).write_text("\n".join(lines), encoding="utf-8")
        print(f"サマリーを書き出しました: {config['summary']}")


if __name__ == "__main__":
    
