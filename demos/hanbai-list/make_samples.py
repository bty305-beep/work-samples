#!/usr/bin/env python3
"""
動作確認・速度計測用のサンプルExcelを作ります。

    python make_samples.py                          # 商品マスタ30万件・除外リスト5千件
    python make_samples.py --rows 500000 --excludes 8000

Excelはバイナリのためリポジトリに置かず、このスクリプトで生成する形にしています。
samples/ に2ファイルを書き出します。

    商品マスタ.xlsx    商品コード・商品名・仕入れ値
    除外リスト.xlsx    除外コード・除外理由

実運用で必ず起きることを、意図的に混ぜてあります。

    除外リストの商品コードが全角          ＡＢ－０００１２３
    除外リストの商品コードが小文字        ab-000123
    商品コードの前後に空白                "AB-000123 "（全角スペースを含む）
    商品コードの途中に空白                "AB- 000123"
    マスタに存在しない除外コード          廃番。除外リストにだけ残っている
    仕入れ値が「1,200円」「１２００」表記  数値として読めない
    仕入れ値が空欄・「－」                 価格を計算できない
    商品コードの重複                      マスタに同じコードが2行ある

生成後、そのまま generate.py を実行すれば処理時間が測れます。
"""

import argparse
import time
from pathlib import Path

import numpy as np
import pandas as pd

from generate import write_workbook

HERE = Path(__file__).resolve().parent
SAMPLES = HERE / "samples"

# 乱数を固定しておくと、作り直しても同じサンプルになります（計測を比較できます）。
SEED = 20260801

PREFIXES = ["AB", "CD", "EF", "GH", "JK", "LM", "NP", "QR"]

MATERIALS = ["ステンレス", "アルミ", "スチール", "木製", "樹脂", "ガラス", "帆布", "陶器"]
CATEGORIES = ["収納ラック", "作業台", "保存容器", "台車", "工具箱", "照明", "マット", "脚立"]
VARIANTS = ["スタンダード", "ワイド", "スリム", "折りたたみ", "キャスター付", "業務用"]
SIZES = ["S", "M", "L", "XL", "幅60", "幅90", "幅120"]

REASONS = ["取扱終了", "メーカー指定", "在庫なし", "季節終了", "型番切替"]

# 半角 → 全角の変換表（ＡＢ－０００１２３ を作るため）
TO_ZENKAKU = str.maketrans(
    {chr(c): chr(c - 0x20 + 0xFF00) for c in range(0x21, 0x7F)} | {" ": "　"}
)

# 仕入れ値の汚れ方。上2つは拾えます。下2つは拾えないので除外一覧に出します。
DIRTY_COSTS = ["1,200円", "１２００", "", "－"]


def make_master(rows, duplicates, dirty, rng):
    unique = rows - duplicates

    index = np.arange(unique)
    prefix = np.array(PREFIXES)[index % len(PREFIXES)]
    codes = [f"{p}-{i:06d}" for p, i in zip(prefix, index)]

    names = [
        f"{MATERIALS[a]}{CATEGORIES[b]} {VARIANTS[c]} {SIZES[d]}"
        for a, b, c, d in zip(
            rng.integers(0, len(MATERIALS), unique),
            rng.integers(0, len(CATEGORIES), unique),
            rng.integers(0, len(VARIANTS), unique),
            rng.integers(0, len(SIZES), unique),
        )
    ]

    # 安いものから高いものまで幅を持たせます（設定ファイルの段階掛け率を試せます）。
    costs = np.clip(np.round(rng.lognormal(7.0, 0.9, unique)), 100, 80000).astype("int64")

    df = pd.DataFrame({"商品コード": codes, "商品名": names, "仕入れ値": costs})

    # 仕入れ値の表記ゆれを混ぜます。
    if dirty:
        df["仕入れ値"] = df["仕入れ値"].astype("object")
        spots = rng.choice(unique, size=dirty, replace=False)
        for i, spot in enumerate(spots):
            df.iat[spot, 2] = DIRTY_COSTS[i % len(DIRTY_COSTS)]

    # 同じ商品コードの行を混ぜます（マスタの二重登録）。
    if duplicates:
        copies = df.iloc[rng.choice(unique, size=duplicates, replace=False)]
        df = pd.concat([df, copies], ignore_index=True)

    # 実際のマスタはコード順に並んでいるとは限りません。
    return df.sample(frac=1.0, random_state=SEED).reset_index(drop=True)


def scatter_codes(codes, rng):
    """除外リスト側の商品コードに、現場で起きる表記ゆれを混ぜます。"""
    kinds = rng.choice(
        ["そのまま", "全角", "小文字", "前後空白", "途中空白"],
        size=len(codes),
        p=[0.55, 0.15, 0.12, 0.10, 0.08],
    )
    out = []
    for code, kind in zip(codes, kinds):
        if kind == "全角":
            code = code.translate(TO_ZENKAKU)
        elif kind == "小文字":
            code = code.lower()
        elif kind == "前後空白":
            code = f" {code}　"
        elif kind == "途中空白":
            code = code.replace("-", "- ")
        out.append(code)
    return out, kinds


def make_excludes(master_codes, count, missing, rng):
    picked = list(rng.choice(master_codes, size=count - missing, replace=False))
    codes, kinds = scatter_codes(picked, rng)

    # 廃番。除外リストには残っているが、商品マスタにはもう無いコード。
    codes += [f"ZZ-{900000 + i:06d}" for i in range(missing)]
    kinds = list(kinds) + ["マスタに無い"] * missing

    df = pd.DataFrame({
        "除外コード": codes,
        "除外理由": list(rng.choice(REASONS, size=len(codes))),
    })
    return df.sample(frac=1.0, random_state=SEED).reset_index(drop=True), pd.Series(kinds)


def main():
    parser = argparse.ArgumentParser(description="販売リスト生成のサンプルExcelを作ります。")
    parser.add_argument("--rows", type=int, default=300_000, help="商品マスタの行数（既定 300000）")
    parser.add_argument("--excludes", type=int, default=5_000, help="除外リストの行数（既定 5000）")
    parser.add_argument("--duplicates", type=int, default=200, help="商品コードが重複する行数")
    parser.add_argument("--dirty", type=int, default=160, help="仕入れ値の表記がゆれている行数")
    parser.add_argument("--missing", type=int, default=120, help="マスタに存在しない除外コードの件数")
    parser.add_argument("--csv", action="store_true", help="同じ内容をCSVでも書き出す（読み込み速度の比較用）")
    args = parser.parse_args()

    rng = np.random.default_rng(SEED)
    SAMPLES.mkdir(parents=True, exist_ok=True)

    print(f"サンプルを作成します: {SAMPLES}")

    clock = time.perf_counter()
    master = make_master(args.rows, args.duplicates, args.dirty, rng)
    excludes, kinds = make_excludes(
        master["商品コード"].unique(), args.excludes, args.missing, rng
    )
    print(f"  組み立て: {time.perf_counter() - clock:.2f}秒")

    clock = time.perf_counter()
    write_workbook(SAMPLES / "商品マスタ.xlsx", [("商品マスタ", master)])
    print(f"  商品マスタ.xlsx: {len(master):,}行（{time.perf_counter() - clock:.2f}秒）")

    clock = time.perf_counter()
    write_workbook(SAMPLES / "除外リスト.xlsx", [("除外リスト", excludes)])
    print(f"  除外リスト.xlsx: {len(excludes):,}行（{time.perf_counter() - clock:.2f}秒）")

    if args.csv:
        clock = time.perf_counter()
        master.to_csv(SAMPLES / "商品マスタ.csv", index=False, encoding="utf-8-sig")
        excludes.to_csv(SAMPLES / "除外リスト.csv", index=False, encoding="utf-8-sig")
        print(f"  CSVも書き出しました（{time.perf_counter() - clock:.2f}秒）")

    print()
    print("除外リストの商品コードの内訳:")
    notes = {
        "そのまま": "",
        "マスタに無い": "  ← 廃番。除外リストにだけ残っています",
    }
    for kind, count in kinds.value_counts().items():
        note = notes.get(kind, "  ← 正規化しないと突合できません")
        print(f"  {kind}: {count:,}件{note}")

    print()
    print("できました。続けて次を実行すると、販売リストが生成されます。")
    print("  python generate.py")


if __name__ == "__main__":
    main()
