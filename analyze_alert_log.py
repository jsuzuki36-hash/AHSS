# analyze_alert_log.py
from pathlib import Path
import csv
from collections import Counter, defaultdict

import matplotlib.pyplot as plt

# 日本語フォント対策（Windowsで効きやすい順に候補）
plt.rcParams["font.family"] = [
    "Yu Gothic", "Yu Gothic UI", "Meiryo", "MS Gothic", "MS PGothic"
]
plt.rcParams["axes.unicode_minus"] = False  # マイナス記号の文字化け対策

CSV_NAME = "alert_log.csv"   # VBAが出力する固定名
DELIM = ";"                  # reasonsの区切り


def read_rows(csv_path: Path):
    """文字コードを複数試してCSVを読む（Windows/Excelあるある対応）"""
    encodings = ["utf-8-sig", "utf-8", "cp932", "shift_jis"]
    last_err = None
    for enc in encodings:
        try:
            with csv_path.open("r", encoding=enc, newline="") as f:
                reader = csv.DictReader(f)
                rows = list(reader)
            return rows, enc
        except Exception as e:
            last_err = e
    raise last_err


def split_reasons(reason_str: str):
    """reasons列を分解して、空や余計な空白を除外"""
    if not reason_str:
        return []
    parts = [p.strip() for p in reason_str.split(DELIM)]
    return [p for p in parts if p]


def plot_counter(title: str, counter: Counter, out_png: Path):
    """Counterを棒グラフにしてPNG保存（色指定しない）"""
    labels = list(counter.keys())
    values = list(counter.values())

    plt.figure(figsize=(12, 6))
    plt.bar(labels, values)
    plt.title(title)
    plt.ylabel("count")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close()


def plot_species_reason_stacked(title: str, sr_counts, out_png: Path, top_reasons=None):
    """
    種類×異常内訳を積み上げ棒で表示
    sr_counts: dict[(species, reason)] = count
    top_reasons: 表示する理由のリスト（Noneなら全）
    """
    # 種類一覧
    species_list = sorted({s for (s, _) in sr_counts.keys()})

    # 理由一覧（多すぎると見づらいので、必要なら上位だけ）
    if top_reasons is None:
        reason_list = sorted({r for (_, r) in sr_counts.keys()})
    else:
        reason_list = top_reasons

    # 種類ごとに理由の件数を並べる
    data = {r: [] for r in reason_list}
    for s in species_list:
        for r in reason_list:
            data[r].append(sr_counts.get((s, r), 0))

    # 積み上げ
    plt.figure(figsize=(13, 7))
    bottom = [0] * len(species_list)
    for r in reason_list:
        vals = data[r]
        plt.bar(species_list, vals, bottom=bottom, label=r)
        bottom = [b + v for b, v in zip(bottom, vals)]

    plt.title(title)
    plt.ylabel("count")
    plt.xticks(rotation=45, ha="right")
    plt.legend()
    plt.tight_layout()
    plt.savefig(out_png, dpi=150, bbox_inches="tight")
    plt.close()


def main():
    base = Path(__file__).resolve().parent
    csv_path = base / CSV_NAME
    if not csv_path.exists():
        raise FileNotFoundError(f"{csv_path} が見つかりません。VBAでCSV出力したフォルダと同じ場所に置いてください。")

    rows, enc = read_rows(csv_path)
    print(f"読み込み成功: {CSV_NAME} / encoding={enc}")
    print("総件数:", len(rows))

    # 1) 種類別件数
    species_counter = Counter(r.get("species", "").strip() for r in rows)
    if "" in species_counter:
        # 空欄があれば「(空欄)」扱い
        species_counter["(空欄)"] = species_counter.pop("")
    print("\n種類別件数:")
    for k, v in species_counter.most_common():
        print(k, v)

    # 2) 異常理由別件数（reasonsを分解）
    reason_counter = Counter()
    for r in rows:
        for reason in split_reasons(r.get("reasons", "")):
            reason_counter[reason] += 1

    print("\n異常理由別件数:")
    for k, v in reason_counter.most_common():
        print(k, v)

    # 3) 種類×異常内訳（reasons分解してカウント）
    sr_counts = defaultdict(int)
    for r in rows:
        sp = (r.get("species", "") or "").strip() or "(空欄)"
        reasons = split_reasons(r.get("reasons", ""))
        for reason in reasons:
            sr_counts[(sp, reason)] += 1

    # --- グラフ出力（PNG）
    out1 = base / "chart_species_counts.png"
    out2 = base / "chart_reason_counts.png"
    out3 = base / "chart_species_reason_stacked.png"

    plot_counter("AHSS: Species Counts", species_counter, out1)
    plot_counter("AHSS: Reason Counts (split by ';')", reason_counter, out2)

    # 理由が多い場合は上位だけにする（見やすさ優先）
    top = [k for k, _ in reason_counter.most_common(8)]
    plot_species_reason_stacked("AHSS: Species x Reasons (Top 8 Reasons)", sr_counts, out3, top_reasons=top)

    print("\nPNGを出力しました:")
    print(out1.name)
    print(out2.name)
    print(out3.name)


if __name__ == "__main__":
    main()