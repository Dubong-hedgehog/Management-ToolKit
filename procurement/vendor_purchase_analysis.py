"""
vendor_purchase_analysis.py
구매처(공급업체)별 구매 패턴을 피벗해서 "이 회사, 특정 업체에 너무 몰려있는
거 아닌가?"를 숫자로 보여주는 도구. 특정 업종에 국한되지 않는 범용 구매
데이터(품목카테고리 기준)를 전제로 한다.

1) ABC 분석(파레토)
   구매처별 누적 구매액 비중으로 A(상위 ~80%)/B(~95%)/C(나머지) 등급을 매긴다.
   재고/구매관리에서 널리 쓰이는 표준 기법.

2) 구매처 집중도 - HHI(허핀달-허쉬만 지수)
   HHI = sum(구매처 점유율(%)^2). 원래는 공정거래위원회·미국 DOJ/FTC가
   기업결합 심사에서 시장 집중도를 판단할 때 쓰는 지표(<1500 경쟁적,
   1500~2500 다소 집중, >2500 고집중)인데, 여기서는 같은 계산식을 그대로
   "우리 회사가 특정 공급처에 얼마나 의존하고 있는지"를 보는 데 재사용한다.
   공급망 리스크 관리에서 실제로 흔히 쓰이는 응용이다.
   (참고: 한국 공정거래위원회 기업결합 심사기준, 미국 DOJ/FTC Horizontal
   Merger Guidelines)

3) 구매처별 월별 구매액 추이

4) 단가 급등 이상탐지
   품목카테고리 x 구매처 조합별로 최근 1개월 평균단가를 그 이전 6개월
   평균단가와 비교해서 유의미하게(예: 30% 이상) 튀는 경우를 잡아낸다.

Usage:
    python procurement/vendor_purchase_analysis.py
    python procurement/vendor_purchase_analysis.py --as-of 2026-07-29

Input:
    procurement/sample_data/purchases.csv
        발주번호, 발주일자, 공급업체, 품목카테고리, 수량, 단가, 금액

Output:
    procurement/output/vendor_purchase_report.xlsx
        - 구매처_ABC등급 / 구매처_HHI / 단가급등_이상탐지
    procurement/output/vendor_pareto.png       (파레토 차트)
    procurement/output/vendor_monthly_trend.png (구매처별 월별 추이)
    procurement/output/price_spike.png          (단가 급등 이상탐지 대상 추이)
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import PALETTE, save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402

HERE = Path(__file__).resolve().parent

PRICE_SPIKE_RATIO = 1.3     # 최근 1개월 평균단가가 이전 6개월 대비 30% 이상 오르면 이상신호
PRICE_SPIKE_MIN_BASELINE_MONTHS = 3  # 비교할 이전 구간이 최소 이만큼은 있어야 신뢰

HHI_UNCONCENTRATED = 1500
HHI_MODERATE = 2500


def load_purchases(path) -> pd.DataFrame:
    df = load_csv(path)
    df["발주일자"] = pd.to_datetime(df["발주일자"])
    return df


# ------------------------------------------------------------------
# 1) ABC 분석(파레토)
# ------------------------------------------------------------------
def build_abc_analysis(df: pd.DataFrame) -> pd.DataFrame:
    by_vendor = df.groupby("공급업체")["금액"].sum().sort_values(ascending=False).reset_index()
    total = by_vendor["금액"].sum()
    by_vendor["구매비중"] = by_vendor["금액"] / total
    by_vendor["누적비중"] = by_vendor["구매비중"].cumsum()

    def _grade(cum):
        if cum <= 0.80:
            return "A"
        if cum <= 0.95:
            return "B"
        return "C"

    by_vendor["등급"] = by_vendor["누적비중"].map(_grade)
    return by_vendor


# ------------------------------------------------------------------
# 2) HHI 집중도
# ------------------------------------------------------------------
def compute_hhi(df: pd.DataFrame) -> tuple[float, str, pd.DataFrame]:
    by_vendor = df.groupby("공급업체")["금액"].sum()
    share_pct = (by_vendor / by_vendor.sum()) * 100
    hhi = float((share_pct ** 2).sum())

    if hhi < HHI_UNCONCENTRATED:
        level = "경쟁적(낮은 집중도)"
    elif hhi < HHI_MODERATE:
        level = "다소 집중"
    else:
        level = "고집중(공급처 의존 리스크 높음)"

    detail = share_pct.sort_values(ascending=False).reset_index()
    detail.columns = ["공급업체", "구매비중(%)"]
    return hhi, level, detail


# ------------------------------------------------------------------
# 3) 월별 구매처 추이 (차트 전용 데이터)
# ------------------------------------------------------------------
def monthly_vendor_pivot(df: pd.DataFrame) -> pd.DataFrame:
    m = df.copy()
    m["월"] = m["발주일자"].dt.to_period("M").astype(str)
    return m.groupby(["공급업체", "월"])["금액"].sum().unstack(0).fillna(0)


# ------------------------------------------------------------------
# 4) 단가 급등 이상탐지
# ------------------------------------------------------------------
def detect_price_spikes(df: pd.DataFrame, as_of: pd.Timestamp) -> pd.DataFrame:
    m = df.copy()
    m["월"] = m["발주일자"].dt.to_period("M")
    recent_month = as_of.to_period("M")
    baseline_end = recent_month - 1
    baseline_start = recent_month - 6

    rows = []
    for (vendor, cat), g in m.groupby(["공급업체", "품목카테고리"]):
        recent = g[g["월"] == recent_month]
        baseline = g[(g["월"] >= baseline_start) & (g["월"] <= baseline_end)]
        if len(recent) < 3 or baseline["월"].nunique() < PRICE_SPIKE_MIN_BASELINE_MONTHS:
            continue
        recent_avg = recent["단가"].mean()
        baseline_avg = baseline["단가"].mean()
        if baseline_avg <= 0:
            continue
        ratio = recent_avg / baseline_avg
        if ratio >= PRICE_SPIKE_RATIO:
            rows.append({
                "공급업체": vendor, "품목카테고리": cat,
                "이전평균단가": round(baseline_avg, -1),
                f"{recent_month}평균단가": round(recent_avg, -1),
                "상승률": round(ratio - 1, 3),
                "설명": f"{vendor}/{cat} 단가 {baseline_avg:,.0f}원 -> {recent_avg:,.0f}원 ({(ratio-1):.0%} 상승)",
            })
    return pd.DataFrame(rows).sort_values("상승률", ascending=False) if rows else pd.DataFrame(
        columns=["공급업체", "품목카테고리", "이전평균단가", "상승률", "설명"]
    )


# ------------------------------------------------------------------
# 차트
# ------------------------------------------------------------------
def plot_pareto(abc_df: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, ax1 = plt.subplots(figsize=(9, 5))
    colors = {"A": "#2E5EAA", "B": "#F2B705", "C": "#999999"}
    bar_colors = [colors[g] for g in abc_df["등급"]]
    ax1.bar(abc_df["공급업체"], abc_df["금액"], color=bar_colors)
    ax1.set_ylabel("구매액 (원)")
    ax1.tick_params(axis="x", rotation=45)

    ax2 = ax1.twinx()
    ax2.plot(abc_df["공급업체"], abc_df["누적비중"] * 100, color="#C0102A", marker="o")
    ax2.set_ylabel("누적 비중 (%)")
    ax2.set_ylim(0, 105)
    ax2.axhline(80, color="#2E5EAA", linestyle="--", linewidth=1)
    ax2.axhline(95, color="#F2B705", linestyle="--", linewidth=1)

    ax1.set_title("구매처별 파레토 분석 (막대=구매액, 선=누적비중, 파랑=A/노랑=B/회색=C)")
    fig.tight_layout()
    return save_chart(fig, output_path)


def plot_vendor_monthly_trend(pivot: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, ax = plt.subplots(figsize=(10, 5.5))
    top_vendors = pivot.sum().sort_values(ascending=False).index[:6]
    for i, vendor in enumerate(top_vendors):
        ax.plot(pivot.index, pivot[vendor], marker="o", label=vendor, color=PALETTE[i % len(PALETTE)])
    ax.set_title("구매처별 월별 구매액 추이 (상위 6개 업체)")
    ax.set_ylabel("구매액 (원)")
    ax.tick_params(axis="x", rotation=60)
    ax.legend(fontsize=8)
    fig.tight_layout()
    return save_chart(fig, output_path)


def plot_price_spike(df: pd.DataFrame, spike_df: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, ax = plt.subplots(figsize=(9, 5))
    if spike_df.empty:
        ax.text(0.5, 0.5, "단가 급등 이상신호 없음", ha="center", va="center", transform=ax.transAxes)
    else:
        m = df.copy()
        m["월"] = m["발주일자"].dt.to_period("M").astype(str)
        for i, (_, r) in enumerate(spike_df.iterrows()):
            g = m[(m["공급업체"] == r["공급업체"]) & (m["품목카테고리"] == r["품목카테고리"])]
            trend = g.groupby("월")["단가"].mean()
            ax.plot(trend.index, trend.values, marker="o",
                    label=f"{r['공급업체']}/{r['품목카테고리']}", color=PALETTE[i % len(PALETTE)])
        ax.tick_params(axis="x", rotation=60)
        ax.legend(fontsize=8)
    ax.set_title("단가 급등 이상탐지 대상 추이")
    ax.set_ylabel("평균 단가 (원)")
    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(abc_df, hhi, hhi_level, spike_df, as_of) -> None:
    print_section(f"구매처 패턴 분석 ({as_of.date()} 기준)")

    print("\n[1] 구매처 ABC 등급")
    for _, r in abc_df.iterrows():
        print(f"  [{r['등급']}] {r['공급업체']}: {r['금액']:,.0f}원 (비중 {r['구매비중']:.1%}, 누적 {r['누적비중']:.1%})")

    print(f"\n[2] 구매처 집중도(HHI) = {hhi:.0f} -> {hhi_level}")
    print(f"  (기준: <{HHI_UNCONCENTRATED} 경쟁적 / {HHI_UNCONCENTRATED}~{HHI_MODERATE} 다소집중 / >{HHI_MODERATE} 고집중)")

    print("\n[3] 단가 급등 이상탐지")
    if spike_df.empty:
        print("  해당 없음")
    else:
        for _, r in spike_df.iterrows():
            print(f"  - {r['설명']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "purchases.csv")
    parser.add_argument("--as-of", default=None, help="YYYY-MM-DD (기본값: 데이터의 최신 날짜)")
    args = parser.parse_args()

    df = load_purchases(args.input)
    as_of = pd.Timestamp(args.as_of) if args.as_of else df["발주일자"].max()

    abc_df = build_abc_analysis(df)
    hhi, hhi_level, hhi_detail = compute_hhi(df)
    pivot = monthly_vendor_pivot(df)
    spike_df = detect_price_spikes(df, as_of)

    excel_path = save_excel_report({
        "구매처_ABC등급": abc_df,
        "구매처_HHI": hhi_detail.assign(HHI=round(hhi, 1), 집중도평가=hhi_level),
        "단가급등_이상탐지": spike_df,
    }, HERE / "output" / "vendor_purchase_report.xlsx")

    pareto_path = plot_pareto(abc_df, HERE / "output" / "vendor_pareto.png")
    trend_path = plot_vendor_monthly_trend(pivot, HERE / "output" / "vendor_monthly_trend.png")
    spike_path = plot_price_spike(df, spike_df, HERE / "output" / "price_spike.png")

    print_summary(abc_df, hhi, hhi_level, spike_df, as_of)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {pareto_path}, {trend_path}, {spike_path}")


if __name__ == "__main__":
    main()
