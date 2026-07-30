"""
contract_document_registry.py
법무팀 관점의 계약서 관리 도구. 구매/계약(procurement) 쪽 WBS 트래커가
"이행 일정이 지연되는지/언제 만료되는지"를 본다면, 이 스크립트는 계약서
"문서 그 자체"의 관리 상태를 본다 - 날짜 임박 알림과는 다른 각도다.

1) 원본(하드본) 보관 현황
   하드본을 실제로 보유하고 있는지, 위치는 어딘지, 분실/미회수 상태인
   계약이 몇 건인지. 소프트본(스캔본)조차 없는 계약은 원본을 잃어버리면
   그 계약의 존재를 증명할 방법이 없어지는 가장 위험한 경우라 별도로 뽑는다.

2) 필수조항 체크 매트릭스
   계약유형별로 손해배상상한/관할법원/비밀유지/불가항력/지식재산권귀속
   같은 표준 보호조항이 실제로 얼마나 포함돼 있는지를 계약유형 x 조항
   히트맵으로 본다. 계약유형에 따라 애초에 해당 안 되는 조항(예:
   임대차계약의 지식재산권귀속)은 "해당없음"으로 분모에서 뺀다.

Usage:
    python legal/contract_document_registry.py

Input:
    legal/sample_data/contracts_registry.csv
        계약번호, 계약명, 계약유형, 담당부서, 상대방, 체결일, 계약금액,
        하드본보관여부(Y/N), 하드본보관위치, 소프트본보유여부(Y/N),
        손해배상상한/관할법원/비밀유지/불가항력/지식재산권귀속 (각 Y/N/해당없음)

Output:
    legal/output/contract_document_report.xlsx
        - 하드본_보관현황 / 원본유실_위험목록 / 조항_포함율_요약 / 조항_누락_상세
    legal/output/hardcopy_status.png       (하드본 보관현황)
    legal/output/clause_coverage_heatmap.png (계약유형 x 조항 포함율 히트맵)
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
CLAUSES = ["손해배상상한", "관할법원", "비밀유지", "불가항력", "지식재산권귀속"]
# 계약금액이 이 이상인데 손해배상상한 조항이 없으면 특히 위험하다고 본다
HIGH_VALUE_THRESHOLD = 50_000_000


def load_contracts(path) -> pd.DataFrame:
    df = load_csv(path)
    df["체결일"] = pd.to_datetime(df["체결일"])
    return df


# ------------------------------------------------------------------
# 1) 원본(하드본) 보관 현황
# ------------------------------------------------------------------
def build_hardcopy_status(df: pd.DataFrame) -> pd.DataFrame:
    return df.groupby(["하드본보관여부", "하드본보관위치"]).size().reset_index(name="건수").sort_values(
        "건수", ascending=False
    )


def build_lost_original_risk(df: pd.DataFrame) -> pd.DataFrame:
    """하드본이 없는데 소프트본조차 없는(또는 둘 다 없는) 최고위험 계약."""
    risky = df[(df["하드본보관여부"] == "N") | (df["소프트본보유여부"] == "N")].copy()
    risky["위험도"] = "높음"
    risky.loc[(risky["하드본보관여부"] == "N") & (risky["소프트본보유여부"] == "N"), "위험도"] = "매우높음(원본 증빙 불가)"
    return risky[[
        "계약번호", "계약명", "계약유형", "담당부서", "상대방", "계약금액",
        "하드본보관여부", "하드본보관위치", "소프트본보유여부", "위험도",
    ]].sort_values("위험도")


# ------------------------------------------------------------------
# 2) 필수조항 체크 매트릭스
# ------------------------------------------------------------------
def build_clause_coverage(df: pd.DataFrame) -> pd.DataFrame:
    """계약유형 x 조항 별 포함율(%). '해당없음'은 분모에서 제외한다."""
    rows = []
    for ctype, g in df.groupby("계약유형"):
        for clause in CLAUSES:
            applicable = g[g[clause] != "해당없음"]
            if len(applicable) == 0:
                continue
            coverage = (applicable[clause] == "Y").mean()
            rows.append({"계약유형": ctype, "조항": clause, "포함율": round(coverage, 3), "대상건수": len(applicable)})
    return pd.DataFrame(rows)


def build_clause_gap_detail(df: pd.DataFrame) -> pd.DataFrame:
    """조항이 누락된 개별 계약 목록 (고액 계약 우선)."""
    rows = []
    for _, r in df.iterrows():
        for clause in CLAUSES:
            if r[clause] == "N":
                rows.append({
                    "계약번호": r["계약번호"], "계약명": r["계약명"], "계약유형": r["계약유형"],
                    "담당부서": r["담당부서"], "계약금액": r["계약금액"], "누락조항": clause,
                    "고액계약여부": "Y" if r["계약금액"] >= HIGH_VALUE_THRESHOLD else "N",
                })
    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values(["고액계약여부", "계약금액"], ascending=[False, False])


# ------------------------------------------------------------------
# 차트
# ------------------------------------------------------------------
def plot_hardcopy_status(df: pd.DataFrame, output_path) -> Path:
    setup_style()
    fig, axes = plt.subplots(1, 2, figsize=(13, 5.5))

    status_ct = df["하드본보관위치"].value_counts()
    colors = {"법무팀 캐비닛": "#2E5EAA", "지점 보관함": "#76B041", "분실": "#C0102A", "미회수(날인본 회수 전)": "#F2B705"}
    bar_colors = [colors.get(k, "#999999") for k in status_ct.index]
    axes[0].bar(status_ct.index, status_ct.values, color=bar_colors)
    axes[0].set_title("하드본(원본) 보관 위치별 현황")
    axes[0].set_ylabel("계약 건수")
    axes[0].tick_params(axis="x", rotation=15)
    for i, v in enumerate(status_ct.values):
        axes[0].text(i, v + 0.3, str(v), ha="center", fontsize=9)

    soft_ct = df["소프트본보유여부"].value_counts()
    axes[1].pie(soft_ct.values, labels=[f"보유({k})" if k == "Y" else f"미보유({k})" for k in soft_ct.index],
                autopct="%1.0f%%", colors=["#2E5EAA", "#C0102A"], textprops={"fontsize": 10})
    axes[1].set_title("소프트본(스캔본) 보유 비율")

    fig.tight_layout()
    return save_chart(fig, output_path)


def plot_clause_heatmap(coverage: pd.DataFrame, output_path) -> Path:
    setup_style()
    pivot = coverage.pivot(index="계약유형", columns="조항", values="포함율").reindex(columns=CLAUSES)

    fig, ax = plt.subplots(figsize=(10, 6))
    im = ax.imshow(pivot.values * 100, cmap="RdYlGn", aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(pivot.columns)))
    ax.set_xticklabels(pivot.columns)
    ax.set_yticks(range(len(pivot.index)))
    ax.set_yticklabels(pivot.index)
    for i in range(pivot.shape[0]):
        for j in range(pivot.shape[1]):
            val = pivot.values[i, j]
            if pd.isna(val):
                ax.text(j, i, "해당없음", ha="center", va="center", fontsize=8, color="gray")
                continue
            val_pct = val * 100
            ax.text(j, i, f"{val_pct:.0f}%", ha="center", va="center",
                     color="white" if val_pct < 40 or val_pct > 75 else "black", fontsize=10)
    ax.set_title("계약유형별 필수조항 포함율 (%)")
    fig.colorbar(im, ax=ax, label="포함율 (%)", fraction=0.046, pad=0.04)
    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(hardcopy_status, lost_risk, coverage, gap_detail) -> None:
    print_section("계약서 관리 현황")

    print("\n[1] 하드본(원본) 보관 현황")
    for _, r in hardcopy_status.iterrows():
        print(f"  - {r['하드본보관여부']} / {r['하드본보관위치']}: {r['건수']}건")

    print("\n[2] 원본 유실 위험 목록")
    if lost_risk.empty:
        print("  해당 없음")
    else:
        for _, r in lost_risk.iterrows():
            print(f"  - [{r['위험도']}] {r['계약명']}({r['계약번호']}, {r['담당부서']}) 하드본:{r['하드본보관여부']}/{r['하드본보관위치']} 소프트본:{r['소프트본보유여부']}")

    print("\n[3] 계약유형별 필수조항 포함율")
    for ctype, g in coverage.groupby("계약유형"):
        parts = ", ".join(f"{r['조항']} {r['포함율']:.0%}" for _, r in g.iterrows())
        print(f"  - {ctype}: {parts}")

    print("\n[4] 조항 누락 상세 (고액계약 우선)")
    if gap_detail.empty:
        print("  누락 없음")
    else:
        for _, r in gap_detail.head(10).iterrows():
            flag = " [고액계약]" if r["고액계약여부"] == "Y" else ""
            print(f"  - {r['계약명']}({r['계약번호']}) {r['누락조항']} 누락{flag} (계약금액 {r['계약금액']:,.0f}원)")
        if len(gap_detail) > 10:
            print(f"  ... 외 {len(gap_detail) - 10}건 (엑셀 참고)")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "contracts_registry.csv")
    args = parser.parse_args()

    df = load_contracts(args.input)

    hardcopy_status = build_hardcopy_status(df)
    lost_risk = build_lost_original_risk(df)
    coverage = build_clause_coverage(df)
    gap_detail = build_clause_gap_detail(df)

    excel_path = save_excel_report({
        "하드본_보관현황": hardcopy_status,
        "원본유실_위험목록": lost_risk,
        "조항_포함율_요약": coverage,
        "조항_누락_상세": gap_detail,
    }, HERE / "output" / "contract_document_report.xlsx")

    hardcopy_chart = plot_hardcopy_status(df, HERE / "output" / "hardcopy_status.png")
    heatmap_chart = plot_clause_heatmap(coverage, HERE / "output" / "clause_coverage_heatmap.png")

    print_summary(hardcopy_status, lost_risk, coverage, gap_detail)
    print(f"\n엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {hardcopy_chart}, {heatmap_chart}")


if __name__ == "__main__":
    main()
