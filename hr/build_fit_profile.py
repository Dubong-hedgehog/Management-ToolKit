"""
build_fit_profile.py
"이 회사에 잘 맞는 사람은 어떤 패턴을 보이는가"를 과거 재직자 이력서에서
뽑아내 `company_fit_profile.json`으로 저장하는 도구. resume_fit_report.py가
신규 지원자를 평가할 때 이 프로파일을 참고자료로 사용한다.

장기재직군(재직중이고 근속 12개월 이상 - 핏이 맞았을 가능성이 높다고 보는
긍정 패턴)과 조기퇴사군(퇴사했고 근속 6개월 이하 - 핏이 안 맞았을 가능성이
높다고 보는 위험 패턴) 둘 다를 학습 대상으로 삼는다. 긍정 패턴만 보면
"이미 뽑힌 사람들은 원래 비슷비슷하다"는 생존편향에 빠지기 쉬워서, 조기
퇴사 사례의 공통점도 대비군으로 함께 본다.

※ 중요: "재직 중 = 핏이 좋음", "조기퇴사 = 핏이 안 좋음"은 편의상의 근사치
일 뿐 확정적 인과관계가 아니다(그 사람 개인 문제가 아니라 조직/상사/처우
문제로 퇴사했을 수도 있음). 이 프로파일은 참고 신호로만 쓰고, 채용 여부를
이 결과 하나로 최종 결정하면 안 된다. 카테고리/키워드도 성별·나이·출신학교
같은 인적 속성은 다루지 않고 "이력서에 서술된 행동/경험"만 본다
(hr/fit_taxonomy.py 참고).

Claude API 키(.env의 ANTHROPIC_API_KEY)가 설정돼 있으면 Claude가 이력서
텍스트를 직접 읽고 패턴을 요약하고, 없으면 hr/fit_taxonomy.py의 키워드
사전으로 규칙 기반 매칭을 한다 - 정교함은 떨어지지만 항상 결과가 나온다.

Usage:
    python hr/build_fit_profile.py

Input:
    hr/sample_data/past_hires_resumes.csv
        사번, 이름, 포지션, 입사일, 재직상태(재직중/퇴사), 퇴사일, 근속개월수,
        그룹(장기재직(긍정패턴)/조기퇴사(위험패턴)/기타), 이력서요약

Output:
    hr/output/company_fit_profile.json  (resume_fit_report.py의 입력)
    hr/output/fit_profile_report.xlsx   (그룹별 카테고리 점수 표)
    hr/output/fit_category_comparison.png (긍정군 vs 위험군 카테고리별 점수 비교)
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timezone
from pathlib import Path

import matplotlib.pyplot as plt
import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from common.chart_style import save_chart, setup_style  # noqa: E402
from common.excel_io import load_csv, save_excel_report  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.llm_utils import call_claude, extract_json  # noqa: E402
from common.llm_utils import is_configured as llm_is_configured  # noqa: E402
from fit_taxonomy import CATEGORIES, CATEGORY_KEYWORDS  # noqa: E402

HERE = Path(__file__).resolve().parent
POSITIVE_GROUP = "장기재직(긍정패턴)"
RISK_GROUP = "조기퇴사(위험패턴)"


def load_resumes(path) -> pd.DataFrame:
    return load_csv(path)


# ------------------------------------------------------------------
# 규칙 기반 폴백
# ------------------------------------------------------------------
def _keyword_scores(resumes: list[str]) -> dict[str, float]:
    """카테고리별로, 해당 카테고리 키워드가 하나라도 들어간 이력서 비율(%)."""
    scores = {}
    for cat, keywords in CATEGORY_KEYWORDS.items():
        hits = sum(any(kw in r for kw in keywords) for r in resumes)
        scores[cat] = round(hits / len(resumes) * 100, 1) if resumes else 0.0
    return scores


def build_profile_fallback(positive_resumes: list[str], risk_resumes: list[str]) -> dict:
    return {
        "method": "keyword_fallback",
        "categories": {
            cat: {
                "positive_score": _keyword_scores(positive_resumes)[cat],
                "risk_score": _keyword_scores(risk_resumes)[cat],
                "description": CATEGORIES[cat],
            }
            for cat in CATEGORIES
        },
        "narrative_summary": (
            "ANTHROPIC_API_KEY가 설정되지 않아 키워드 매칭 기반으로 생성된 참고용 "
            "프로파일입니다. 정교한 정성 분석은 API 연동 후 다시 생성해보세요."
        ),
    }


# ------------------------------------------------------------------
# Claude 기반
# ------------------------------------------------------------------
def build_profile_llm(positive_resumes: list[str], risk_resumes: list[str]) -> dict:
    category_list = "\n".join(f"- {k}: {v}" for k, v in CATEGORIES.items())
    prompt = f"""다음은 한 회사의 과거 채용 데이터입니다. 두 그룹의 이력서 요약이 주어집니다.

[그룹 A: 장기재직 - 입사 후 오래 근무 중인 사람들의 이력서]
{chr(10).join(f"{i+1}. {r}" for i, r in enumerate(positive_resumes))}

[그룹 B: 조기퇴사 - 입사 6개월 이내 퇴사한 사람들의 이력서]
{chr(10).join(f"{i+1}. {r}" for i, r in enumerate(risk_resumes))}

아래 8개 카테고리 기준으로, 그룹 A에서 공통적으로 관찰되는 패턴(positive_signal)과
그룹 B에서 공통적으로 관찰되는 패턴(risk_signal)을 요약하고, 각 그룹에서 그
카테고리가 얼마나 두드러지는지 0~100 점수(positive_score, risk_score)로
매겨주세요. 성별·나이·출신학교 등 인적 속성은 절대 근거로 쓰지 말고,
이력서에 서술된 행동/경험만 근거로 삼으세요.

카테고리:
{category_list}

아래 JSON 형식으로만 답하세요 (다른 설명 없이):
{{
  "categories": {{
    "<카테고리키>": {{"positive_signal": "...", "risk_signal": "...", "positive_score": 0, "risk_score": 0}},
    ...
  }},
  "narrative_summary": "전체 종합 요약 3~4문장"
}}"""
    response = call_claude(prompt, max_tokens=3000)
    parsed = extract_json(response)
    parsed["method"] = "claude_llm"
    return parsed


# ------------------------------------------------------------------
def plot_comparison(profile: dict, output_path) -> Path:
    setup_style()
    cats = list(profile["categories"].keys())
    pos = [profile["categories"][c]["positive_score"] for c in cats]
    risk = [profile["categories"][c]["risk_score"] for c in cats]

    fig, ax = plt.subplots(figsize=(11, 6))
    x = range(len(cats))
    ax.bar([i - 0.2 for i in x], pos, width=0.4, label="장기재직(긍정패턴)", color="#2E5EAA")
    ax.bar([i + 0.2 for i in x], risk, width=0.4, label="조기퇴사(위험패턴)", color="#C0102A")
    ax.set_xticks(list(x))
    ax.set_xticklabels(cats, rotation=25, ha="right", fontsize=9)
    ax.set_ylabel("점수 / 언급 비율 (%)")
    ax.set_title("카테고리별 장기재직군 vs 조기퇴사군 비교")
    ax.legend(fontsize=9)
    fig.tight_layout()
    return save_chart(fig, output_path)


def print_summary(profile: dict) -> None:
    print_section(f"회사 핏 프로파일 생성 완료 (method={profile['method']})")
    for cat, info in profile["categories"].items():
        print(f"\n[{cat}] 긍정군 {info['positive_score']} / 위험군 {info['risk_score']}")
        if "positive_signal" in info:
            print(f"  긍정 신호: {info['positive_signal']}")
            print(f"  위험 신호: {info['risk_signal']}")
    print(f"\n종합: {profile['narrative_summary']}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default=HERE / "sample_data" / "past_hires_resumes.csv")
    parser.add_argument("--force-fallback", action="store_true", help="API 키가 있어도 강제로 키워드 폴백 사용")
    args = parser.parse_args()

    df = load_resumes(args.input)
    positive_resumes = df[df["그룹"] == POSITIVE_GROUP]["이력서요약"].tolist()
    risk_resumes = df[df["그룹"] == RISK_GROUP]["이력서요약"].tolist()

    if not positive_resumes or not risk_resumes:
        raise SystemExit("장기재직/조기퇴사 그룹 데이터가 부족합니다. 샘플 데이터를 확인하세요.")

    use_llm = llm_is_configured() and not args.force_fallback
    if use_llm:
        try:
            profile = build_profile_llm(positive_resumes, risk_resumes)
        except Exception as e:  # noqa: BLE001
            print(f"[경고] Claude 호출 실패({e}), 키워드 폴백으로 전환합니다.")
            profile = build_profile_fallback(positive_resumes, risk_resumes)
    else:
        profile = build_profile_fallback(positive_resumes, risk_resumes)

    profile["generated_at"] = datetime.now(timezone.utc).isoformat()
    profile["source"] = "hr/sample_data/past_hires_resumes.csv (전부 가상 데이터)"
    profile["n_positive"] = len(positive_resumes)
    profile["n_risk"] = len(risk_resumes)

    output_json = HERE / "output" / "company_fit_profile.json"
    output_json.parent.mkdir(parents=True, exist_ok=True)
    output_json.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")

    table = pd.DataFrame([
        {"카테고리": cat, "긍정군_점수": info["positive_score"], "위험군_점수": info["risk_score"],
         "설명": info.get("description", CATEGORIES.get(cat, ""))}
        for cat, info in profile["categories"].items()
    ])
    excel_path = save_excel_report({"카테고리별_점수": table}, HERE / "output" / "fit_profile_report.xlsx")
    chart_path = plot_comparison(profile, HERE / "output" / "fit_category_comparison.png")

    print_summary(profile)
    print(f"\n핏 프로파일 저장: {output_json}")
    print(f"엑셀 리포트 저장: {excel_path}")
    print(f"차트 저장: {chart_path}")


if __name__ == "__main__":
    main()
