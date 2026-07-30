"""
resume_fit_report.py
신규 지원자 이력서 한 장을 넣으면, build_fit_profile.py가 만들어둔 회사 핏
프로파일과 비교해서 적합도 리포트를 PDF로 뽑아주는 도구.

가장 중요한 건 "회사/업무 적합성이 어떨 것 같은지" 요약 추론과 "면접에서
뭘 중점적으로 물어봐야 하는지"이고, 카테고리별 점수 차트는 그걸 보조하는
시각 자료다(CONVENTIONS.md 9번 원칙과 동일).

Claude API 키(.env의 ANTHROPIC_API_KEY)가 있으면 Claude가 이력서 원문을
직접 읽고 정성 평가 + 맞춤 질문을 생성한다. 없으면 hr/fit_taxonomy.py의
키워드 사전으로 점수를 근사하고, 프로파일 대비 점수 차가 큰 카테고리에서
정적 질문은행(QUESTION_BANK)의 질문을 뽑아 대체한다 - 정교함은 떨어지지만
항상 리포트가 나온다.

※ 이 리포트는 참고 자료다. 최종 채용 판단은 사람이 한다 - 특히 폴백
경로(키워드 매칭)는 이력서 표현을 얼마나 "그럴듯하게 썼는지"에 좌우되기
쉬워서 그 자체로 당락을 결정하면 안 된다.

Usage:
    python hr/resume_fit_report.py
    python hr/resume_fit_report.py --resume path/to/resume.txt --candidate-name "홍길동" --position 마케팅

Input:
    hr/output/company_fit_profile.json  (build_fit_profile.py를 먼저 실행해야 함)
    지원자 이력서 텍스트 파일 (기본값: hr/sample_data/sample_candidate_resume.txt)

Output:
    hr/output/resume_fit_report_<지원자명>.pdf
    hr/output/candidate_score_chart_<지원자명>.png
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import matplotlib.pyplot as plt

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from reportlab.lib import colors  # noqa: E402
from reportlab.lib.pagesizes import A4  # noqa: E402
from reportlab.lib.styles import ParagraphStyle  # noqa: E402
from reportlab.lib.units import mm  # noqa: E402
from reportlab.platypus import (  # noqa: E402
    Image, ListFlowable, ListItem, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

from common.chart_style import save_chart, setup_style  # noqa: E402
from common.format_utils import print_section  # noqa: E402
from common.llm_utils import call_claude, extract_json  # noqa: E402
from common.llm_utils import is_configured as llm_is_configured  # noqa: E402
from common.pdf_statement import FONT_BOLD, FONT_REGULAR, _register_korean_font  # noqa: E402
from fit_taxonomy import CATEGORIES, CATEGORY_KEYWORDS, QUESTION_BANK  # noqa: E402

HERE = Path(__file__).resolve().parent
GAP_THRESHOLD = 25  # 프로파일 긍정군 점수 대비 이 이상 낮으면 "확인 필요"로 본다


def load_profile(path) -> dict:
    path = Path(path)
    if not path.exists():
        raise SystemExit(
            f"'{path}'가 없습니다. 먼저 python hr/build_fit_profile.py 를 실행해서 "
            "회사 핏 프로파일을 만들어주세요."
        )
    return json.loads(path.read_text(encoding="utf-8"))


def load_candidate_resume(path) -> str:
    path = Path(path)
    if not path.exists():
        raise SystemExit(f"이력서 파일을 찾을 수 없습니다: {path}")
    return path.read_text(encoding="utf-8")


# ------------------------------------------------------------------
# 규칙 기반 폴백
# ------------------------------------------------------------------
def score_resume_fallback(resume_text: str) -> dict[str, float]:
    scores = {}
    for cat, keywords in CATEGORY_KEYWORDS.items():
        hit = sum(kw in resume_text for kw in keywords)
        scores[cat] = round(min(hit / max(len(keywords), 1) * 130, 100), 1)  # 여유있게 스케일링
    return scores


def generate_report_fallback(profile: dict, resume_text: str) -> dict:
    candidate_scores = score_resume_fallback(resume_text)
    category_notes, focus, questions = {}, [], []

    for cat in CATEGORIES:
        pos_score = profile["categories"][cat]["positive_score"]
        cand_score = candidate_scores[cat]
        gap = pos_score - cand_score
        if gap >= GAP_THRESHOLD:
            category_notes[cat] = f"이력서에서 관련 키워드가 적게 확인됨 (프로파일 긍정군 {pos_score} vs 지원자 {cand_score})"
            focus.append(cat)
            questions.extend(QUESTION_BANK.get(cat, [])[:1])
        else:
            category_notes[cat] = f"프로파일 긍정군과 비슷한 수준 (긍정군 {pos_score} vs 지원자 {cand_score})"

    avg_pos = sum(c["positive_score"] for c in profile["categories"].values()) / len(profile["categories"])
    avg_cand = sum(candidate_scores.values()) / len(candidate_scores)
    if avg_cand >= avg_pos * 0.85:
        overall = "키워드 매칭 기준으로는 기존 장기재직군과 유사한 패턴입니다."
    elif avg_cand >= avg_pos * 0.6:
        overall = "일부 항목에서 기존 장기재직군 대비 약한 신호가 있어 면접에서 확인이 필요합니다."
    else:
        overall = "여러 항목에서 기존 장기재직군 대비 신호가 약합니다. 면접에서 중점적으로 확인해보세요."
    overall += " (※ 키워드 매칭 기반 근사치이며, API 연동 시 더 정교한 정성 평가로 대체됩니다.)"

    return {
        "method": "keyword_fallback",
        "candidate_scores": candidate_scores,
        "category_notes": category_notes,
        "overall_recommendation": overall,
        "interview_focus": focus or ["특별히 두드러지는 약점 신호는 없습니다."],
        "interview_questions": questions or ["일반적인 역량 검증 질문으로 진행하세요."],
    }


# ------------------------------------------------------------------
# Claude 기반
# ------------------------------------------------------------------
def generate_report_llm(profile: dict, resume_text: str, candidate_name: str, position: str) -> dict:
    category_list = "\n".join(f"- {k}: {v}" for k, v in CATEGORIES.items())
    profile_summary = "\n".join(
        f"- {cat}: 긍정군 {info['positive_score']}점 (신호: {info.get('positive_signal', 'N/A')}) / "
        f"위험군 {info['risk_score']}점 (신호: {info.get('risk_signal', 'N/A')})"
        for cat, info in profile["categories"].items()
    )
    prompt = f"""당신은 채용 담당자를 돕는 보조 도구입니다. 아래는 이 회사의 "핏 프로파일"
(과거 장기재직자 vs 조기퇴사자 이력서에서 뽑은 참고 패턴)과, 신규 지원자의 이력서입니다.

[회사 핏 프로파일]
{profile_summary}

[지원자 정보]
이름: {candidate_name}
지원 포지션: {position}

[지원자 이력서]
{resume_text}

카테고리:
{category_list}

지원자의 이력서를 프로파일과 비교해서 카테고리별 점수(0~100)와 짧은 근거를 매기고,
전체적인 회사/업무 적합성에 대한 종합 의견(3~4문장, 추론이라는 점을 명시)과,
면접에서 중점적으로 확인해야 할 포인트(2~4개), 그 포인트를 검증할 맞춤 면접 질문
(3~5개)을 제안하세요. 성별·나이·출신학교 등 인적 속성은 절대 언급하거나 근거로
쓰지 마세요. 이 리포트는 참고자료일 뿐 최종 채용 여부를 결정하는 게 아니라는 점을
overall_recommendation에 한 문장으로 포함하세요.

아래 JSON 형식으로만 답하세요 (다른 설명 없이):
{{
  "candidate_scores": {{"<카테고리키>": 0, ...}},
  "category_notes": {{"<카테고리키>": "근거 한 문장", ...}},
  "overall_recommendation": "종합 의견",
  "interview_focus": ["...", "..."],
  "interview_questions": ["...", "..."]
}}"""
    response = call_claude(prompt, max_tokens=2500)
    parsed = extract_json(response)
    parsed["method"] = "claude_llm"
    return parsed


# ------------------------------------------------------------------
# 차트 + PDF
# ------------------------------------------------------------------
def plot_candidate_scores(profile: dict, report: dict, output_path) -> Path:
    setup_style()
    cats = list(CATEGORIES.keys())
    pos = [profile["categories"][c]["positive_score"] for c in cats]
    cand = [report["candidate_scores"].get(c, 0) for c in cats]

    fig, ax = plt.subplots(figsize=(11, 6))
    x = range(len(cats))
    ax.bar([i - 0.2 for i in x], pos, width=0.4, label="회사 프로파일(긍정군 기준)", color="#2E5EAA")
    ax.bar([i + 0.2 for i in x], cand, width=0.4, label="지원자", color="#76B041")
    ax.set_xticks(list(x))
    ax.set_xticklabels(cats, rotation=25, ha="right", fontsize=9)
    ax.set_ylabel("점수")
    ax.set_ylim(0, 105)
    ax.set_title("지원자 vs 회사 핏 프로파일(긍정군)")
    ax.legend(fontsize=9)
    fig.tight_layout()
    return save_chart(fig, output_path)


def render_pdf(report: dict, candidate_name: str, position: str, chart_path: Path, output_path: Path) -> Path:
    _register_korean_font()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc = SimpleDocTemplate(str(output_path), pagesize=A4, topMargin=18 * mm, bottomMargin=18 * mm,
                             leftMargin=18 * mm, rightMargin=18 * mm)

    title_style = ParagraphStyle("Title", fontName=FONT_BOLD, fontSize=18, alignment=1, spaceAfter=10)
    h2_style = ParagraphStyle("H2", fontName=FONT_BOLD, fontSize=12, spaceBefore=12, spaceAfter=6)
    body_style = ParagraphStyle("Body", fontName=FONT_REGULAR, fontSize=10, leading=15)
    meta_style = ParagraphStyle("Meta", fontName=FONT_REGULAR, fontSize=9.5, alignment=1,
                                 textColor=colors.grey, spaceAfter=14)

    story = [
        Paragraph("지원자 적합도 리포트", title_style),
        Paragraph(f"지원자: {candidate_name}  |  지원 포지션: {position}  |  생성 방식: {report['method']}",
                  meta_style),
        Paragraph("종합 의견", h2_style),
        Paragraph(report["overall_recommendation"], body_style),
        Spacer(1, 8),
        Image(str(chart_path), width=170 * mm, height=170 * mm * 0.55),
        Paragraph("면접 중점 포인트", h2_style),
        ListFlowable([ListItem(Paragraph(p, body_style)) for p in report["interview_focus"]], bulletType="bullet"),
        Paragraph("맞춤 면접 질문", h2_style),
        ListFlowable([ListItem(Paragraph(q, body_style)) for q in report["interview_questions"]], bulletType="1"),
        Paragraph("카테고리별 근거", h2_style),
    ]

    table_data = [["카테고리", "점수", "근거"]]
    for cat in CATEGORIES:
        score = report["candidate_scores"].get(cat, "-")
        note = report["category_notes"].get(cat, "-")
        table_data.append([cat, str(score), Paragraph(note, body_style)])
    table = Table(table_data, colWidths=[42 * mm, 16 * mm, 116 * mm])
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), FONT_BOLD),
        ("FONTNAME", (0, 1), (-1, -1), FONT_REGULAR),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.whitesmoke),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
    ]))
    story.append(table)
    story.append(Spacer(1, 10))
    story.append(Paragraph(
        "※ 이 리포트는 샘플 데이터 기반 참고 자료입니다. 최종 채용 판단에 단독으로 사용하지 마세요.",
        ParagraphStyle("Footnote", fontName=FONT_REGULAR, fontSize=8, textColor=colors.grey),
    ))

    doc.build(story)
    return output_path


def print_summary(report: dict, candidate_name: str) -> None:
    print_section(f"{candidate_name} 적합도 리포트 (method={report['method']})")
    print(f"\n종합 의견: {report['overall_recommendation']}")
    print("\n면접 중점 포인트:")
    for p in report["interview_focus"]:
        print(f"  - {p}")
    print("\n맞춤 면접 질문:")
    for q in report["interview_questions"]:
        print(f"  - {q}")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--profile", default=HERE / "output" / "company_fit_profile.json")
    parser.add_argument("--resume", default=HERE / "sample_data" / "sample_candidate_resume.txt")
    parser.add_argument("--candidate-name", default="홍길동(샘플)")
    parser.add_argument("--position", default="마케팅")
    parser.add_argument("--force-fallback", action="store_true")
    args = parser.parse_args()

    profile = load_profile(args.profile)
    resume_text = load_candidate_resume(args.resume)

    use_llm = llm_is_configured() and not args.force_fallback
    if use_llm:
        try:
            report = generate_report_llm(profile, resume_text, args.candidate_name, args.position)
        except Exception as e:  # noqa: BLE001
            print(f"[경고] Claude 호출 실패({e}), 키워드 폴백으로 전환합니다.")
            report = generate_report_fallback(profile, resume_text)
    else:
        report = generate_report_fallback(profile, resume_text)

    safe_name = "".join(c for c in args.candidate_name if c.isalnum()) or "candidate"
    chart_path = plot_candidate_scores(profile, report, HERE / "output" / f"candidate_score_chart_{safe_name}.png")
    pdf_path = render_pdf(report, args.candidate_name, args.position, chart_path,
                           HERE / "output" / f"resume_fit_report_{safe_name}.pdf")

    print_summary(report, args.candidate_name)
    print(f"\n차트 저장: {chart_path}")
    print(f"PDF 리포트 저장: {pdf_path}")


if __name__ == "__main__":
    main()
