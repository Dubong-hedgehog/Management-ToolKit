"""
pdf_statement.py
손익계산서/재무상태표처럼 "당기/전기 2열 비교 + 로마숫자 대분류 + 볼드 소계"
형식의 공식 재무제표 스타일 PDF를 만드는 공통 렌더러.

손익계산서, 재무상태표 둘 다 시각적 구조가 같기 때문에(제목 → 기간/회사명 →
과목·당기·전기 표) 이 모듈 하나로 두 문서를 모두 만든다. 새로운 재무제표
유형이 추가되더라도 이 렌더러는 그대로 재사용하면 된다.
"""
from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

_FONT_REGISTERED = False
FONT_REGULAR = "NanumGothic"
FONT_BOLD = "NanumGothicBold"


def _register_korean_font() -> None:
    """koreanize-matplotlib에 번들된 나눔고딕 TTF를 reportlab에 등록한다.

    reportlab 기본 폰트는 한글을 지원하지 않아서, 별도 폰트 파일을 등록하지
    않으면 한글이 전부 깨진다. 이미 requirements.txt에 있는
    koreanize-matplotlib 패키지 안의 폰트를 재사용해서 새 의존성을 늘리지 않는다.
    """
    global _FONT_REGISTERED
    if _FONT_REGISTERED:
        return

    import koreanize_matplotlib

    font_dir = Path(koreanize_matplotlib.__file__).parent / "fonts"
    pdfmetrics.registerFont(TTFont(FONT_REGULAR, str(font_dir / "NanumGothic.ttf")))
    pdfmetrics.registerFont(TTFont(FONT_BOLD, str(font_dir / "NanumGothicBold.ttf")))
    _FONT_REGISTERED = True


@dataclass
class StatementRow:
    """재무제표 표 한 줄.

    level 0 = 로마숫자 대분류 (Ⅰ, Ⅱ ...), level 1 = 중분류 (1), (2) ...,
    level 2 = 개별 계정과목.
    detail_* 는 개별 계정 금액(들여쓴 안쪽 금액열), total_* 는 그 구간의
    합계/소계 금액(바깥쪽 금액열)이다. 계정마다 둘 중 하나만 채운다.
    """
    label: str
    level: int = 2
    detail_current: Optional[float] = None
    total_current: Optional[float] = None
    detail_prior: Optional[float] = None
    total_prior: Optional[float] = None
    bold: bool = False


def _fmt(amount: Optional[float]) -> str:
    if amount is None:
        return ""
    return f"{amount:,.0f}"


def render_statement_pdf(
    rows: list[StatementRow],
    title: str,
    period_lines: list[str],
    company_name: str,
    output_path: str | Path,
    current_header: str = "당기",
    prior_header: str = "전기",
) -> Path:
    """공식 재무제표 스타일 PDF를 생성한다.

    Args:
        rows: 표에 들어갈 줄들 (StatementRow 리스트)
        title: 문서 제목 (예: "손 익 계 산 서")
        period_lines: 제목 아래 표시할 기간 안내 줄들
                      (예: ["제 8(당)기 2026년 1월 1일부터 2026년 3월 31일까지",
                            "제 7(전)기 2025년 1월 1일부터 2025년 12월 31일까지"])
        company_name: 회사명
        output_path: 저장할 PDF 경로
        current_header: 표 헤더의 당기 열 제목 (예: "제8(당)기")
        prior_header: 표 헤더의 전기 열 제목 (예: "제7(전)기")
    """
    _register_korean_font()
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    doc = SimpleDocTemplate(
        str(output_path), pagesize=A4,
        topMargin=18 * mm, bottomMargin=18 * mm, leftMargin=16 * mm, rightMargin=16 * mm,
    )

    title_style = ParagraphStyle(
        "Title", fontName=FONT_BOLD, fontSize=20, alignment=1, spaceAfter=4,
        leading=26,
    )
    period_style = ParagraphStyle("Period", fontName=FONT_REGULAR, fontSize=9.5, alignment=1, leading=14)
    meta_style = ParagraphStyle("Meta", fontName=FONT_REGULAR, fontSize=9.5, leading=13)
    meta_style_right = ParagraphStyle("MetaRight", fontName=FONT_REGULAR, fontSize=9.5, alignment=2, leading=13)

    story = []
    # 제목 (밑줄 두 줄 느낌을 내기 위해 제목 바로 아래 얇은 선을 그린 표 하나를 둔다)
    title_text = "  ".join(list(title.replace(" ", "")))  # 글자 사이 자간을 벌려서 원본 서식과 비슷하게
    story.append(Paragraph(title_text, title_style))
    underline_tbl = Table([[""]], colWidths=[45 * mm], rowHeights=[1.6])
    underline_tbl.setStyle(TableStyle([
        ("LINEBELOW", (0, 0), (-1, -1), 1.6, colors.black),
        ("LINEABOVE", (0, 0), (-1, -1), 0.4, colors.black),
    ]))
    centered_underline = Table([[underline_tbl]], colWidths=[doc.width])
    centered_underline.setStyle(TableStyle([("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    story.append(centered_underline)
    story.append(Spacer(1, 6))

    for line in period_lines:
        story.append(Paragraph(line, period_style))
    story.append(Spacer(1, 6))

    meta_tbl = Table(
        [[Paragraph(f"회사명 : {company_name}", meta_style), Paragraph("(단위 : 원)", meta_style_right)]],
        colWidths=[doc.width * 0.7, doc.width * 0.3],
    )
    story.append(meta_tbl)
    story.append(Spacer(1, 4))

    # ---- 본문 표 ----
    label_w = doc.width * 0.34
    amt_w = (doc.width - label_w) / 4

    header1 = ["과      목", current_header, "", prior_header, ""]
    header2 = ["", "금 액", "금 액", "금 액", "금 액"]
    data = [header1, header2]

    row_styles = []  # (row_index, level, bold)
    for i, row in enumerate(rows):
        r_idx = len(data)
        indent = "    " * row.level
        label_cell = Paragraph(indent + row.label, ParagraphStyle(
            f"label{i}", fontName=FONT_BOLD if row.bold else FONT_REGULAR, fontSize=9,
            leading=13,
        ))
        data.append([
            label_cell,
            _fmt(row.detail_current), _fmt(row.total_current),
            _fmt(row.detail_prior), _fmt(row.total_prior),
        ])
        row_styles.append((r_idx, row.bold))

    table = Table(data, colWidths=[label_w, amt_w, amt_w, amt_w, amt_w], repeatRows=2)

    style_cmds = [
        ("SPAN", (1, 0), (2, 0)), ("SPAN", (3, 0), (4, 0)),
        ("SPAN", (0, 0), (0, 1)),
        ("FONTNAME", (0, 0), (-1, 1), FONT_BOLD),
        ("FONTSIZE", (0, 0), (-1, -1), 9),
        ("ALIGN", (0, 0), (-1, 1), "CENTER"),
        ("ALIGN", (1, 2), (-1, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.black),
        ("BACKGROUND", (0, 0), (-1, 1), colors.whitesmoke),
        ("TOPPADDING", (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (1, 0), (-1, -1), 6),
    ]
    for r_idx, bold in row_styles:
        if bold:
            style_cmds.append(("FONTNAME", (1, r_idx), (-1, r_idx), FONT_BOLD))

    table.setStyle(TableStyle(style_cmds))
    story.append(table)

    doc.build(story)
    return output_path
