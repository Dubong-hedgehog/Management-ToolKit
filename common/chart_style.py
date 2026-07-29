"""
chart_style.py
모든 카테고리 스크립트가 같은 느낌의 차트를 그리도록 스타일을 통일하는 모듈.

한글 폰트가 없는 환경(리눅스 서버 등)에서 실행해도 깨지지 않도록
koreanize-matplotlib이 설치돼 있으면 자동으로 나눔고딕 계열 폰트를 등록한다.
못 찾으면 경고만 남기고 기본 폰트로 진행한다 (그래프 자체는 정상 생성됨).
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt

matplotlib.use("Agg")  # 화면 없는 서버 환경에서도 그래프 파일 저장 가능하게

_KOREAN_FONT_CANDIDATES = [
    "NanumGothic",
    "Malgun Gothic",
    "AppleGothic",
    "NanumBarunGothic",
]


def setup_style() -> None:
    """차트 공통 스타일(색상, 한글 폰트, 그리드 등) 적용."""
    # 시스템에 한글 폰트가 없는 환경(리눅스 서버/CI 등)을 위해
    # koreanize-matplotlib이 설치돼 있으면 자동으로 나눔고딕을 등록해준다.
    try:
        import koreanize_matplotlib  # noqa: F401
    except ImportError:
        pass

    available = {f.name for f in fm.fontManager.ttflist}
    for candidate in _KOREAN_FONT_CANDIDATES:
        if candidate in available:
            plt.rcParams["font.family"] = candidate
            break
    else:
        print("[chart_style] 경고: 한글 폰트를 찾지 못했습니다. 라벨이 깨질 수 있습니다. "
              "(리포트 실행에는 영향 없음, 필요시 나눔고딕 설치)")

    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 120
    plt.rcParams["axes.grid"] = True
    plt.rcParams["grid.alpha"] = 0.3
    plt.rcParams["axes.spines.top"] = False
    plt.rcParams["axes.spines.right"] = False


# 카테고리 전반에서 재사용할 팔레트 (경영지원팀 문서 톤에 맞춘 차분한 색)
PALETTE = ["#2E5EAA", "#E4572E", "#76B041", "#F4A300", "#8E44AD", "#17A398"]


def save_chart(fig, output_path: str | Path, tight: bool = True) -> Path:
    """차트를 파일로 저장하고 저장 경로를 반환한다."""
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if tight:
        fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight")
    plt.close(fig)
    return output_path
