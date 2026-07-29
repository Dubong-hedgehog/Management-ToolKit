"""
chart_style.py
모든 카테고리 스크립트가 같은 느낌의 차트를 그리도록 스타일을 통일하는 모듈.
"""
from __future__ import annotations

from pathlib import Path

import matplotlib
import matplotlib.font_manager as fm
import matplotlib.pyplot as plt

matplotlib.use("Agg")

_KOREAN_FONT_CANDIDATES = [
    "NanumGothic",
    "Malgun Gothic",
    "AppleGothic",
    "NanumBarunGothic",
]


def setup_style() -> None:
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
        print("[chart_style] 경고: 한글 폰트를 찾지 못했습니다. 라벨이 깨질 수 있습니다.")

    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 120
    plt.rcParams["axes.grid"] = True
    plt.rcParams["grid.alpha"] = 0.3
    plt.rcParams["axes.spines.top"] = False
    plt.rcParams["axes.spines.right"] = False


PALETTE = ["#2E5EAA", "#E4572E", "#76B041", "#F4A300", "#8E44AD", "#17A398"]


def save_chart(fig, output_path: str | Path, tight: bool = True, dpi: int = 200) -> Path:
    """차트를 PNG로 저장한다.

    dpi 기본값을 200으로 고정한다 — README에 작게 임베드됐을 때도 글자가
    뭉개지지 않게 하기 위함(예전엔 rcParams의 figure.dpi=120을 그대로
    따라가서 화질이 낮았음). 행 수가 적어 캔버스가 작아지는 표 형태 차트는
    호출하는 쪽에서 figsize를 넉넉히(가로 10~12in 이상) 잡아줘야 이 dpi
    상향의 효과가 제대로 난다.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    if tight:
        fig.tight_layout()
    fig.savefig(output_path, bbox_inches="tight", dpi=dpi)
    plt.close(fig)
    return output_path
