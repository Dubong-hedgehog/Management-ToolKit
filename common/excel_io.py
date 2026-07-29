"""
excel_io.py
경영지원팀 업무에서 가장 많이 쓰는 '엑셀 읽고 쓰기'를 공통화한 모듈.
"""
from __future__ import annotations

from pathlib import Path
from typing import Mapping

import pandas as pd


def load_excel(path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {path}")
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def load_csv(path: str | Path, **kwargs) -> pd.DataFrame:
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"CSV 파일을 찾을 수 없습니다: {path}")
    kwargs.setdefault("encoding", "utf-8-sig")
    return pd.read_csv(path, **kwargs)


def save_excel_report(
    sheets: Mapping[str, pd.DataFrame],
    output_path: str | Path,
    autofit: bool = True,
    autofilter: bool = True,
) -> Path:
    """여러 DataFrame을 시트별로 저장한다.

    autofilter=True(기본값)면 각 시트 헤더 행에 엑셀 필터 드롭다운을 켜준다.
    데이터가 있는 시트를 열었을 때 바로 "위험군만", "이상패턴만" 같은 식으로
    특정 컬럼 값 기준으로 걸러볼 수 있게 하기 위함. 빈 시트는 건너뛴다.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)
            worksheet = writer.sheets[safe_name]

            if autofit:
                for col_idx, col in enumerate(df.columns):
                    max_len = max(
                        [len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()]
                    )
                    worksheet.set_column(col_idx, col_idx, min(max_len + 2, 40))

            if autofilter and len(df) > 0 and len(df.columns) > 0:
                worksheet.autofilter(0, 0, len(df), len(df.columns) - 1)

    return output_path
