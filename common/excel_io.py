"""
excel_io.py
경영지원팀 업무에서 가장 많이 쓰는 '엑셀 읽고 쓰기'를 공통화한 모듈.

모든 카테고리(finance, hr, procurement ...)의 스크립트는 이 모듈을 통해
엑셀을 읽고 쓴다. 이렇게 하면:
  1) 파일마다 pandas 옵션을 다르게 쓰는 실수를 줄이고
  2) 나중에 엑셀 대신 구글시트/DB로 데이터 소스를 바꿔도 이 파일만 고치면 된다.
"""
from __future__ import annotations

from pathlib import Path
from typing import Mapping

import pandas as pd


def load_excel(path: str | Path, sheet_name: str | int = 0) -> pd.DataFrame:
    """엑셀 파일을 DataFrame으로 읽어온다.

    Args:
        path: 읽을 엑셀 파일 경로 (.xlsx)
        sheet_name: 시트명 또는 인덱스 (기본값: 첫 번째 시트)

    Returns:
        pandas.DataFrame
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"엑셀 파일을 찾을 수 없습니다: {path}")
    return pd.read_excel(path, sheet_name=sheet_name, engine="openpyxl")


def load_csv(path: str | Path, **kwargs) -> pd.DataFrame:
    """CSV 파일을 DataFrame으로 읽어온다. (인코딩은 기본 UTF-8, 필요시 kwargs로 override)"""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(f"CSV 파일을 찾을 수 없습니다: {path}")
    kwargs.setdefault("encoding", "utf-8-sig")
    return pd.read_csv(path, **kwargs)


def save_excel_report(
    sheets: Mapping[str, pd.DataFrame],
    output_path: str | Path,
    autofit: bool = True,
) -> Path:
    """여러 개의 DataFrame을 시트별로 나눠서 하나의 엑셀 리포트로 저장한다.

    Args:
        sheets: {시트이름: DataFrame} 형태의 딕셔너리
        output_path: 저장할 경로 (.xlsx)
        autofit: 컬럼 너비를 내용에 맞게 자동 조정할지 여부

    Returns:
        저장된 파일의 Path
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        for sheet_name, df in sheets.items():
            # 엑셀 시트명은 31자 제한이 있어 잘라준다.
            safe_name = sheet_name[:31]
            df.to_excel(writer, sheet_name=safe_name, index=False)

            if autofit:
                worksheet = writer.sheets[safe_name]
                for col_idx, col in enumerate(df.columns):
                    max_len = max(
                        [len(str(col))] + [len(str(v)) for v in df[col].astype(str).tolist()]
                    )
                    worksheet.set_column(col_idx, col_idx, min(max_len + 2, 40))

    return output_path
