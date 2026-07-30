"""
sheet_io.py
Google Sheets를 데이터 소스로 쓰기 위한 공통 연동 모듈. (.env 기반 설정)
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

import pandas as pd
from dotenv import load_dotenv

load_dotenv()

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def is_configured() -> bool:
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    cred_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
    return bool(sheet_id and cred_path and Path(cred_path).exists())


def load_transactions_from_sheet(
    sheet_id: Optional[str] = None,
    worksheet_name: Optional[str] = None,
    credentials_path: Optional[str] = None,
) -> pd.DataFrame:
    try:
        import gspread
        from google.oauth2.service_account import Credentials
    except ImportError as e:
        raise ImportError(
            "gspread/google-auth가 설치돼 있지 않습니다. pip install -r requirements.txt 로 설치해주세요."
        ) from e

    sheet_id = sheet_id or os.getenv("GOOGLE_SHEET_ID")
    worksheet_name = worksheet_name or os.getenv("GOOGLE_WORKSHEET_NAME", "Sheet1")
    credentials_path = credentials_path or os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")

    if not sheet_id:
        raise ValueError("GOOGLE_SHEET_ID가 설정되지 않았습니다 (.env 확인).")
    if not credentials_path or not Path(credentials_path).exists():
        raise FileNotFoundError(f"서비스 계정 키 파일을 찾을 수 없습니다: {credentials_path}")

    creds = Credentials.from_service_account_file(credentials_path, scopes=_SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(worksheet_name)
    records = worksheet.get_all_records()
    return pd.DataFrame(records)
