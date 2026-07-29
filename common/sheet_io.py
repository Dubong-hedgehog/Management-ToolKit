"""
sheet_io.py
Google Sheets를 데이터 소스로 쓰기 위한 공통 연동 모듈.

회사 계정이나 시트가 바뀌어도 코드를 고칠 필요가 없도록, 시트 ID / 서비스 계정
키 경로 / 탭 이름을 전부 .env 파일에서 읽어온다. .env는 .gitignore에 포함되어
있어 실제 값은 절대 커밋되지 않는다 (.env.example 참고).

사용 전 준비:
  1) Google Cloud Console에서 서비스 계정을 만들고 JSON 키를 다운로드
  2) 대상 스프레드시트를 그 서비스 계정 이메일(...@...iam.gserviceaccount.com)과
     '공유' (뷰어 권한이면 충분)
  3) .env.example을 .env로 복사하고 GOOGLE_SHEET_ID / GOOGLE_SERVICE_ACCOUNT_FILE 등을 채우기

나중에 회사/계정이 바뀌면 .env 안의 값만 새로 채우면 되고, 이 파일이나
스크립트 코드는 건드릴 필요가 없다.
"""
from __future__ import annotations

import os
from pathlib import Path
from typing import Optional

import pandas as pd
from dotenv import load_dotenv

load_dotenv()  # 프로젝트 루트의 .env를 읽어 os.environ에 채워줌 (없으면 조용히 무시)

_SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]


def is_configured() -> bool:
    """.env에 구글 시트 연동에 필요한 값이 다 채워져 있는지 확인.

    스크립트들은 이 값이 True일 때만 실제 구글 시트에 붙고, False면 로컬
    sample_data로 자동 폴백한다 (누구나 클론 후 바로 실행 가능해야 하므로).
    """
    sheet_id = os.getenv("GOOGLE_SHEET_ID")
    cred_path = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE")
    return bool(sheet_id and cred_path and Path(cred_path).exists())


def load_transactions_from_sheet(
    sheet_id: Optional[str] = None,
    worksheet_name: Optional[str] = None,
    credentials_path: Optional[str] = None,
) -> pd.DataFrame:
    """구글 스프레드시트 한 탭을 통째로 읽어 DataFrame으로 반환.

    인자를 안 넘기면 .env 값(GOOGLE_SHEET_ID, GOOGLE_WORKSHEET_NAME,
    GOOGLE_SERVICE_ACCOUNT_FILE)을 사용한다. 시트 첫 행은 헤더로 취급된다.
    """
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
        raise FileNotFoundError(
            f"서비스 계정 키 파일을 찾을 수 없습니다: {credentials_path} "
            "(.env의 GOOGLE_SERVICE_ACCOUNT_FILE 경로 확인)"
        )

    creds = Credentials.from_service_account_file(credentials_path, scopes=_SCOPES)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(sheet_id)
    worksheet = sheet.worksheet(worksheet_name)

    records = worksheet.get_all_records()
    return pd.DataFrame(records)
