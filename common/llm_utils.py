"""
llm_utils.py
Claude API(Anthropic) 연동 공통 유틸. `.env`의 ANTHROPIC_API_KEY로 설정한다.

이 프로젝트의 다른 외부 연동(sheet_io.py의 구글시트, notify_utils.py의
이메일/슬랙/팀즈)과 같은 원칙을 따른다 - 키가 없으면 에러를 던지지 않고
"미설정" 상태를 알려주며, 호출하는 스크립트가 규칙 기반 폴백 로직으로
대체하게 한다. 그래야 API 키 없이 클론해도 스크립트가 항상 끝까지 돈다.

주의: 실제로 API를 호출하면 사용자 본인의 Anthropic API 키 사용량에 따라
비용이 발생한다(무료 아님). 이 저장소 자체는 비용을 대신 내주지 않는다.
"""
from __future__ import annotations

import json
import os
from typing import Optional

from dotenv import load_dotenv

load_dotenv()

DEFAULT_MODEL = os.getenv("ANTHROPIC_MODEL", "claude-sonnet-5")


def is_configured() -> bool:
    return bool(os.getenv("ANTHROPIC_API_KEY"))


def call_claude(prompt: str, system: Optional[str] = None, max_tokens: int = 2000) -> str:
    """Claude에 프롬프트를 보내고 텍스트 응답을 받는다.

    키가 없는 상태에서 호출하면 예외를 던진다 - 반드시 호출 전에
    is_configured()로 먼저 확인하고, 미설정이면 폴백 로직을 타도록 스크립트를
    작성할 것.
    """
    try:
        import anthropic
    except ImportError as e:
        raise ImportError(
            "anthropic 패키지가 설치돼 있지 않습니다. pip install -r requirements.txt 로 설치해주세요."
        ) from e

    client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
    kwargs = {"system": system} if system else {}
    response = client.messages.create(
        model=DEFAULT_MODEL,
        max_tokens=max_tokens,
        messages=[{"role": "user", "content": prompt}],
        **kwargs,
    )
    return "".join(block.text for block in response.content if block.type == "text")


def extract_json(text: str) -> dict:
    """Claude 응답 텍스트에서 JSON을 안전하게 추출한다.

    ```json ... ``` 코드펜스로 감싸서 응답하는 경우가 많아 그 형태를 먼저
    처리하고, 그래도 실패하면 첫 '{'부터 마지막 '}'까지를 잘라 재시도한다.
    """
    text = text.strip()
    if text.startswith("```"):
        parts = text.split("```")
        text = parts[1] if len(parts) > 1 else text
        if text.startswith("json"):
            text = text[4:]
    text = text.strip()
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        start, end = text.find("{"), text.rfind("}")
        if start != -1 and end != -1:
            return json.loads(text[start : end + 1])
        raise
