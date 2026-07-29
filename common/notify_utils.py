"""
notify_utils.py
계약만료 임박 같은 "누군가 챙겨야 하는" 알림을 보내는 공통 모듈. (.env 기반 설정)

기본 채널은 이메일이고, NOTIFY_CHANNEL을 slack 또는 teams로 바꾸면 웹훅으로
보낸다. 자격증명이 하나도 설정 안 돼 있어도 에러 없이 콘솔에 출력 + 알림
내역을 파일로 남기는 것으로 대체하기 때문에, 클론만 받아도(별도 설정 없이)
바로 실행이 된다 — 이 저장소의 다른 모듈들(sheet_io.py 등)과 동일한 설계
원칙이다.

.env 설정 예시:
    NOTIFY_CHANNEL=email          # email(기본값) | slack | teams

    # channel=email 일 때
    SMTP_HOST=smtp.gmail.com
    SMTP_PORT=587
    SMTP_USER=you@example.com
    SMTP_PASSWORD=xxxx
    NOTIFY_EMAIL_FROM=you@example.com
    NOTIFY_EMAIL_TO=team@example.com

    # channel=slack 일 때
    SLACK_WEBHOOK_URL=https://hooks.slack.com/services/xxx

    # channel=teams 일 때
    TEAMS_WEBHOOK_URL=https://xxx.webhook.office.com/xxx
"""
from __future__ import annotations

import os
import smtplib
from email.mime.text import MIMEText
from pathlib import Path

import pandas as pd
from dotenv import load_dotenv

load_dotenv()


def get_channel() -> str:
    return os.getenv("NOTIFY_CHANNEL", "email").strip().lower()


def is_configured(channel: str | None = None) -> bool:
    channel = channel or get_channel()
    if channel == "email":
        return bool(
            os.getenv("SMTP_HOST") and os.getenv("SMTP_USER") and os.getenv("SMTP_PASSWORD")
            and os.getenv("NOTIFY_EMAIL_TO")
        )
    if channel == "slack":
        return bool(os.getenv("SLACK_WEBHOOK_URL"))
    if channel == "teams":
        return bool(os.getenv("TEAMS_WEBHOOK_URL"))
    return False


def _send_email(subject: str, body: str) -> None:
    host = os.getenv("SMTP_HOST")
    port = int(os.getenv("SMTP_PORT", "587"))
    user = os.getenv("SMTP_USER")
    password = os.getenv("SMTP_PASSWORD")
    sender = os.getenv("NOTIFY_EMAIL_FROM", user)
    to = os.getenv("NOTIFY_EMAIL_TO")

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = sender
    msg["To"] = to

    with smtplib.SMTP(host, port) as server:
        server.starttls()
        server.login(user, password)
        server.sendmail(sender, [to], msg.as_string())


def _send_webhook(url: str, subject: str, body: str) -> None:
    import json
    import urllib.request

    payload = json.dumps({"text": f"*{subject}*\n{body}"}).encode("utf-8")
    req = urllib.request.Request(url, data=payload, headers={"Content-Type": "application/json"})
    urllib.request.urlopen(req, timeout=10)


def send_alert(subject: str, body: str, fallback_log_path: str | Path | None = None) -> str:
    """알림을 보낸다. 반환값은 실제로 무슨 일이 일어났는지에 대한 짧은 설명.

    채널이 설정돼 있지 않으면 예외를 던지지 않고 콘솔 출력 + (지정됐다면)
    fallback_log_path에 CSV로 누적 기록하는 것으로 대체한다. 데모/포트폴리오
    용도로 자격증명 없이도 항상 "정상 동작"하는 걸 보여주기 위함이고, 실제
    운영에서는 .env만 채우면 그대로 이메일/슬랙/팀즈로 나간다.
    """
    channel = get_channel()

    if not is_configured(channel):
        print(f"[알림 미발송 - {channel} 설정 없음] {subject}\n{body}")
        if fallback_log_path:
            fallback_log_path = Path(fallback_log_path)
            fallback_log_path.parent.mkdir(parents=True, exist_ok=True)
            row = pd.DataFrame([{"발송시각": pd.Timestamp.now(), "채널": f"{channel}(미설정)", "제목": subject, "내용": body}])
            if fallback_log_path.exists():
                row.to_csv(fallback_log_path, mode="a", header=False, index=False, encoding="utf-8-sig")
            else:
                row.to_csv(fallback_log_path, index=False, encoding="utf-8-sig")
        return f"미발송(채널 미설정: {channel}) - 콘솔/로그로만 기록"

    try:
        if channel == "email":
            _send_email(subject, body)
        elif channel == "slack":
            _send_webhook(os.getenv("SLACK_WEBHOOK_URL"), subject, body)
        elif channel == "teams":
            _send_webhook(os.getenv("TEAMS_WEBHOOK_URL"), subject, body)
        else:
            raise ValueError(f"알 수 없는 NOTIFY_CHANNEL: {channel}")
        return f"발송완료({channel})"
    except Exception as e:  # noqa: BLE001 - 알림 실패로 본 작업이 죽으면 안 됨
        print(f"[알림 발송 실패 - {channel}] {e}")
        return f"발송실패({channel}): {e}"
