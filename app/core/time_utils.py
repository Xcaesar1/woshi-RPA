from __future__ import annotations

from datetime import datetime, timedelta, timezone


BEIJING_TZ = timezone(timedelta(hours=8), name="Asia/Shanghai")
DISPLAY_DATETIME_FORMAT = "%Y-%m-%d %H:%M:%S"
TASK_ID_DATETIME_FORMAT = "%Y%m%d-%H%M%S"


def beijing_now() -> datetime:
    return datetime.now(BEIJING_TZ)


def beijing_now_iso() -> str:
    return beijing_now().isoformat(timespec="seconds")


def beijing_now_display() -> str:
    return beijing_now().strftime(DISPLAY_DATETIME_FORMAT)


def beijing_task_id_timestamp() -> str:
    return beijing_now().strftime(TASK_ID_DATETIME_FORMAT)


def beijing_threshold_iso(*, seconds: int = 0, days: int = 0) -> str:
    return (beijing_now() - timedelta(seconds=seconds, days=days)).isoformat(timespec="seconds")


def format_datetime_for_display(value: object) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return text
    if "T" not in text:
        return text

    try:
        parsed = datetime.fromisoformat(text)
    except ValueError:
        return text.replace("T", " ")

    if parsed.tzinfo is None:
        # Historical Docker records were stored as UTC-like naive ISO strings.
        parsed = parsed.replace(tzinfo=timezone.utc)
    return parsed.astimezone(BEIJING_TZ).strftime(DISPLAY_DATETIME_FORMAT)
