from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


@dataclass(frozen=True)
class WorkflowDefinition:
    name: str
    label: str


ROOT_DIR = Path(__file__).resolve().parents[2]
APP_DIR = ROOT_DIR / "app"
DATA_DIR = ROOT_DIR / "data"
UPLOADS_DIR = DATA_DIR / "uploads"
JOBS_DIR = DATA_DIR / "jobs"
RESULTS_DIR = DATA_DIR / "results"
LOGS_DIR = DATA_DIR / "logs"
BROWSER_DIR = DATA_DIR / "browser"
BROWSER_PROFILE_DIR = BROWSER_DIR / "profile_playwright"
DB_DIR = DATA_DIR / "db"
DB_PATH = DB_DIR / "tasks.sqlite3"
SQLITE_JOURNAL_MODE = os.environ.get("SQLITE_JOURNAL_MODE", "WAL").strip().upper() or "WAL"

RESOURCE_DIR = ROOT_DIR
DEFAULT_CONFIG_PATH = ROOT_DIR / "lingxing_rpa.local.json"
EXAMPLE_MANIFESTS = {
    "fba_manifest.txt": ROOT_DIR / "fba_manifest.txt",
    "fba_manifest.xlsx": ROOT_DIR / "fba_manifest.xlsx",
}


def env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    if raw is None or not raw.strip():
        return default
    try:
        return int(raw.strip())
    except ValueError:
        return default


def env_bool(name: str, default: bool) -> bool:
    raw = os.environ.get(name)
    if raw is None:
        return default
    return raw.strip().lower() in {"1", "true", "yes", "on"}


REDIS_URL = os.environ.get("REDIS_URL", "redis://127.0.0.1:6379/0")
REDIS_TASK_QUEUE_KEY = os.environ.get("REDIS_TASK_QUEUE_KEY", "lingxing:tasks:queue")
REDIS_TASK_QUEUE_MEMBERS_KEY = os.environ.get("REDIS_TASK_QUEUE_MEMBERS_KEY", "lingxing:tasks:queue_members")
REDIS_BROWSER_SLOT_PREFIX = os.environ.get("REDIS_BROWSER_SLOT_PREFIX", "lingxing:browser:slot")
REDIS_QUEUE_POP_LOCK_KEY = os.environ.get("REDIS_QUEUE_POP_LOCK_KEY", "lingxing:tasks:queue_pop_lock")
REDIS_WORKER_HEARTBEAT_KEY = os.environ.get("REDIS_WORKER_HEARTBEAT_KEY", "lingxing:worker:heartbeat")

HEARTBEAT_INTERVAL_SECONDS = env_int("HEARTBEAT_INTERVAL_SECONDS", 15)
STALE_TASK_TIMEOUT_SECONDS = env_int("STALE_TASK_TIMEOUT_SECONDS", 10 * 60)
WORKER_POLL_INTERVAL_SECONDS = env_int("WORKER_POLL_INTERVAL_SECONDS", 5)
TASK_RETENTION_DAYS = env_int("TASK_RETENTION_DAYS", 30)
TASK_CLEANUP_INTERVAL_SECONDS = env_int("TASK_CLEANUP_INTERVAL_SECONDS", 6 * 60 * 60)
BROWSER_MAX_CONCURRENCY = max(1, env_int("BROWSER_MAX_CONCURRENCY", 1))
BROWSER_SLOT_LOCK_TTL_SECONDS = env_int("BROWSER_SLOT_LOCK_TTL_SECONDS", max(HEARTBEAT_INTERVAL_SECONDS * 4, 120))
QUEUE_POP_LOCK_TTL_SECONDS = env_int("QUEUE_POP_LOCK_TTL_SECONDS", max(WORKER_POLL_INTERVAL_SECONDS * 3, 30))
WORKER_HEARTBEAT_TTL_SECONDS = env_int("WORKER_HEARTBEAT_TTL_SECONDS", max(HEARTBEAT_INTERVAL_SECONDS * 4, 60))
PLAYWRIGHT_HEADLESS = env_bool("PLAYWRIGHT_HEADLESS", False)
RECENT_LOG_LINE_COUNT = 120

DEFAULT_WORKFLOW = WorkflowDefinition(
    name="lingxing_fba_download_and_process",
    label="领星 FBA 下载并整理",
)
WORKFLOW_REGISTRY = {
    DEFAULT_WORKFLOW.name: DEFAULT_WORKFLOW,
}


def ensure_app_directories() -> None:
    for path in [
        DATA_DIR,
        UPLOADS_DIR,
        JOBS_DIR,
        RESULTS_DIR,
        LOGS_DIR,
        BROWSER_DIR,
        BROWSER_PROFILE_DIR,
        DB_DIR,
    ]:
        path.mkdir(parents=True, exist_ok=True)
