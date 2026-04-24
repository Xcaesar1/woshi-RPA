from __future__ import annotations

import sqlite3
from pathlib import Path

from app.core.config import DB_PATH, SQLITE_JOURNAL_MODE, ensure_app_directories


def get_connection() -> sqlite3.Connection:
    ensure_app_directories()
    connection = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    connection.row_factory = sqlite3.Row
    connection.execute(f"PRAGMA journal_mode={SQLITE_JOURNAL_MODE}")
    connection.execute("PRAGMA foreign_keys=ON")
    return connection


def row_to_dict(row: sqlite3.Row | None) -> dict | None:
    if row is None:
        return None
    return dict(row)


def init_db() -> None:
    ensure_app_directories()
    with get_connection() as connection:
        connection.executescript(
            """
            CREATE TABLE IF NOT EXISTS tasks (
                id TEXT PRIMARY KEY,
                workflow_name TEXT NOT NULL,
                original_filename TEXT NOT NULL,
                submitter TEXT NOT NULL,
                remark TEXT,
                status TEXT NOT NULL,
                created_at TEXT NOT NULL,
                queued_at TEXT,
                started_at TEXT,
                heartbeat_at TEXT,
                finished_at TEXT,
                upload_path TEXT NOT NULL,
                job_dir TEXT NOT NULL,
                result_zip_path TEXT,
                result_primary_file TEXT,
                log_path TEXT NOT NULL,
                error_message TEXT,
                total_fba_count INTEGER NOT NULL DEFAULT 0,
                success_fba_count INTEGER NOT NULL DEFAULT 0,
                failed_fba_count INTEGER NOT NULL DEFAULT 0
            );
            CREATE INDEX IF NOT EXISTS idx_tasks_status_created_at
                ON tasks(status, created_at DESC);
            CREATE INDEX IF NOT EXISTS idx_tasks_submitter_created_at
                ON tasks(submitter, created_at DESC);
            """
        )
