from __future__ import annotations

import secrets

from app.core.time_utils import beijing_now_iso, beijing_task_id_timestamp, beijing_threshold_iso
from app.core.db import get_connection, row_to_dict
from app.models.task import (
    TASK_STATUS_FAILED,
    TASK_STATUS_PENDING,
    TASK_STATUS_QUEUED,
    TASK_STATUS_RUNNING,
    TERMINAL_TASK_STATUSES,
)


def iso_now() -> str:
    return beijing_now_iso()


def generate_task_id() -> str:
    return f"{beijing_task_id_timestamp()}-{secrets.token_hex(3)}"


def create_task(
    *,
    task_id: str,
    workflow_name: str,
    original_filename: str,
    submitter: str,
    remark: str | None,
    upload_path: str,
    job_dir: str,
    log_path: str,
    total_fba_count: int,
) -> dict:
    created_at = iso_now()
    queued_at = iso_now()
    with get_connection() as connection:
        connection.execute(
            """
            INSERT INTO tasks (
                id, workflow_name, original_filename, submitter, remark, status,
                created_at, queued_at, upload_path, job_dir, log_path,
                total_fba_count, success_fba_count, failed_fba_count
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, 0)
            """,
            (
                task_id,
                workflow_name,
                original_filename,
                submitter,
                remark,
                TASK_STATUS_PENDING,
                created_at,
                None,
                upload_path,
                job_dir,
                log_path,
                total_fba_count,
            ),
        )
        connection.execute(
            "UPDATE tasks SET status = ?, queued_at = ? WHERE id = ?",
            (TASK_STATUS_QUEUED, queued_at, task_id),
        )
    return get_task(task_id)


def get_task(task_id: str) -> dict | None:
    with get_connection() as connection:
        row = connection.execute("SELECT * FROM tasks WHERE id = ?", (task_id,)).fetchone()
    return row_to_dict(row)


def list_tasks(*, submitter: str | None = None, status: str | None = None) -> list[dict]:
    query = "SELECT * FROM tasks WHERE 1=1"
    params: list[str] = []
    if submitter:
        query += " AND submitter = ?"
        params.append(submitter)
    if status:
        query += " AND status = ?"
        params.append(status)
    query += " ORDER BY created_at DESC"
    with get_connection() as connection:
        rows = connection.execute(query, params).fetchall()
    return [dict(row) for row in rows]


def list_task_ids_by_status(status: str) -> list[str]:
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id FROM tasks
            WHERE status = ?
            ORDER BY COALESCE(queued_at, created_at) ASC, created_at ASC
            """,
            (status,),
        ).fetchall()
    return [row["id"] for row in rows]


def count_tasks(*, status: str | None = None) -> int:
    with get_connection() as connection:
        if status:
            row = connection.execute("SELECT COUNT(*) AS count FROM tasks WHERE status = ?", (status,)).fetchone()
        else:
            row = connection.execute("SELECT COUNT(*) AS count FROM tasks").fetchone()
    return int(row["count"]) if row else 0


def claim_task(task_id: str) -> dict | None:
    now = iso_now()
    with get_connection() as connection:
        updated = connection.execute(
            """
            UPDATE tasks
            SET status = ?, started_at = COALESCE(started_at, ?), heartbeat_at = ?
            WHERE id = ? AND status = ?
            """,
            (TASK_STATUS_RUNNING, now, now, task_id, TASK_STATUS_QUEUED),
        ).rowcount
    if updated == 0:
        return None
    return get_task(task_id)


def claim_next_queued_task() -> dict | None:
    connection = get_connection()
    try:
        connection.isolation_level = None
        connection.execute("BEGIN IMMEDIATE")
        row = connection.execute(
            """
            SELECT * FROM tasks
            WHERE status = ?
            ORDER BY COALESCE(queued_at, created_at) ASC, created_at ASC
            LIMIT 1
            """,
            (TASK_STATUS_QUEUED,),
        ).fetchone()
        if row is None:
            connection.execute("COMMIT")
            return None

        now = iso_now()
        updated = connection.execute(
            """
            UPDATE tasks
            SET status = ?, started_at = COALESCE(started_at, ?), heartbeat_at = ?
            WHERE id = ? AND status = ?
            """,
            (TASK_STATUS_RUNNING, now, now, row["id"], TASK_STATUS_QUEUED),
        ).rowcount
        connection.execute("COMMIT")
        if updated == 0:
            return None
    except Exception:
        connection.execute("ROLLBACK")
        raise
    finally:
        connection.close()
    return get_task(row["id"])


def touch_task_heartbeat(task_id: str) -> None:
    with get_connection() as connection:
        connection.execute(
            "UPDATE tasks SET heartbeat_at = ? WHERE id = ?",
            (iso_now(), task_id),
        )


def mark_task_finished(
    *,
    task_id: str,
    status: str,
    result_zip_path: str | None,
    result_primary_file: str | None,
    error_message: str | None,
    total_fba_count: int,
    success_fba_count: int,
    failed_fba_count: int,
) -> None:
    now = iso_now()
    with get_connection() as connection:
        connection.execute(
            """
            UPDATE tasks
            SET status = ?,
                finished_at = ?,
                heartbeat_at = ?,
                result_zip_path = ?,
                result_primary_file = ?,
                error_message = ?,
                total_fba_count = ?,
                success_fba_count = ?,
                failed_fba_count = ?
            WHERE id = ?
            """,
            (
                status,
                now,
                now,
                result_zip_path,
                result_primary_file,
                error_message,
                total_fba_count,
                success_fba_count,
                failed_fba_count,
                task_id,
            ),
        )


def mark_task_failed(task_id: str, error_message: str) -> None:
    task = get_task(task_id)
    total = int(task.get("total_fba_count") or 0) if task else 0
    success = int(task.get("success_fba_count") or 0) if task else 0
    failed = 0
    if task:
        failed = max(total - success, 0)
        if failed == 0:
            failed = int(task.get("failed_fba_count") or 0)
    mark_task_finished(
        task_id=task_id,
        status=TASK_STATUS_FAILED,
        result_zip_path=task.get("result_zip_path") if task else None,
        result_primary_file=task.get("result_primary_file") if task else None,
        error_message=error_message,
        total_fba_count=total,
        success_fba_count=success,
        failed_fba_count=failed,
    )


def reset_stale_running_tasks(timeout_seconds: int) -> list[str]:
    threshold = beijing_threshold_iso(seconds=timeout_seconds)
    stale_ids: list[str] = []
    with get_connection() as connection:
        rows = connection.execute(
            """
            SELECT id FROM tasks
            WHERE status = ?
              AND COALESCE(heartbeat_at, started_at, queued_at, created_at) < ?
            """,
            (TASK_STATUS_RUNNING, threshold),
        ).fetchall()
        stale_ids = [row["id"] for row in rows]
        for task_id in stale_ids:
            connection.execute(
                """
                UPDATE tasks
                SET status = ?, finished_at = ?, error_message = ?
                WHERE id = ?
                """,
                (
                    TASK_STATUS_FAILED,
                    iso_now(),
                    "worker 心跳超时，任务已自动标记为失败",
                    task_id,
                ),
            )
    return stale_ids


def list_expired_terminal_tasks(retention_days: int) -> list[dict]:
    threshold = beijing_threshold_iso(days=retention_days)
    placeholders = ",".join(["?"] * len(TERMINAL_TASK_STATUSES))
    params = [*sorted(TERMINAL_TASK_STATUSES), threshold]
    with get_connection() as connection:
        rows = connection.execute(
            f"""
            SELECT * FROM tasks
            WHERE status IN ({placeholders})
              AND COALESCE(finished_at, created_at) < ?
            ORDER BY COALESCE(finished_at, created_at) ASC
            """,
            params,
        ).fetchall()
    return [dict(row) for row in rows]


def delete_tasks(task_ids: list[str]) -> None:
    if not task_ids:
        return
    placeholders = ",".join(["?"] * len(task_ids))
    with get_connection() as connection:
        connection.execute(f"DELETE FROM tasks WHERE id IN ({placeholders})", task_ids)
