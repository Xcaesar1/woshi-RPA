from __future__ import annotations

import argparse
import os
import threading
import time
from pathlib import Path

from app.core.config import (
    HEARTBEAT_INTERVAL_SECONDS,
    STALE_TASK_TIMEOUT_SECONDS,
    TASK_CLEANUP_INTERVAL_SECONDS,
    TASK_RETENTION_DAYS,
    WORKER_POLL_INTERVAL_SECONDS,
    ensure_app_directories,
)
from app.core.db import init_db
from app.models.task import TASK_STATUS_QUEUED
from app.services.file_service import append_log_line
from app.services.queue_service import (
    BrowserSlotLease,
    acquire_queue_pop_lock,
    acquire_browser_slot,
    clear_runtime_locks,
    list_queue_member_ids,
    pop_task_id,
    record_worker_heartbeat,
    requeue_missing_task_ids,
)
from app.services.task_service import (
    claim_task,
    list_expired_terminal_tasks,
    list_task_ids_by_status,
    reset_stale_running_tasks,
    touch_task_heartbeat,
)
from app.services.workflow_service import cleanup_expired_tasks, process_task


class StoppableThread(threading.Thread):
    def __init__(self, interval_seconds: int):
        super().__init__(daemon=True)
        self.interval_seconds = interval_seconds
        self._stop_event = threading.Event()

    def stop(self) -> None:
        self._stop_event.set()


class TaskHeartbeatThread(StoppableThread):
    def __init__(self, task_id: str, interval_seconds: int):
        super().__init__(interval_seconds)
        self.task_id = task_id

    def run(self) -> None:
        while not self._stop_event.wait(self.interval_seconds):
            touch_task_heartbeat(self.task_id)


class WorkerHeartbeatThread(StoppableThread):
    def __init__(self, worker_id: str, interval_seconds: int):
        super().__init__(interval_seconds)
        self.worker_id = worker_id

    def run(self) -> None:
        while not self._stop_event.wait(self.interval_seconds):
            record_worker_heartbeat(self.worker_id)


class BrowserSlotRenewThread(StoppableThread):
    def __init__(self, lease: BrowserSlotLease, interval_seconds: int):
        super().__init__(interval_seconds)
        self.lease = lease

    def run(self) -> None:
        while not self._stop_event.wait(self.interval_seconds):
            self.lease.extend()


def build_worker_id() -> str:
    return f"{os.environ.get('HOSTNAME') or os.environ.get('COMPUTERNAME') or 'worker'}-{os.getpid()}"


def reconcile_queue_state() -> dict[str, int]:
    stale_ids = reset_stale_running_tasks(STALE_TASK_TIMEOUT_SECONDS)
    queued_ids = list_task_ids_by_status(TASK_STATUS_QUEUED)
    queued_members = list_queue_member_ids()
    missing_ids = [task_id for task_id in queued_ids if task_id not in queued_members]
    requeued_ids = requeue_missing_task_ids(missing_ids)
    return {
        "stale_count": len(stale_ids),
        "requeued_count": len(requeued_ids),
    }


def cleanup_history() -> int:
    return cleanup_expired_tasks(list_expired_terminal_tasks(TASK_RETENTION_DAYS))


def reconcile_runtime_state() -> dict[str, int]:
    summary = reconcile_queue_state()
    summary["cleaned_count"] = cleanup_history()
    return summary


def process_claimed_task(task: dict, lease: BrowserSlotLease) -> None:
    log_path = task.get("log_path")
    if log_path:
        append_log_line(Path(log_path), f"worker 已领取任务，使用浏览器执行槽 #{lease.slot_index}")

    heartbeat = TaskHeartbeatThread(task["id"], HEARTBEAT_INTERVAL_SECONDS)
    renew = BrowserSlotRenewThread(lease, HEARTBEAT_INTERVAL_SECONDS)
    touch_task_heartbeat(task["id"])
    heartbeat.start()
    renew.start()
    try:
        process_task(task)
    finally:
        heartbeat.stop()
        renew.stop()
        heartbeat.join(timeout=1)
        renew.join(timeout=1)


def run_worker(*, once: bool = False, poll_interval: int = WORKER_POLL_INTERVAL_SECONDS) -> None:
    ensure_app_directories()
    init_db()
    clear_runtime_locks()

    worker_id = build_worker_id()
    record_worker_heartbeat(worker_id)
    worker_heartbeat = WorkerHeartbeatThread(worker_id, HEARTBEAT_INTERVAL_SECONDS)
    worker_heartbeat.start()

    last_maintenance_at = 0.0
    try:
        while True:
            now = time.time()
            reconcile_queue_state()
            if last_maintenance_at == 0.0 or now - last_maintenance_at >= TASK_CLEANUP_INTERVAL_SECONDS:
                cleanup_history()
                last_maintenance_at = now

            pop_lease = acquire_queue_pop_lock(blocking_timeout_seconds=poll_interval)
            if pop_lease is None:
                if once:
                    return
                continue

            lease = None
            try:
                task_id = pop_task_id()
                if task_id is None:
                    if once:
                        return
                    time.sleep(poll_interval)
                    continue

                lease = acquire_browser_slot()
                if lease is None:
                    if once:
                        return
                    continue

                task = claim_task(task_id)
                if task is None:
                    if once:
                        return
                    continue

                process_claimed_task(task, lease)
                if once:
                    return
            finally:
                if lease is not None:
                    lease.release()
                pop_lease.release()
    finally:
        worker_heartbeat.stop()
        worker_heartbeat.join(timeout=1)


def main() -> None:
    parser = argparse.ArgumentParser(description="领星内网页面任务 worker")
    parser.add_argument("--once", action="store_true", help="只尝试领取并执行一次")
    parser.add_argument("--poll-interval", type=int, default=WORKER_POLL_INTERVAL_SECONDS)
    args = parser.parse_args()
    run_worker(once=args.once, poll_interval=args.poll_interval)


if __name__ == "__main__":
    main()
