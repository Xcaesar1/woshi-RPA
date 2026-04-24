from __future__ import annotations

import time
from dataclasses import dataclass
from datetime import datetime, timedelta
from functools import lru_cache

from redis import Redis
from redis.exceptions import LockError
from redis.lock import Lock

from app.core.config import (
    BROWSER_MAX_CONCURRENCY,
    BROWSER_SLOT_LOCK_TTL_SECONDS,
    HEARTBEAT_INTERVAL_SECONDS,
    REDIS_BROWSER_SLOT_PREFIX,
    REDIS_QUEUE_POP_LOCK_KEY,
    REDIS_TASK_QUEUE_KEY,
    REDIS_TASK_QUEUE_MEMBERS_KEY,
    REDIS_URL,
    REDIS_WORKER_HEARTBEAT_KEY,
    QUEUE_POP_LOCK_TTL_SECONDS,
    WORKER_POLL_INTERVAL_SECONDS,
    WORKER_HEARTBEAT_TTL_SECONDS,
)


@dataclass
class BrowserSlotLease:
    slot_index: int
    lock: Lock

    def extend(self, ttl_seconds: int = BROWSER_SLOT_LOCK_TTL_SECONDS) -> bool:
        try:
            return bool(self.lock.extend(ttl_seconds, replace_ttl=True))
        except (LockError, AttributeError):
            return False

    def release(self) -> None:
        try:
            if self.lock.owned():
                self.lock.release()
        except LockError:
            return


@dataclass
class QueuePopLease:
    lock: Lock

    def release(self) -> None:
        try:
            if self.lock.owned():
                self.lock.release()
        except LockError:
            return


@lru_cache(maxsize=1)
def get_redis_client() -> Redis:
    return Redis.from_url(
        REDIS_URL,
        decode_responses=True,
        health_check_interval=30,
        socket_connect_timeout=5,
        socket_timeout=max(WORKER_POLL_INTERVAL_SECONDS * 3, 30),
    )


def enqueue_task(task_id: str) -> bool:
    client = get_redis_client()
    added = client.sadd(REDIS_TASK_QUEUE_MEMBERS_KEY, task_id)
    if added:
        client.rpush(REDIS_TASK_QUEUE_KEY, task_id)
        return True
    return False


def dequeue_task_id(timeout_seconds: int) -> str | None:
    client = get_redis_client()
    item = client.blpop(REDIS_TASK_QUEUE_KEY, timeout=max(1, int(timeout_seconds)))
    if not item:
        return None
    _, task_id = item
    client.srem(REDIS_TASK_QUEUE_MEMBERS_KEY, task_id)
    return task_id


def pop_task_id() -> str | None:
    client = get_redis_client()
    task_id = client.lpop(REDIS_TASK_QUEUE_KEY)
    if not task_id:
        return None
    client.srem(REDIS_TASK_QUEUE_MEMBERS_KEY, task_id)
    return task_id


def remove_task_from_queue(task_id: str) -> None:
    client = get_redis_client()
    client.lrem(REDIS_TASK_QUEUE_KEY, 0, task_id)
    client.srem(REDIS_TASK_QUEUE_MEMBERS_KEY, task_id)


def list_queue_member_ids() -> set[str]:
    client = get_redis_client()
    return {item for item in client.smembers(REDIS_TASK_QUEUE_MEMBERS_KEY) if item}


def list_queue_items() -> list[str]:
    client = get_redis_client()
    return [item for item in client.lrange(REDIS_TASK_QUEUE_KEY, 0, -1) if item]


def requeue_missing_task_ids(task_ids: list[str]) -> list[str]:
    requeued: list[str] = []
    for task_id in task_ids:
        if enqueue_task(task_id):
            requeued.append(task_id)
    return requeued


def queue_depth() -> int:
    return int(get_redis_client().llen(REDIS_TASK_QUEUE_KEY))


def record_worker_heartbeat(worker_id: str) -> str:
    now = datetime.now().isoformat(timespec="seconds")
    client = get_redis_client()
    client.hset(REDIS_WORKER_HEARTBEAT_KEY, worker_id, now)
    client.expire(REDIS_WORKER_HEARTBEAT_KEY, WORKER_HEARTBEAT_TTL_SECONDS)
    return now


def get_worker_heartbeat_snapshot(max_age_seconds: int | None = None) -> dict[str, str]:
    client = get_redis_client()
    raw = client.hgetall(REDIS_WORKER_HEARTBEAT_KEY)
    if not raw:
        return {}

    threshold_seconds = max_age_seconds or WORKER_HEARTBEAT_TTL_SECONDS
    cutoff = datetime.now() - timedelta(seconds=threshold_seconds)
    alive: dict[str, str] = {}
    stale_ids: list[str] = []
    for worker_id, timestamp_text in raw.items():
        try:
            timestamp = datetime.fromisoformat(timestamp_text)
        except ValueError:
            stale_ids.append(worker_id)
            continue
        if timestamp >= cutoff:
            alive[worker_id] = timestamp_text
        else:
            stale_ids.append(worker_id)

    if stale_ids:
        client.hdel(REDIS_WORKER_HEARTBEAT_KEY, *stale_ids)
    return alive


def latest_worker_heartbeat() -> str | None:
    snapshot = get_worker_heartbeat_snapshot()
    if not snapshot:
        return None
    return max(snapshot.values())


def is_any_worker_alive() -> bool:
    return bool(get_worker_heartbeat_snapshot())


def acquire_browser_slot(
    *,
    blocking_timeout_seconds: int | None = None,
    retry_interval_seconds: float | None = None,
) -> BrowserSlotLease | None:
    client = get_redis_client()
    retry_interval = retry_interval_seconds or min(HEARTBEAT_INTERVAL_SECONDS / 3, 2.0)
    deadline = time.time() + blocking_timeout_seconds if blocking_timeout_seconds is not None else None

    while True:
        for slot_index in range(1, BROWSER_MAX_CONCURRENCY + 1):
            lock = client.lock(
                f"{REDIS_BROWSER_SLOT_PREFIX}:{slot_index}",
                timeout=BROWSER_SLOT_LOCK_TTL_SECONDS,
                blocking=False,
                thread_local=False,
            )
            if lock.acquire(blocking=False):
                return BrowserSlotLease(slot_index=slot_index, lock=lock)

        if deadline is not None and time.time() >= deadline:
            return None
        time.sleep(retry_interval)


def count_browser_slots_in_use() -> int:
    client = get_redis_client()
    in_use = 0
    for slot_index in range(1, BROWSER_MAX_CONCURRENCY + 1):
        if client.exists(f"{REDIS_BROWSER_SLOT_PREFIX}:{slot_index}"):
            in_use += 1
    return in_use


def acquire_queue_pop_lock(
    *,
    blocking_timeout_seconds: int | None = None,
    retry_interval_seconds: float | None = None,
) -> QueuePopLease | None:
    client = get_redis_client()
    retry_interval = retry_interval_seconds or min(HEARTBEAT_INTERVAL_SECONDS / 3, 2.0)
    deadline = time.time() + blocking_timeout_seconds if blocking_timeout_seconds is not None else None

    while True:
        lock = client.lock(
            REDIS_QUEUE_POP_LOCK_KEY,
            timeout=QUEUE_POP_LOCK_TTL_SECONDS,
            blocking=False,
            thread_local=False,
        )
        if lock.acquire(blocking=False):
            return QueuePopLease(lock=lock)

        if deadline is not None and time.time() >= deadline:
            return None
        time.sleep(retry_interval)


def clear_runtime_locks() -> int:
    client = get_redis_client()
    keys = [REDIS_QUEUE_POP_LOCK_KEY]
    keys.extend(f"{REDIS_BROWSER_SLOT_PREFIX}:{slot_index}" for slot_index in range(1, BROWSER_MAX_CONCURRENCY + 1))
    existing_keys = [key for key in keys if client.exists(key)]
    if not existing_keys:
        return 0
    return int(client.delete(*existing_keys))
