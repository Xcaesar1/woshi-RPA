from __future__ import annotations

from pathlib import Path

from app.core.config import (
    BROWSER_MAX_CONCURRENCY,
    BROWSER_PROFILE_DIR,
    DEFAULT_CONFIG_PATH,
    RESOURCE_DIR,
    RESULTS_DIR,
    WORKFLOW_REGISTRY,
)
from app.models.task import (
    TASK_STATUS_FAILED,
    TASK_STATUS_LABELS,
    TERMINAL_TASK_STATUSES,
    normalize_batch_status,
)
from app.services.file_service import (
    append_log_line,
    build_job_directories,
    cleanup_task_artifacts,
    cleanup_submission_files,
    create_result_zip,
    default_result_zip_path,
    load_json_file,
    locate_job_manifest,
    resolve_primary_result_file,
    save_uploaded_manifest,
    tail_text_file,
)
from app.services.queue_service import (
    count_browser_slots_in_use,
    enqueue_task,
    is_any_worker_alive,
    latest_worker_heartbeat,
    queue_depth,
    remove_task_from_queue,
)
from app.services.task_service import (
    count_tasks,
    create_task,
    delete_tasks,
    generate_task_id,
    get_task,
    list_tasks,
    mark_task_failed,
    mark_task_finished,
)
from lingxing_rpa_runner import parse_manifest_file, run_manifest_job


def get_workflow_options() -> list[dict[str, str]]:
    return [{"name": item.name, "label": item.label} for item in WORKFLOW_REGISTRY.values()]


def validate_workflow_name(workflow_name: str) -> str:
    if workflow_name not in WORKFLOW_REGISTRY:
        raise ValueError("不支持的流程类型")
    return workflow_name


def build_task_view(task: dict) -> dict:
    view = dict(task)
    status = task.get("status", TASK_STATUS_FAILED)
    view["status_label"] = TASK_STATUS_LABELS.get(status, status)
    view["workflow_label"] = WORKFLOW_REGISTRY.get(task.get("workflow_name"), WORKFLOW_REGISTRY[next(iter(WORKFLOW_REGISTRY))]).label
    result_zip_path = task.get("result_zip_path")
    view["can_download"] = bool(result_zip_path and Path(result_zip_path).exists() and status in TERMINAL_TASK_STATUSES)
    view["detail_url"] = f"/tasks/{task['id']}"
    view["download_url"] = f"/api/tasks/{task['id']}/download"
    return view


def list_task_views(*, submitter: str | None = None, status: str | None = None) -> list[dict]:
    return [build_task_view(task) for task in list_tasks(submitter=submitter, status=status)]


def create_task_submission(*, manifest_upload, workflow_name: str, submitter: str, remark: str | None) -> dict:
    workflow_name = validate_workflow_name(workflow_name)
    submitter = (submitter or "").strip()
    remark = (remark or "").strip() or None
    if not submitter:
        raise ValueError("提交人不能为空")

    task_id = generate_task_id()
    job_paths = build_job_directories(task_id)
    upload_path: Path | None = None
    input_manifest_path: Path | None = None
    task_created = False
    log_path = job_paths["logs"] / "task.log"
    log_path.touch()

    try:
        upload_path, input_manifest_path, original_filename = save_uploaded_manifest(
            manifest_upload,
            task_id=task_id,
            input_dir=job_paths["input"],
        )
        append_log_line(log_path, "任务已创建，开始校验上传清单")
        fba_codes = parse_manifest_file(input_manifest_path)
        if not fba_codes:
            raise ValueError("清单中未解析到任何 FBA 号")
        append_log_line(log_path, f"清单校验通过，共解析到 {len(fba_codes)} 个 FBA，任务已入队")
        task = create_task(
            task_id=task_id,
            workflow_name=workflow_name,
            original_filename=original_filename,
            submitter=submitter,
            remark=remark,
            upload_path=str(upload_path),
            job_dir=str(job_paths["job_dir"]),
            log_path=str(log_path),
            total_fba_count=len(fba_codes),
        )
        task_created = True
        enqueue_task(task_id)
        return build_task_view(task)
    except Exception:
        if task_created:
            try:
                remove_task_from_queue(task_id)
            except Exception:
                pass
            delete_tasks([task_id])
        cleanup_submission_files(job_paths["job_dir"], upload_path)
        raise


def build_task_error_message(batch_report: dict) -> str | None:
    fatal_error = batch_report.get("fatal_error") or {}
    if fatal_error.get("error"):
        return str(fatal_error["error"])

    failed_results = [item for item in batch_report.get("results", []) if item.get("status") != "success"]
    if failed_results:
        return f"共有 {len(failed_results)} 个 FBA 执行失败"
    return None


def process_task(task: dict) -> dict:
    task_id = task["id"]
    job_dir = Path(task["job_dir"])
    log_path = Path(task["log_path"])

    def log(message: str) -> None:
        append_log_line(log_path, message)

    try:
        log("任务进入运行中")
        manifest_path = locate_job_manifest(job_dir)
        batch_report = run_manifest_job(
            manifest_path=manifest_path,
            resource_dir=RESOURCE_DIR,
            job_dir=job_dir,
            profile_dir=BROWSER_PROFILE_DIR,
            config_path=DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else None,
            log_callback=log,
        )
        result_zip_path = create_result_zip(
            job_dir=job_dir,
            result_zip_path=default_result_zip_path(task_id),
            batch_report=batch_report,
            log_path=log_path,
        )
        primary_result_path = resolve_primary_result_file(job_dir, batch_report)
        final_status = normalize_batch_status(batch_report.get("status"))
        error_message = build_task_error_message(batch_report)
        mark_task_finished(
            task_id=task_id,
            status=final_status,
            result_zip_path=str(result_zip_path),
            result_primary_file=str(primary_result_path) if primary_result_path else None,
            error_message=error_message,
            total_fba_count=len(batch_report.get("fba_codes", [])),
            success_fba_count=int(batch_report.get("success_count", 0)),
            failed_fba_count=int(batch_report.get("failed_count", 0)),
        )
        log(f"任务执行结束，状态：{TASK_STATUS_LABELS.get(final_status, final_status)}")
    except Exception as exc:
        log(f"任务执行异常：{exc}")
        mark_task_failed(task_id, str(exc))
    return get_task_detail(task_id)


def get_system_status() -> dict[str, object]:
    try:
        browser_slots_in_use = count_browser_slots_in_use()
        queue_depth_value = queue_depth()
        worker_alive = is_any_worker_alive()
        worker_recent_heartbeat = latest_worker_heartbeat()
        error_message = None
    except Exception as exc:
        browser_slots_in_use = 0
        queue_depth_value = 0
        worker_alive = False
        worker_recent_heartbeat = None
        error_message = str(exc)
    return {
        "queued_count": count_tasks(status="QUEUED"),
        "running_count": count_tasks(status="RUNNING"),
        "queue_depth": queue_depth_value,
        "worker_alive": worker_alive,
        "worker_recent_heartbeat": worker_recent_heartbeat,
        "browser_slots_total": BROWSER_MAX_CONCURRENCY,
        "browser_slots_in_use": browser_slots_in_use,
        "redis_error": error_message,
    }


def cleanup_expired_tasks(expired_tasks: list[dict]) -> int:
    if not expired_tasks:
        return 0
    for task in expired_tasks:
        cleanup_task_artifacts(task)
    delete_tasks([task["id"] for task in expired_tasks])
    return len(expired_tasks)


def get_task_detail(task_id: str) -> dict:
    task = get_task(task_id)
    if task is None:
        raise KeyError(task_id)

    task_view = build_task_view(task)
    job_dir = Path(task["job_dir"])
    batch_report = load_json_file(job_dir / "reports" / "batch_report.json")
    recent_log = tail_text_file(Path(task["log_path"])) if task.get("log_path") else ""
    recent_log_lines = [line for line in recent_log.splitlines() if line.strip()]
    current_stage = recent_log_lines[-1] if recent_log_lines else task_view["status_label"]
    fba_results = []
    if batch_report:
        for item in batch_report.get("results", []):
            normalized_status = normalize_batch_status(item.get("status"))
            fba_results.append(
                {
                    "fba_code": item.get("fba_code"),
                    "status": normalized_status,
                    "status_label": TASK_STATUS_LABELS.get(normalized_status),
                    "warehouse_count": item.get("warehouse_count"),
                    "download_count": len(item.get("downloaded_files", [])),
                    "output_workbook": item.get("processing_output_workbook"),
                    "report_file": item.get("processing_report_file"),
                    "error": item.get("error"),
                    "error_code": item.get("error_code"),
                }
            )

    task_view.update(
        {
            "current_stage": current_stage,
            "recent_log": recent_log,
            "recent_log_lines": recent_log_lines,
            "batch_report_summary": {
                "status": batch_report.get("status") if batch_report else task_view["status"],
                "success_count": batch_report.get("success_count") if batch_report else task.get("success_fba_count"),
                "failed_count": batch_report.get("failed_count") if batch_report else task.get("failed_fba_count"),
                "fba_codes": batch_report.get("fba_codes", []) if batch_report else [],
                "started_at": batch_report.get("started_at") if batch_report else task.get("started_at"),
                "finished_at": batch_report.get("finished_at") if batch_report else task.get("finished_at"),
                "fatal_error": (batch_report.get("fatal_error") or {}).get("error") if batch_report else None,
            },
            "fba_results": fba_results,
        }
    )
    return task_view
