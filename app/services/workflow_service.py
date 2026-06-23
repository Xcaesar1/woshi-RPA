from __future__ import annotations

import json
import re
from pathlib import Path

from app.core.config import (
    AMAZON_HL_WORKFLOW_NAME,
    BROWSER_MAX_CONCURRENCY,
    BROWSER_PROFILE_DIR,
    DEFAULT_CONFIG_PATH,
    LINGXING_WORKFLOW_NAME,
    RESOURCE_DIR,
    RESULTS_DIR,
    WORKFLOW_REGISTRY,
)
from app.core.time_utils import beijing_now_display, format_datetime_for_display
from app.models.task import (
    TASK_STATUS_FAILED,
    TASK_STATUS_LABELS,
    TERMINAL_TASK_STATUSES,
    normalize_batch_status,
)
from app.services.file_service import (
    ALLOWED_AMAZON_HL_SUFFIXES,
    append_log_line,
    build_job_directories,
    cleanup_task_artifacts,
    cleanup_submission_files,
    create_user_result_download,
    default_result_zip_path,
    load_json_file,
    locate_job_manifest,
    locate_job_manifests,
    resolve_primary_result_file,
    save_uploaded_manifests,
    save_text_manifest,
    save_uploaded_manifest,
    tail_text_file,
)
from app.services.amazon_hl_csv_service import (
    AmazonHlShipment,
    convert_amazon_hl_shipment_to_source_workbook,
    parse_amazon_hl_csv_shipments,
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
from lingxing_excel_processor import process_workbooks
from lingxing_rpa_runner import parse_manifest_file, run_manifest_job


FBA_TEXT_TOKEN_RE = re.compile(r"[A-Za-z0-9-]+")
FBA_CODE_RE = re.compile(r"^FBA[A-Z0-9-]+$")
LOG_TIMESTAMP_RE = re.compile(r"\[(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2})\]")
TASK_TIME_FIELDS = ("created_at", "queued_at", "started_at", "heartbeat_at", "finished_at")


def get_workflow_options() -> list[dict[str, str]]:
    return [{"name": item.name, "label": item.label} for item in WORKFLOW_REGISTRY.values()]


def validate_workflow_name(workflow_name: str) -> str:
    if workflow_name not in WORKFLOW_REGISTRY:
        raise ValueError("不支持的流程类型")
    return workflow_name


def validate_hl_upload_filename(filename: str | None) -> None:
    suffix = Path(filename or "").suffix.lower()
    if suffix not in ALLOWED_AMAZON_HL_SUFFIXES:
        raise ValueError("HL 发货请上传 Amazon 后台导出的 .csv 文件")


def task_requires_browser(task: dict) -> bool:
    return task.get("workflow_name") != AMAZON_HL_WORKFLOW_NAME


def parse_hl_shipments_from_paths(paths: list[Path]) -> list[AmazonHlShipment]:
    shipments: list[AmazonHlShipment] = []
    for path in paths:
        try:
            shipments.extend(parse_amazon_hl_csv_shipments(path))
        except Exception as exc:
            raise ValueError(f"{path.name} 解析失败：{exc}") from exc

    if not shipments:
        raise ValueError("HL 发货 CSV 中未解析到任何 FBA")

    seen: set[str] = set()
    duplicated: list[str] = []
    for shipment in shipments:
        if shipment.fba_code in seen:
            duplicated.append(shipment.fba_code)
        seen.add(shipment.fba_code)
    if duplicated:
        raise ValueError(f"HL 发货 CSV 中存在重复 FBA：{', '.join(dict.fromkeys(duplicated))}")
    return shipments


def build_task_view(task: dict) -> dict:
    view = dict(task)
    for field in TASK_TIME_FIELDS:
        if field in view:
            view[field] = format_datetime_for_display(view.get(field))
    status = task.get("status", TASK_STATUS_FAILED)
    view["status_label"] = TASK_STATUS_LABELS.get(status, status)
    view["workflow_label"] = WORKFLOW_REGISTRY.get(task.get("workflow_name"), WORKFLOW_REGISTRY[next(iter(WORKFLOW_REGISTRY))]).label
    result_zip_path = task.get("result_zip_path")
    view["can_download"] = bool(result_zip_path and Path(result_zip_path).exists() and status in TERMINAL_TASK_STATUSES)
    view["detail_url"] = f"/tasks/{task['id']}"
    view["download_url"] = f"/api/tasks/{task['id']}/download"
    fba_codes = resolve_task_fba_codes(task)
    view["fba_codes"] = fba_codes
    view["task_display_id"] = format_task_display_id(fba_codes, fallback=task["id"])
    view["internal_task_id"] = task["id"]
    return view


def resolve_task_fba_codes(task: dict) -> list[str]:
    job_dir_value = task.get("job_dir")
    if not job_dir_value:
        return []

    job_dir = Path(job_dir_value)
    batch_report = load_json_file(job_dir / "reports" / "batch_report.json")
    report_codes = batch_report.get("fba_codes", []) if batch_report else []
    if report_codes:
        return [str(code).strip().upper() for code in report_codes if str(code).strip()]

    try:
        manifest_paths = locate_job_manifests(job_dir)
        if manifest_paths and all(path.suffix.lower() == ".csv" for path in manifest_paths):
            shipments = parse_hl_shipments_from_paths(manifest_paths)
            return [shipment.fba_code for shipment in shipments]
        manifest_path = manifest_paths[0]
        return parse_manifest_file(manifest_path)
    except Exception:
        return []


def format_task_display_id(fba_codes: list[str], *, fallback: str) -> str:
    if not fba_codes:
        return fallback
    if len(fba_codes) == 1:
        return fba_codes[0]
    return f"{fba_codes[0]} 等 {len(fba_codes)} 个"


def is_legacy_utc_task_time(value: object) -> bool:
    text = str(value or "")
    if "T" not in text:
        return False
    timezone_part = text[19:]
    return "+" not in timezone_part and "-" not in timezone_part and not text.endswith("Z")


def format_recent_log_for_display(log_text: str, *, legacy_utc: bool) -> str:
    if not legacy_utc or not log_text:
        return log_text

    def replace_match(match: re.Match[str]) -> str:
        converted = format_datetime_for_display(match.group(1).replace(" ", "T"))
        return f"[{converted or match.group(1)}]"

    return LOG_TIMESTAMP_RE.sub(replace_match, log_text)


def list_task_views(*, submitter: str | None = None, status: str | None = None) -> list[dict]:
    return [build_task_view(task) for task in list_tasks(submitter=submitter, status=status)]


def parse_fba_text_input(fba_text: str | None) -> list[str]:
    text = (fba_text or "").strip()
    if not text:
        return []

    fba_codes: list[str] = []
    invalid_tokens: list[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#"):
            continue
        tokens = FBA_TEXT_TOKEN_RE.findall(line)
        for token in tokens:
            normalized = token.upper().strip()
            if FBA_CODE_RE.fullmatch(normalized):
                fba_codes.append(normalized)
            else:
                invalid_tokens.append(token)

    if invalid_tokens:
        preview = ", ".join(invalid_tokens[:5])
        raise ValueError(f"FBA号格式不正确，请检查：{preview}")
    return list(dict.fromkeys(fba_codes))


def create_task_submission(
    *,
    manifest_upload=None,
    manifest_uploads=None,
    fba_text: str | None = None,
    workflow_name: str,
    submitter: str,
    remark: str | None = None,
) -> dict:
    workflow_name = validate_workflow_name(workflow_name)
    submitter = (submitter or "").strip()
    remark = (remark or "").strip() or None
    if not submitter:
        raise ValueError("提交人不能为空")
    pasted_fba_codes = parse_fba_text_input(fba_text)
    hl_uploads = [upload for upload in ([manifest_upload] + list(manifest_uploads or [])) if upload is not None]
    if workflow_name == AMAZON_HL_WORKFLOW_NAME:
        if not hl_uploads:
            raise ValueError("HL 发货请上传 Amazon 后台导出的 .csv 文件")
        for upload in hl_uploads:
            validate_hl_upload_filename(upload.filename)
    elif workflow_name == LINGXING_WORKFLOW_NAME:
        if not pasted_fba_codes:
            raise ValueError("正常/UPS 流程请直接粘贴 FBA 号，一行一个")
    elif not manifest_upload and not pasted_fba_codes:
        raise ValueError("请粘贴 FBA 号，或上传清单文件")

    task_id = generate_task_id()
    job_paths = build_job_directories(task_id)
    upload_path: Path | None = None
    input_manifest_path: Path | None = None
    task_created = False
    log_path = job_paths["logs"] / "task.log"
    log_path.touch()

    try:
        if workflow_name == AMAZON_HL_WORKFLOW_NAME:
            upload_path, input_manifest_paths, original_filename = save_uploaded_manifests(
                hl_uploads,
                task_id=task_id,
                input_dir=job_paths["input"],
                allowed_suffixes=ALLOWED_AMAZON_HL_SUFFIXES,
                invalid_message="HL 发货请上传 Amazon 后台导出的 .csv 文件",
            )
            append_log_line(log_path, f"任务已创建，开始校验 {len(input_manifest_paths)} 个 Amazon HL CSV 文件")
            shipments = parse_hl_shipments_from_paths(input_manifest_paths)
            fba_codes = [shipment.fba_code for shipment in shipments]
        elif pasted_fba_codes:
            upload_path, input_manifest_path, original_filename = save_text_manifest(
                "\n".join(pasted_fba_codes),
                task_id=task_id,
                input_dir=job_paths["input"],
            )
            append_log_line(log_path, "任务已创建，开始校验粘贴的 FBA 号")
            fba_codes = pasted_fba_codes
        elif manifest_upload:
            upload_path, input_manifest_path, original_filename = save_uploaded_manifest(
                manifest_upload,
                task_id=task_id,
                input_dir=job_paths["input"],
            )
            append_log_line(log_path, "任务已创建，开始校验上传清单")
            fba_codes = parse_manifest_file(input_manifest_path)
        else:
            raise ValueError("请粘贴 FBA 号，或上传清单文件")
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


def process_amazon_hl_csv_task(task: dict, log) -> dict:
    job_dir = Path(task["job_dir"])
    manifest_paths = locate_job_manifests(job_dir, allowed_suffixes=ALLOWED_AMAZON_HL_SUFFIXES)
    reports_root = job_dir / "reports"
    reports_root.mkdir(parents=True, exist_ok=True)
    source_dir = job_dir / "downloads" / "amazon_hl"
    output_dir = job_dir / "output" / "amazon_hl"
    started_at = beijing_now_display()

    shipments = parse_hl_shipments_from_paths(manifest_paths)
    results: list[dict] = []
    converted_results: list[dict] = []

    for sequence, shipment in enumerate(shipments, start=1):
        fba_code = shipment.fba_code
        safe_fba = re.sub(r"[^A-Za-z0-9._-]+", "_", fba_code).strip("_") or fba_code
        report_path = reports_root / f"{safe_fba}_automation_report.json"
        result = {
            "fba_code": fba_code,
            "status": "pending",
            "started_at": started_at,
            "fba_root": str(job_dir / safe_fba),
            "downloads_dir": str(source_dir),
            "output_dir": str(output_dir),
            "screenshots_dir": str(job_dir / "screenshots" / safe_fba),
            "downloaded_files": [],
            "processing_output_workbook": None,
            "processing_output_files": [],
            "processing_report_file": None,
            "error_code": None,
            "error": None,
            "traceback": None,
            "failure_screenshot": None,
            "amazon_hl_summary": {
                "cargo_name": shipment.cargo_name,
                "destination": shipment.destination,
                "sku_count": shipment.sku_count,
                "carton_count": shipment.carton_count,
                "total_quantity": shipment.total_quantity,
            },
        }
        results.append(result)
        log(f"[{fba_code}] 开始解析 Amazon HL CSV")
        try:
            source_workbook, _ = convert_amazon_hl_shipment_to_source_workbook(
                shipment,
                source_dir,
                resource_dir=RESOURCE_DIR,
            )
            result["downloaded_files"] = [
                {
                    "sequence": sequence,
                    "warehouse_code": "AMAZON-HL",
                    "path": str(source_workbook),
                    "filename": source_workbook.name,
                    "source": "amazon_csv",
                }
            ]
            converted_results.append(result)
            log(f"[{fba_code}] CSV 已转换为标准源表")
        except Exception as exc:
            import traceback

            result["status"] = "failed"
            result["error_code"] = "amazon_hl_csv_error"
            result["error"] = str(exc)
            result["traceback"] = traceback.format_exc()
            log(f"[{fba_code}] HL CSV 转换失败：{exc}")

    if converted_results:
        log("开始整理 HL CSV 批次 Excel")
        try:
            process_report = process_workbooks(RESOURCE_DIR, source_dir, output_dir)
            for result in converted_results:
                result["processing_output_workbook"] = process_report.get("output_workbook")
                result["processing_output_files"] = process_report.get("processing_output_files", [])
                result["processing_report_file"] = process_report.get("report_file")
                result["processing_anomalies"] = process_report.get("anomalies", [])
                result["status"] = "success"
                log(f"[{result['fba_code']}] HL CSV 整理完成")
        except Exception as exc:
            import traceback

            error_traceback = traceback.format_exc()
            for result in converted_results:
                result["status"] = "failed"
                result["error_code"] = "amazon_hl_csv_process_error"
                result["error"] = str(exc)
                result["traceback"] = error_traceback
            log(f"HL CSV 批次整理失败：{exc}")

    finished_at = beijing_now_display()
    for result in results:
        result["finished_at"] = finished_at
        safe_fba = re.sub(r"[^A-Za-z0-9._-]+", "_", result["fba_code"]).strip("_") or result["fba_code"]
        (reports_root / f"{safe_fba}_automation_report.json").write_text(
            json.dumps(result, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    success_count = sum(1 for result in results if result.get("status") == "success")
    failed_count = len(results) - success_count
    if success_count and failed_count:
        batch_status = "partial_success"
    elif success_count:
        batch_status = "success"
    else:
        batch_status = "failed"
    batch_report = {
        "batch_dir": str(job_dir),
        "manifest_path": str(manifest_paths[0]),
        "manifest_paths": [str(path) for path in manifest_paths],
        "resource_dir": str(RESOURCE_DIR),
        "work_dir": str(job_dir.parent),
        "config_file": None,
        "started_at": started_at,
        "finished_at": finished_at,
        "fba_codes": [shipment.fba_code for shipment in shipments],
        "success_count": success_count,
        "failed_count": failed_count,
        "status": batch_status,
        "results": results,
        "fatal_error": None,
        "source": "amazon_hl_csv",
    }
    (reports_root / "batch_report.json").write_text(json.dumps(batch_report, ensure_ascii=False, indent=2), encoding="utf-8")
    log(f"HL CSV 批次完成，状态：{batch_report['status']}")
    return batch_report


def simplify_failure_reason(error_code: str | None, error: str | None) -> str:
    code = (error_code or "").strip()
    text = (error or "").strip()
    lower_text = text.lower()

    if code == "shipment_not_found" or "未搜索到" in text:
        return "领星里没有查到这个 FBA。请先确认 FBA 号是否填错, 或这个 FBA 是否属于当前账号能看到的店铺。"
    if code in {"shipment_search_timeout", "shipment_detail_timeout"} or "详情页超时" in text or "搜索" in text and "超时" in text:
        return "领星页面打开或搜索太慢, 这次没有成功进入详情页。可以稍后重试一次。"
    if code in {"export_response_not_excel", "download_not_excel"} or "不是 excel" in lower_text:
        return "领星没有返回正常的 Excel 文件。通常是这个货件还没准备好, 或页面上的导出还不能用。"
    if code in {"export_modal_not_found", "download_dialog_not_found"}:
        return "点击下载后没有出现导出确认框。可能是领星页面变了, 或当前货件状态还不能下载。"
    if code in {"download_button_not_found", "shipment_cards_not_found"}:
        return "没有找到下载按钮。请确认这个货件已经进入箱子标签页面, 并且页面里有仓库卡片。"
    if code == "home_or_login_page_not_ready":
        return "领星页面这次没有正常加载出来。系统会自动重试；如果连续出现, 请稍后再提交或联系负责人检查网络。"
    if "ssl" in lower_text or "unexpected_eof" in lower_text or "eof occurred" in lower_text:
        return "下载时领星连接中断了一次。通常是临时网络问题, 重新提交这几个 FBA 即可。"
    if code in {"login_failed", "login_fields_not_found", "credentials_missing"} or "登录" in text:
        return "领星登录没有成功。请联系负责人检查账号密码, 或确认当前是否需要重新登录。"
    if code == "shipment_not_ready_for_box_labels" or "箱子标签" in text:
        return "这个货件还没到可以下载箱子标签的步骤。等领星里进入箱子标签后再提交。"
    if "bad offset" in lower_text or "not a zip file" in lower_text:
        return "下载到的 Excel 文件不完整。可以先重试一次, 如果连续失败再联系负责人。"
    if text:
        return text
    return "任务执行时遇到问题。可以先重试一次, 如果还失败再把详情页截图发给负责人。"


def build_friendly_failure_notice(task_view: dict, fba_results: list[dict], batch_report: dict | None) -> dict | None:
    status = task_view.get("status")
    if status not in {"FAILED", "PARTIAL_SUCCESS"}:
        return None

    failed_items = [
        {
            "fba_code": item.get("fba_code") or "-",
            "reason": simplify_failure_reason(item.get("error_code"), item.get("error")),
        }
        for item in fba_results
        if item.get("status") != "SUCCESS"
    ]

    fatal_error = (batch_report or {}).get("fatal_error") or {}
    if not failed_items and fatal_error.get("error"):
        failed_items.append(
            {
                "fba_code": task_view.get("task_display_id") or task_view.get("id") or "-",
                "reason": simplify_failure_reason(fatal_error.get("error_code"), fatal_error.get("error")),
            }
        )

    if status == "PARTIAL_SUCCESS":
        title = "有一部分 FBA 没处理成功"
        message = "已经成功的文件可以先下载使用。没成功的 FBA 请按下面原因检查后再重新提交。"
    else:
        title = "这次任务没有处理成功"
        message = "系统已经尝试打开领星并执行下载, 但中途遇到了下面的问题。"

    if not failed_items:
        failed_items.append(
            {
                "fba_code": task_view.get("task_display_id") or task_view.get("id") or "-",
                "reason": simplify_failure_reason(None, task_view.get("error_message")),
            }
        )

    return {
        "title": title,
        "message": message,
        "failed_items": failed_items[:8],
        "suggestions": [
            "先检查 FBA 号有没有填错。",
            "如果你手动在领星也查不到, 说明这个 FBA 当前不能处理。",
            "如果你手动能查到, 把这个详情页截图发给负责人。",
        ],
    }


def process_task(task: dict) -> dict:
    task_id = task["id"]
    job_dir = Path(task["job_dir"])
    log_path = Path(task["log_path"])

    def log(message: str) -> None:
        append_log_line(log_path, message)

    try:
        log("任务进入运行中")
        manifest_path = locate_job_manifest(job_dir)
        if task.get("workflow_name") == AMAZON_HL_WORKFLOW_NAME:
            batch_report = process_amazon_hl_csv_task(task, log)
        else:
            batch_report = run_manifest_job(
                manifest_path=manifest_path,
                resource_dir=RESOURCE_DIR,
                job_dir=job_dir,
                profile_dir=BROWSER_PROFILE_DIR,
                config_path=DEFAULT_CONFIG_PATH if DEFAULT_CONFIG_PATH.exists() else None,
                log_callback=log,
            )
        result_download_path = create_user_result_download(
            job_dir=job_dir,
            result_path=default_result_zip_path(task_id),
        )
        primary_result_path = resolve_primary_result_file(job_dir, batch_report)
        final_status = normalize_batch_status(batch_report.get("status"))
        error_message = build_task_error_message(batch_report)
        mark_task_finished(
            task_id=task_id,
            status=final_status,
            result_zip_path=str(result_download_path) if result_download_path else None,
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
    recent_log = format_recent_log_for_display(recent_log, legacy_utc=is_legacy_utc_task_time(task.get("created_at")))
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
                "started_at": format_datetime_for_display(batch_report.get("started_at") if batch_report else task.get("started_at")),
                "finished_at": format_datetime_for_display(batch_report.get("finished_at") if batch_report else task.get("finished_at")),
                "fatal_error": (batch_report.get("fatal_error") or {}).get("error") if batch_report else None,
            },
            "fba_results": fba_results,
            "friendly_error": build_friendly_failure_notice(task_view, fba_results, batch_report),
        }
    )
    return task_view
