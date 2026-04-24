from __future__ import annotations

import json
import shutil
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile

from app.core.time_utils import beijing_now_display
from app.core.config import (
    EXAMPLE_MANIFESTS,
    JOBS_DIR,
    RECENT_LOG_LINE_COUNT,
    RESULTS_DIR,
    UPLOADS_DIR,
    ensure_app_directories,
)
from lingxing_rpa_runner import sanitize_filename_part


ALLOWED_MANIFEST_SUFFIXES = {".txt", ".xlsx"}


def now_display() -> str:
    return beijing_now_display()


def sanitize_upload_name(filename: str) -> str:
    source = Path(filename or "manifest.txt")
    suffix = source.suffix.lower()
    safe_stem = sanitize_filename_part(source.stem)
    if not suffix:
        suffix = ".txt"
    return f"{safe_stem}{suffix}"


def build_job_directories(task_id: str) -> dict[str, Path]:
    ensure_app_directories()
    job_dir = JOBS_DIR / task_id
    directories = {
        "job_dir": job_dir,
        "input": job_dir / "input",
        "downloads": job_dir / "downloads",
        "output": job_dir / "output",
        "logs": job_dir / "logs",
        "screenshots": job_dir / "screenshots",
        "reports": job_dir / "reports",
    }
    for path in directories.values():
        path.mkdir(parents=True, exist_ok=True)
    return directories


def save_uploaded_manifest(upload_file, task_id: str, input_dir: Path) -> tuple[Path, Path, str]:
    ensure_app_directories()
    original_filename = sanitize_upload_name(upload_file.filename or "manifest.txt")
    suffix = Path(original_filename).suffix.lower()
    if suffix not in ALLOWED_MANIFEST_SUFFIXES:
        raise ValueError("只支持上传 .txt 或 .xlsx 文件")

    upload_path = UPLOADS_DIR / f"{task_id}_{original_filename}"
    input_path = input_dir / original_filename
    with upload_path.open("wb") as upload_handle:
        shutil.copyfileobj(upload_file.file, upload_handle)
    shutil.copy2(upload_path, input_path)
    return upload_path, input_path, original_filename


def save_text_manifest(fba_text: str, task_id: str, input_dir: Path) -> tuple[Path, Path, str]:
    ensure_app_directories()
    original_filename = "pasted_fba_manifest.txt"
    upload_path = UPLOADS_DIR / f"{task_id}_{original_filename}"
    input_path = input_dir / original_filename
    normalized_lines = [line.strip().upper() for line in fba_text.splitlines() if line.strip()]
    content = "\n".join(normalized_lines).strip() + "\n"
    upload_path.write_text(content, encoding="utf-8")
    input_path.write_text(content, encoding="utf-8")
    return upload_path, input_path, original_filename


def cleanup_submission_files(job_dir: Path, upload_path: Path | None = None) -> None:
    if upload_path and upload_path.exists():
        upload_path.unlink()
    if job_dir.exists():
        shutil.rmtree(job_dir, ignore_errors=True)


def cleanup_task_artifacts(task: dict) -> None:
    upload_path = task.get("upload_path")
    if upload_path:
        upload_file = Path(upload_path)
        if upload_file.exists():
            upload_file.unlink()

    result_zip_path = task.get("result_zip_path")
    if result_zip_path:
        result_file = Path(result_zip_path)
        if result_file.exists():
            result_file.unlink()

    job_dir = task.get("job_dir")
    if job_dir:
        shutil.rmtree(job_dir, ignore_errors=True)


def append_log_line(log_path: Path, message: str) -> None:
    log_path.parent.mkdir(parents=True, exist_ok=True)
    with log_path.open("a", encoding="utf-8") as handle:
        handle.write(f"[{now_display()}] {message}\n")


def tail_text_file(path: Path | None, max_lines: int = RECENT_LOG_LINE_COUNT) -> str:
    if path is None or not path.exists():
        return ""
    lines = path.read_text(encoding="utf-8", errors="ignore").splitlines()
    return "\n".join(lines[-max_lines:])


def load_json_file(path: Path | None) -> dict | None:
    if path is None or not path.exists():
        return None
    return json.loads(path.read_text(encoding="utf-8"))


def locate_job_manifest(job_dir: Path) -> Path:
    input_dir = job_dir / "input"
    for path in sorted(input_dir.iterdir()):
        if path.is_file():
            return path
    raise FileNotFoundError(f"任务目录中未找到上传清单：{input_dir}")


def build_error_summary_text(batch_report: dict) -> str:
    lines = [
        f"批次状态：{batch_report.get('status', 'FAILED')}",
        f"成功 FBA 数：{batch_report.get('success_count', 0)}",
        f"失败 FBA 数：{batch_report.get('failed_count', 0)}",
        "",
    ]
    fatal_error = batch_report.get("fatal_error") or {}
    if fatal_error.get("error"):
        lines.extend(
            [
                "批次级错误：",
                str(fatal_error["error"]),
                "",
            ]
        )

    failed_results = [item for item in batch_report.get("results", []) if item.get("status") != "success"]
    if failed_results:
        lines.append("失败明细：")
        for item in failed_results:
            lines.append(f"- {item.get('fba_code')}: {item.get('error') or item.get('error_code') or '未知错误'}")
    else:
        lines.append("本批次没有失败的 FBA。")
    return "\n".join(lines).strip() + "\n"


def add_directory_to_zip(zip_file: ZipFile, base_dir: Path, arc_prefix: str) -> None:
    if not base_dir.exists():
        return
    for path in sorted(base_dir.rglob("*")):
        if path.is_file():
            zip_file.write(path, arcname=str(Path(arc_prefix) / path.relative_to(base_dir)))


def resolve_primary_result_file(job_dir: Path, batch_report: dict) -> Path | None:
    output_root = job_dir / "output"
    for item in batch_report.get("results", []):
        if item.get("status") != "success":
            continue
        workbook_name = item.get("processing_output_workbook")
        if not workbook_name:
            continue
        candidate = output_root / sanitize_filename_part(item.get("fba_code", "")) / workbook_name
        if candidate.exists():
            return candidate

    for candidate in sorted(output_root.rglob("*.xlsx")):
        return candidate
    return None


def create_result_zip(job_dir: Path, result_zip_path: Path, batch_report: dict, log_path: Path) -> Path:
    ensure_app_directories()
    reports_dir = job_dir / "reports"
    error_summary_path = reports_dir / "error_summary.txt"
    error_summary_path.write_text(build_error_summary_text(batch_report), encoding="utf-8")

    if result_zip_path.exists():
        result_zip_path.unlink()

    with ZipFile(result_zip_path, "w", compression=ZIP_DEFLATED) as archive:
        add_directory_to_zip(archive, job_dir / "output", "output")
        add_directory_to_zip(archive, reports_dir, "reports")
        add_directory_to_zip(archive, log_path.parent, "logs")
        if batch_report.get("status") != "success":
            add_directory_to_zip(archive, job_dir / "screenshots", "screenshots")
    return result_zip_path


def collect_output_workbooks(job_dir: Path) -> list[Path]:
    output_root = job_dir / "output"
    if not output_root.exists():
        return []
    return sorted(path for path in output_root.rglob("*.xlsx") if path.is_file() and not path.name.startswith("~$"))


def unique_archive_name(path: Path, used_names: set[str]) -> str:
    base_name = path.name
    if base_name not in used_names:
        used_names.add(base_name)
        return base_name

    parent_name = sanitize_filename_part(path.parent.name)
    candidate = f"{parent_name}_{base_name}"
    if candidate not in used_names:
        used_names.add(candidate)
        return candidate

    index = 2
    while True:
        candidate = f"{path.stem}_{index}{path.suffix}"
        if candidate not in used_names:
            used_names.add(candidate)
            return candidate
        index += 1


def create_user_result_download(job_dir: Path, result_path: Path) -> Path | None:
    ensure_app_directories()
    output_workbooks = collect_output_workbooks(job_dir)
    if not output_workbooks:
        return None

    if len(output_workbooks) == 1:
        return output_workbooks[0]

    if result_path.exists():
        result_path.unlink()

    folder_name = sanitize_filename_part(f"{job_dir.name}_结果文件")
    used_names: set[str] = set()
    with ZipFile(result_path, "w", compression=ZIP_DEFLATED) as archive:
        for workbook in output_workbooks:
            archive_name = unique_archive_name(workbook, used_names)
            archive.write(workbook, arcname=str(Path(folder_name) / archive_name))
    return result_path


def get_example_manifest_path(name: str) -> Path | None:
    return EXAMPLE_MANIFESTS.get(name)


def default_result_zip_path(task_id: str) -> Path:
    return RESULTS_DIR / f"{task_id}.zip"
