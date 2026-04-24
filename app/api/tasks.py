from __future__ import annotations

from pathlib import Path

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.core.config import EXAMPLE_MANIFESTS
from app.models.task import TASK_STATUS_CHOICES, TERMINAL_TASK_STATUSES
from app.services.workflow_service import create_task_submission, get_system_status, get_task_detail, list_task_views


router = APIRouter(prefix="/api")


@router.post("/tasks")
def create_task_api(
    manifest_file: UploadFile = File(...),
    workflow_name: str = Form(...),
    submitter: str = Form(...),
    remark: str = Form(""),
):
    try:
        task = create_task_submission(
            manifest_upload=manifest_file,
            workflow_name=workflow_name,
            submitter=submitter,
            remark=remark,
        )
    except ValueError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc
    finally:
        manifest_file.file.close()

    return {
        "task": task,
        "redirect_url": task["detail_url"],
    }


@router.get("/tasks")
def list_tasks_api(submitter: str | None = None, status: str | None = None):
    if status and status not in TASK_STATUS_CHOICES:
        raise HTTPException(status_code=400, detail="不支持的任务状态")
    return {
        "tasks": list_task_views(submitter=submitter, status=status),
    }


@router.get("/tasks/{task_id}")
def get_task_api(task_id: str):
    try:
        task = get_task_detail(task_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="任务不存在") from exc
    return task


@router.get("/system/status")
def get_system_status_api():
    return get_system_status()


@router.get("/tasks/{task_id}/download")
def download_task_result(task_id: str):
    try:
        task = get_task_detail(task_id)
    except KeyError as exc:
        raise HTTPException(status_code=404, detail="任务不存在") from exc

    result_zip_path = task.get("result_zip_path")
    if task.get("status") not in TERMINAL_TASK_STATUSES or not result_zip_path:
        raise HTTPException(status_code=400, detail="任务尚未生成可下载结果")

    file_path = Path(result_zip_path)
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="结果文件不存在")
    return FileResponse(file_path, filename=file_path.name)


@router.get("/examples/{filename}")
def download_example_file(filename: str):
    example_path = EXAMPLE_MANIFESTS.get(filename)
    if example_path is None or not example_path.exists():
        raise HTTPException(status_code=404, detail="示例文件不存在")
    return FileResponse(example_path, filename=example_path.name)
