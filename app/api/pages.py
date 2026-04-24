from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path

from fastapi import APIRouter, Request
from fastapi.responses import RedirectResponse
from fastapi.templating import Jinja2Templates

from app.core.config import APP_DIR
from app.models.task import TASK_STATUS_CHOICES, TASK_STATUS_LABELS
from app.services.workflow_service import get_system_status, get_task_detail, get_workflow_options, list_task_views


router = APIRouter()
templates = Jinja2Templates(directory="app/templates")
UI_MANIFEST_PATH = APP_DIR / "static" / "ui" / ".vite" / "manifest.json"


@lru_cache(maxsize=1)
def load_ui_manifest() -> dict[str, dict]:
    if not UI_MANIFEST_PATH.exists():
        return {}
    with UI_MANIFEST_PATH.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def collect_entry_css_urls(entry_name: str) -> list[str]:
    manifest = load_ui_manifest()
    entry_key = f"src/entries/{entry_name}.tsx"
    visited: set[str] = set()
    css_files: list[str] = []

    def visit(key: str) -> None:
        if key in visited:
            return
        visited.add(key)
        item = manifest.get(key, {})
        for css_file in item.get("css", []):
            if css_file not in css_files:
                css_files.append(css_file)
        for import_key in item.get("imports", []):
            visit(import_key)

    visit(entry_key)
    version = int(UI_MANIFEST_PATH.stat().st_mtime) if UI_MANIFEST_PATH.exists() else 0
    return [f"/static/ui/{css_file}?v={version}" for css_file in css_files]


def get_entry_script_url(entry_name: str) -> str:
    manifest = load_ui_manifest()
    manifest_key = f"src/entries/{entry_name}.tsx"
    manifest_entry = manifest.get(manifest_key, {})
    file_name = manifest_entry.get("file", f"{entry_name}.js")
    version = int(UI_MANIFEST_PATH.stat().st_mtime) if UI_MANIFEST_PATH.exists() else 0
    return f"/static/ui/{file_name}?v={version}"


def build_template_response(request: Request, template_name: str, *, page_title: str, entry_name: str, page_data: dict):
    response = templates.TemplateResponse(
        request,
        template_name,
        {
            "page_title": page_title,
            "entry_script_url": get_entry_script_url(entry_name),
            "style_urls": collect_entry_css_urls(entry_name),
            "page_data": page_data,
        },
    )
    response.headers["Cache-Control"] = "no-store"
    return response


@router.get("/", include_in_schema=False)
def home() -> RedirectResponse:
    return RedirectResponse(url="/tasks/new", status_code=302)


@router.get("/tasks/new")
def new_task_page(request: Request):
    return build_template_response(
        request,
        "task_new.html",
        page_title="新建任务",
        entry_name="task-new",
        page_data={
            "workflows": get_workflow_options(),
            "example_files": ["fba_manifest.txt", "fba_manifest.xlsx"],
        },
    )


@router.get("/tasks")
def task_list_page(request: Request, submitter: str | None = None, status: str | None = None):
    return build_template_response(
        request,
        "task_list.html",
        page_title="任务列表",
        entry_name="task-list",
        page_data={
            "system_status": get_system_status(),
            "tasks": list_task_views(submitter=submitter, status=status),
            "submitter": submitter or "",
            "status": status or "",
            "status_choices": [(item, TASK_STATUS_LABELS.get(item, item)) for item in TASK_STATUS_CHOICES],
        },
    )


@router.get("/tasks/{task_id}")
def task_detail_page(request: Request, task_id: str):
    detail = get_task_detail(task_id)
    return build_template_response(
        request,
        "task_detail.html",
        page_title=f"任务详情 - {task_id}",
        entry_name="task-detail",
        page_data={
            "task": detail,
        },
    )
