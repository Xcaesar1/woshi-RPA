from __future__ import annotations

from contextlib import asynccontextmanager

from fastapi import FastAPI
from fastapi.staticfiles import StaticFiles

from app.api.pages import router as pages_router
from app.api.tasks import router as tasks_router
from app.core.config import ensure_app_directories
from app.core.db import init_db


@asynccontextmanager
async def lifespan(_: FastAPI):
    ensure_app_directories()
    init_db()
    yield


app = FastAPI(title="领星内网页面任务系统", lifespan=lifespan)
app.mount("/static", StaticFiles(directory="app/static"), name="static")
app.include_router(pages_router)
app.include_router(tasks_router)


@app.get("/healthz")
def healthz() -> dict[str, str]:
    return {"status": "ok"}
