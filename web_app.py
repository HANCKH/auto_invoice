from __future__ import annotations

import json
import logging
import os
import re
import shutil
import threading
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from starlette.requests import Request

from invoice_to_excel import process_invoice_batch


BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "模板文件.xlsx"
TASKS_DIR = BASE_DIR / "work" / "tasks"
STATS_PATH = BASE_DIR / "work" / "stats.json"
STATIC_DIR = BASE_DIR / "static"
TEMPLATES_DIR = BASE_DIR / "templates"

MAX_FILE_SIZE = int(os.getenv("INVOICE_MAX_FILE_SIZE", 20 * 1024 * 1024))
MAX_UPLOAD_SIZE = int(os.getenv("INVOICE_MAX_UPLOAD_SIZE", 200 * 1024 * 1024))
TASK_RETENTION_SECONDS = int(os.getenv("INVOICE_TASK_RETENTION_SECONDS", 24 * 60 * 60))
CHUNK_SIZE = 1024 * 1024
ALLOWED_MIME_TYPES = {"application/pdf", "application/x-pdf", "application/octet-stream"}

logger = logging.getLogger("auto_invoice.web")
logger.setLevel(os.getenv("INVOICE_LOG_LEVEL", "INFO"))

executor = ThreadPoolExecutor(max_workers=2, thread_name_prefix="invoice-task")
stats_lock = threading.Lock()
app = FastAPI(title="Auto Invoice")
app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")
templates = Jinja2Templates(directory=TEMPLATES_DIR)


def _now_iso() -> str:
    return datetime.now(timezone.utc).isoformat()


def _safe_upload_filename(filename: Optional[str]) -> str:
    raw = Path(filename or "").name.strip()
    if not raw:
        raise HTTPException(status_code=400, detail="文件名不能为空")
    safe = re.sub(r"[\\/:*?\"<>|]+", "_", raw)
    safe = re.sub(r"\s+", " ", safe).strip(" .")
    if not safe:
        raise HTTPException(status_code=400, detail="文件名不能为空")
    if Path(safe).suffix.lower() != ".pdf":
        raise HTTPException(status_code=400, detail="只支持上传 PDF 文件")
    return safe


def _unique_path(directory: Path, filename: str) -> Path:
    candidate = directory / filename
    if not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    idx = 1
    while True:
        next_candidate = directory / f"{stem}_{idx}{suffix}"
        if not next_candidate.exists():
            return next_candidate
        idx += 1


def _validate_task_id(task_id: str) -> str:
    try:
        parsed = uuid.UUID(task_id)
    except ValueError as exc:
        raise HTTPException(status_code=404, detail="任务不存在") from exc
    if str(parsed) != task_id:
        raise HTTPException(status_code=404, detail="任务不存在")
    return task_id


def _task_root(task_id: str) -> Path:
    safe_task_id = _validate_task_id(task_id)
    root = (TASKS_DIR / safe_task_id).resolve()
    tasks_root = TASKS_DIR.resolve()
    if tasks_root not in root.parents:
        raise HTTPException(status_code=404, detail="任务不存在")
    if not root.exists():
        raise HTTPException(status_code=404, detail="任务不存在")
    return root


def _task_json_path(task_id: str) -> Path:
    return _task_root(task_id) / "task.json"


def _read_task(task_id: str) -> dict[str, Any]:
    path = _task_json_path(task_id)
    if not path.exists():
        raise HTTPException(status_code=404, detail="任务不存在")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise HTTPException(status_code=500, detail="任务状态文件损坏") from exc


def _write_task(root: Path, task_id: str, status: str, message: str, extra: Optional[dict[str, Any]] = None) -> None:
    path = root / "task.json"
    payload: dict[str, Any] = {}
    if path.exists():
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except json.JSONDecodeError:
            payload = {}

    now = _now_iso()
    payload.update(
        {
            "task_id": task_id,
            "status": status,
            "message": message,
            "updated_at": now,
        }
    )
    if extra:
        payload.update(extra)
    payload.setdefault("created_at", now)

    tmp_path = path.with_suffix(".json.tmp")
    tmp_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(path)


def _default_stats() -> dict[str, Any]:
    return {
        "usage_count": 0,
        "processed_invoice_count": 0,
        "successful_task_count": 0,
        "failed_task_count": 0,
        "last_updated_at": None,
    }


def _read_stats_unlocked() -> dict[str, Any]:
    stats = _default_stats()
    if STATS_PATH.exists():
        try:
            raw = json.loads(STATS_PATH.read_text(encoding="utf-8"))
            if isinstance(raw, dict):
                stats.update(raw)
        except json.JSONDecodeError:
            logger.exception("Stats file is corrupted: %s", STATS_PATH)

    for key in ("usage_count", "processed_invoice_count", "successful_task_count", "failed_task_count"):
        try:
            stats[key] = int(stats.get(key) or 0)
        except (TypeError, ValueError):
            stats[key] = 0
    return stats


def _write_stats_unlocked(stats: dict[str, Any]) -> None:
    STATS_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = STATS_PATH.with_suffix(".json.tmp")
    tmp_path.write_text(json.dumps(stats, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp_path.replace(STATS_PATH)


def _read_stats() -> dict[str, Any]:
    with stats_lock:
        return _read_stats_unlocked()


def _increment_stats(**increments: int) -> dict[str, Any]:
    with stats_lock:
        stats = _read_stats_unlocked()
        for key, amount in increments.items():
            stats[key] = int(stats.get(key) or 0) + int(amount)
        stats["last_updated_at"] = _now_iso()
        _write_stats_unlocked(stats)
        return stats


def _cleanup_old_tasks() -> None:
    if not TASKS_DIR.exists():
        return

    cutoff = datetime.now(timezone.utc).timestamp() - TASK_RETENTION_SECONDS
    for child in TASKS_DIR.iterdir():
        if not child.is_dir():
            continue
        status_path = child / "task.json"
        try:
            if status_path.exists():
                payload = json.loads(status_path.read_text(encoding="utf-8"))
                if payload.get("status") == "RUNNING":
                    continue
            if child.stat().st_mtime < cutoff:
                shutil.rmtree(child, ignore_errors=True)
        except OSError:
            logger.exception("Failed to clean old task directory: %s", child)


def _zip_directory(source_dir: Path, zip_path: Path) -> None:
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for path in sorted(source_dir.rglob("*")):
            if path.is_file():
                archive.write(path, path.relative_to(source_dir.parent))


def _run_invoice_task(task_id: str) -> None:
    root = TASKS_DIR / task_id
    input_dir = root / "input"
    result_dir = root / "output" / "整理结果"
    zip_path = root / "result.zip"

    try:
        _write_task(root, task_id, "RUNNING", "正在处理发票")
        process_invoice_batch(
            input_dir=str(input_dir),
            template=str(TEMPLATE_PATH),
            output_dir=str(result_dir),
            rename=True,
            recursive=False,
            summary_name="A物品清单.xlsx",
            no_summary=False,
        )

        if not result_dir.exists() or not any(result_dir.iterdir()):
            raise RuntimeError("未生成有效的整理结果")

        if zip_path.exists():
            zip_path.unlink()
        _zip_directory(result_dir, zip_path)
        if not zip_path.exists() or zip_path.stat().st_size == 0:
            raise RuntimeError("结果压缩包生成失败")

        task_payload = _read_task(task_id)
        if not task_payload.get("success_stats_recorded"):
            file_count = int(task_payload.get("file_count") or 0)
            _increment_stats(processed_invoice_count=file_count, successful_task_count=1)

        _write_task(root, task_id, "SUCCESS", "处理完成", {"success_stats_recorded": True})
    except Exception as exc:
        logger.exception("Invoice task failed: %s", task_id)
        try:
            task_payload = _read_task(task_id)
        except HTTPException:
            task_payload = {}
        if not task_payload.get("failure_stats_recorded"):
            _increment_stats(failed_task_count=1)
        _write_task(root, task_id, "FAILED", str(exc)[:500] or "处理失败", {"failure_stats_recorded": True})


async def _save_uploads(files: list[UploadFile], input_dir: Path) -> dict[str, int]:
    total_size = 0
    saved_count = 0

    for upload in files:
        filename = _safe_upload_filename(upload.filename)
        content_type = (upload.content_type or "").lower()
        if content_type and content_type not in ALLOWED_MIME_TYPES:
            raise HTTPException(status_code=400, detail=f"{filename} 不是有效的 PDF 文件")

        target = _unique_path(input_dir, filename)
        file_size = 0
        first_chunk = await upload.read(CHUNK_SIZE)
        if not first_chunk:
            raise HTTPException(status_code=400, detail=f"{filename} 是空文件")
        if not first_chunk.lstrip().startswith(b"%PDF-"):
            raise HTTPException(status_code=400, detail=f"{filename} 不是有效的 PDF 文件")

        with target.open("wb") as out:
            chunk = first_chunk
            while chunk:
                file_size += len(chunk)
                total_size += len(chunk)
                if file_size > MAX_FILE_SIZE:
                    raise HTTPException(status_code=400, detail=f"{filename} 超过单文件大小限制")
                if total_size > MAX_UPLOAD_SIZE:
                    raise HTTPException(status_code=400, detail="本次上传超过总大小限制")
                out.write(chunk)
                chunk = await upload.read(CHUNK_SIZE)
        saved_count += 1

    return {"file_count": saved_count, "total_uploaded_size": total_size}


@app.get("/")
async def index(request: Request):
    return templates.TemplateResponse(request, "index.html")


@app.post("/api/tasks/upload")
async def upload_task(files: Optional[list[UploadFile]] = File(default=None)):
    if not files:
        raise HTTPException(status_code=400, detail="请选择至少一个 PDF 文件")
    if not TEMPLATE_PATH.exists():
        raise HTTPException(status_code=500, detail="模板文件不存在")

    _cleanup_old_tasks()

    task_id = str(uuid.uuid4())
    root = TASKS_DIR / task_id
    input_dir = root / "input"
    input_dir.mkdir(parents=True, exist_ok=True)

    try:
        upload_stats = await _save_uploads(files, input_dir)
        _write_task(root, task_id, "PENDING", "任务已创建", upload_stats)
        _increment_stats(usage_count=1)
    except HTTPException:
        shutil.rmtree(root, ignore_errors=True)
        raise
    except Exception as exc:
        shutil.rmtree(root, ignore_errors=True)
        logger.exception("Upload failed")
        raise HTTPException(status_code=500, detail="上传失败") from exc
    finally:
        for upload in files:
            await upload.close()

    executor.submit(_run_invoice_task, task_id)
    return {"task_id": task_id, "status": "PENDING"}


@app.get("/api/stats")
async def app_stats():
    return _read_stats()


@app.get("/api/tasks/{task_id}/status")
async def task_status(task_id: str):
    payload = _read_task(task_id)
    status = payload.get("status")
    download_url = None
    if status == "SUCCESS" and (_task_root(task_id) / "result.zip").exists():
        download_url = f"/api/tasks/{task_id}/download"
    return {
        "task_id": payload.get("task_id", task_id),
        "status": status,
        "message": payload.get("message", ""),
        "download_url": download_url,
    }


@app.get("/api/tasks/{task_id}/download")
async def download_task(task_id: str):
    root = _task_root(task_id)
    payload = _read_task(task_id)
    status = payload.get("status")

    if status in {"PENDING", "RUNNING"}:
        raise HTTPException(status_code=409, detail="任务尚未处理完成")
    if status == "FAILED":
        raise HTTPException(status_code=400, detail=payload.get("message") or "任务处理失败")
    if status != "SUCCESS":
        raise HTTPException(status_code=400, detail="任务状态异常")

    zip_path = root / "result.zip"
    if not zip_path.exists():
        raise HTTPException(status_code=404, detail="结果文件不存在")

    return FileResponse(zip_path, media_type="application/zip", filename="发票整理结果.zip")
