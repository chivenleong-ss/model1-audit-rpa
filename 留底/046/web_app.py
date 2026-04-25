# -*- coding: utf-8 -*-
"""
网页端入口。

职责只保留三类：
1. 接收上传并落盘；
2. 管理任务状态与页面轮询；
3. 串起审计、填报、美化三个后端模块。

页面样式与前端模板继续由 _page.html 负责。
"""

from __future__ import annotations

import os
import re
import sys
import threading
import traceback
import uuid
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, jsonify, render_template_string, request, send_file


def _app_dir() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = _app_dir()
if str(APP_DIR) not in sys.path:
    sys.path.insert(0, str(APP_DIR))

TEMPLATE_XL = APP_DIR / "项目效益审核表.xlsx"
HTML_PATH = APP_DIR / "_page.html"
BG_PATH = APP_DIR / "bg.jpg"
WORK_BASE = APP_DIR / "_jobs"
WORK_BASE.mkdir(parents=True, exist_ok=True)

MAX_UPLOAD_SIZE = 300 * 1024 * 1024
DELIVERY_CANDIDATES = (
    "最终完美交付版_效益审核表.xlsx",
    "交付版_效益审核表.xlsx",
    "自动填报完成_效益审核表.xlsx",
)
MERGED_CANDIDATES = ("4_合并表.xlsx",)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_SIZE


@dataclass
class JobRecord:
    status: str = "running"
    progress: int = 5
    logs: list[str] = field(default_factory=list)
    error_summary: str = ""
    error_detail: str = ""
    delivery_path: str = ""
    merged_path: str = ""
    start_time: float = field(default_factory=lambda: datetime.now().timestamp())
    elapsed: float = 0.0


_JOBS: dict[str, JobRecord] = {}
_LOCK = threading.Lock()


def _format_message(msg, *args) -> str:
    if not args:
        return str(msg)
    try:
        return str(msg) % args
    except Exception:
        return " ".join([str(msg), *[str(arg) for arg in args]]).strip()


def _sanitize_filename(filename: str | None, fallback: str) -> str:
    name = os.path.basename((filename or "").strip())
    if not name:
        return fallback
    name = re.sub(r"[\x00-\x1f]", "_", name)
    name = "".join("_" if ch in '<>:"/\\|?*' else ch for ch in name).strip(" .")
    return name or fallback


def _uploaded(file_storage) -> bool:
    return bool(file_storage and getattr(file_storage, "filename", "").strip())


def _elapsed_seconds(job: JobRecord) -> float:
    if job.elapsed > 0:
        return job.elapsed
    return max(0.0, datetime.now().timestamp() - job.start_time)


def _set_job_progress(job_id: str, progress: int) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return
        job.progress = max(job.progress, min(progress, 100))


def _append_job_log(job_id: str, message: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return
        job.logs.append(message)


def _finish_job_success(job_id: str, delivery_path: str, merged_path: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return
        job.status = "done"
        job.progress = 100
        job.delivery_path = delivery_path
        job.merged_path = merged_path
        job.elapsed = _elapsed_seconds(job)


def _finish_job_error(job_id: str, error_summary: str, error_detail: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return
        job.status = "error"
        job.progress = 100
        job.error_summary = error_summary[:200]
        job.error_detail = error_detail
        job.elapsed = _elapsed_seconds(job)


def _job_payload(job_id: str) -> dict | None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return None
        return {
            "job_id": job_id,
            "status": job.status,
            "progress": job.progress,
            "all_logs": list(job.logs),
            "error_summary": job.error_summary,
            "error_detail": job.error_detail,
            "elapsed": round(_elapsed_seconds(job), 1),
        }


def _find_output_file(result_dir: str | Path, filenames: tuple[str, ...]) -> str:
    base = Path(result_dir)
    for filename in filenames:
        path = base / filename
        if path.exists():
            return str(path)
    return ""


def _read_page_html() -> str:
    if HTML_PATH.exists():
        return HTML_PATH.read_text(encoding="utf-8")
    return f"""
    <h2 style='color:red; font-family:sans-serif; text-align:center; margin-top:50px;'>
        找不到网页界面文件
    </h2>
    <p style='text-align:center; font-family:sans-serif;'>
        请确认 <b>_page.html</b> 与 <b>web_app.exe</b> 放在同一目录。<br>
        <span style='color:gray; font-size:12px;'>当前查找路径：{HTML_PATH}</span>
    </p>
    """


class JobLogger:
    def __init__(self, job_id: str):
        self.job_id = job_id

    def set_progress(self, progress: int) -> None:
        _set_job_progress(self.job_id, progress)

    def _emit(self, prefix: str, msg, *args) -> None:
        _append_job_log(self.job_id, prefix + _format_message(msg, *args))

    def info(self, msg, *args, **kwargs) -> None:
        self._emit("", msg, *args)

    def warning(self, msg, *args, **kwargs) -> None:
        self._emit("⚠️ ", msg, *args)

    def error(self, msg, *args, **kwargs) -> None:
        self._emit("❌ ", msg, *args)

    def debug(self, msg, *args, **kwargs) -> None:
        self._emit("DEBUG ", msg, *args)

    def critical(self, msg, *args, **kwargs) -> None:
        self.error(msg, *args)

    def exception(self, msg, *args, **kwargs) -> None:
        self.error(msg, *args)


@app.route("/")
def index():
    return render_template_string(_read_page_html())


@app.route("/api/submit", methods=["POST"])
def submit():
    voucher = request.files.get("voucher")
    detail = request.files.get("detail")
    invoice = request.files.get("invoice")

    if not _uploaded(voucher) or not _uploaded(detail):
        return jsonify({"error": "需要上传凭证主表和辅助明细账两个文件"}), 400

    job_id = uuid.uuid4().hex[:14]
    job_dir = WORK_BASE / job_id
    job_dir.mkdir(parents=True, exist_ok=False)

    voucher_name = _sanitize_filename(voucher.filename, "voucher.xlsx")
    detail_name = _sanitize_filename(detail.filename, "detail.xlsx")
    invoice_name = _sanitize_filename(getattr(invoice, "filename", ""), "invoice.xlsx")

    voucher_path = job_dir / voucher_name
    detail_path = job_dir / detail_name
    invoice_path = ""

    voucher.save(str(voucher_path))
    detail.save(str(detail_path))
    if _uploaded(invoice):
        invoice_file = job_dir / invoice_name
        invoice.save(str(invoice_file))
        invoice_path = str(invoice_file)

    with _LOCK:
        _JOBS[job_id] = JobRecord()

    thread = threading.Thread(
        target=_pipeline,
        args=(job_id, str(job_dir), str(voucher_path), str(detail_path), invoice_path),
        daemon=True,
        name=f"web-job-{job_id}",
    )
    thread.start()
    return jsonify({"job_id": job_id})


@app.route("/api/status/<job_id>")
def status(job_id: str):
    payload = _job_payload(job_id)
    if payload is None:
        return jsonify({"error": "任务不存在"}), 404
    return jsonify(payload)


@app.route("/api/dl/<job_id>/<which>")
def download(job_id: str, which: str):
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            abort(404)
        if which == "delivery":
            path = job.delivery_path
        elif which == "merged":
            path = job.merged_path
        else:
            abort(404)

    if not path or not os.path.exists(path):
        abort(404)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))


@app.route("/bg.jpg")
def serve_bg():
    if BG_PATH.exists():
        return send_file(str(BG_PATH))
    return "", 404


def _pipeline(job_id: str, job_dir: str, voucher_path: str, detail_path: str, invoice_path: str = ""):
    log = JobLogger(job_id)
    try:
        from audit_pipeline import AuditPipeline

        output_dir = Path(job_dir) / "output"
        pipeline = AuditPipeline(logger=log, template_path=str(TEMPLATE_XL))
        result = pipeline.run(
            voucher_path=voucher_path,
            detail_path=detail_path,
            output_base_dir=str(output_dir),
            invoice_path=invoice_path,
        )
        _finish_job_success(job_id, result.delivery_path, result.merged_path)

    except Exception as exc:
        detail = traceback.format_exc()
        _finish_job_error(job_id, str(exc), detail)
        log.error("流程异常：%s", exc)


if __name__ == "__main__":
    print("\n  " + "─" * 58)
    print("  效益审核自动填报平台 已启动")
    print("  访问地址：http://localhost:5000")
    print("  " + "─" * 58 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
