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

import io
import os
import re
import shutil
import sys
import threading
import traceback
import uuid
import zipfile
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from flask import Flask, abort, jsonify, render_template_string, request, send_file

from batch_grouping import ProfitCenterGroup, build_profit_center_groups


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
    profit_center: str = ""
    files: dict[str, list[str]] = field(default_factory=dict)
    start_time: float = field(default_factory=lambda: datetime.now().timestamp())
    elapsed: float = 0.0


@dataclass
class BatchRecord:
    batch_id: str
    batch_dir: str
    groups: list[ProfitCenterGroup]
    rejected: list[dict]
    started: bool = False
    job_ids: list[str] = field(default_factory=list)


_JOBS: dict[str, JobRecord] = {}
_BATCHES: dict[str, BatchRecord] = {}
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


def _start_job(job_id: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job:
            return
        job.status = "running"
        job.progress = max(job.progress, 5)
        job.start_time = datetime.now().timestamp()


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
            "profit_center": job.profit_center,
            "files": dict(job.files),
        }


def _group_payload(group: ProfitCenterGroup) -> dict:
    return group.to_payload()


def _save_upload_file(file_storage, target_dir: Path, fallback: str) -> str:
    filename = _sanitize_filename(getattr(file_storage, "filename", ""), fallback)
    target = target_dir / filename
    if target.exists():
        stem = target.stem
        suffix = target.suffix
        target = target_dir / f"{stem}_{uuid.uuid4().hex[:6]}{suffix}"
    try:
        file_storage.save(str(target))
    except Exception as exc:
        raise RuntimeError(f"保存上传文件失败：{filename}，原因：{exc}") from exc
    if not target.exists():
        raise RuntimeError(f"上传文件保存失败：{filename}")
    return str(target)


def _copy_to_job_dir(src: str, job_dir: Path) -> str:
    source = Path(src)
    target = job_dir / _sanitize_filename(source.name, source.name)
    if target.exists():
        target = job_dir / f"{target.stem}_{uuid.uuid4().hex[:6]}{target.suffix}"
    try:
        shutil.copy2(str(source), str(target))
    except Exception as exc:
        raise RuntimeError(f"复制分组文件失败：{source.name}，原因：{exc}") from exc
    return str(target)


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
    try:
        job_dir.mkdir(parents=True, exist_ok=False)
        voucher_path = _save_upload_file(voucher, job_dir, "voucher.xlsx")
        detail_path = _save_upload_file(detail, job_dir, "detail.xlsx")
        invoice_path = ""
        if _uploaded(invoice):
            invoice_path = _save_upload_file(invoice, job_dir, "invoice.xlsx")
    except Exception as exc:
        return jsonify({"error": str(exc)}), 400

    with _LOCK:
        _JOBS[job_id] = JobRecord(
            files={
                "voucher": [os.path.basename(voucher_path)],
                "detail": [os.path.basename(detail_path)],
                "invoice": [os.path.basename(invoice_path)] if invoice_path else [],
            }
        )

    thread = threading.Thread(
        target=_pipeline,
        args=(job_id, str(job_dir), voucher_path, detail_path, invoice_path),
        daemon=True,
        name=f"web-job-{job_id}",
    )
    thread.start()
    return jsonify({"job_id": job_id})


@app.route("/api/submit_batch", methods=["POST"])
def submit_batch():
    action = (request.form.get("action") or "preview").strip().lower()
    if action == "start":
        return _start_batch_from_preview()

    upload_files = [item for item in request.files.getlist("files") if _uploaded(item)]
    if not upload_files:
        return jsonify({"error": "请至少上传一个 Excel 或 CSV 文件"}), 400

    batch_id = uuid.uuid4().hex[:14]
    batch_dir = WORK_BASE / f"batch_{batch_id}"
    incoming_dir = batch_dir / "incoming"
    try:
        incoming_dir.mkdir(parents=True, exist_ok=False)
        saved_paths = [
            _save_upload_file(file_storage, incoming_dir, f"upload_{index}.xlsx")
            for index, file_storage in enumerate(upload_files, start=1)
        ]
        groups, rejected = build_profit_center_groups(saved_paths)
    except Exception as exc:
        return jsonify({"error": f"批量上传处理失败：{exc}"}), 400

    rejected_payload = [
        {
            "filename": item.filename,
            "profit_center": item.profit_center,
            "file_type": item.file_type,
            "error": item.error or "文件未能参与分组",
        }
        for item in rejected
    ]

    with _LOCK:
        _BATCHES[batch_id] = BatchRecord(
            batch_id=batch_id,
            batch_dir=str(batch_dir),
            groups=groups,
            rejected=rejected_payload,
        )

    return jsonify(
        {
            "batch_id": batch_id,
            "groups": [_group_payload(group) for group in groups],
            "rejected": rejected_payload,
        }
    )


def _start_batch_from_preview():
    batch_id = (request.form.get("batch_id") or "").strip()
    if not batch_id:
        return jsonify({"error": "缺少批次编号，无法开始处理"}), 400

    with _LOCK:
        batch = _BATCHES.get(batch_id)
        if not batch:
            return jsonify({"error": "批次不存在或已失效，请重新上传"}), 404
        if batch.started:
            return jsonify({"error": "该批次已开始处理，请勿重复提交"}), 400
        batch.started = True
        batch.job_ids = []

    job_items = []
    for group in batch.groups:
        job_id = uuid.uuid4().hex[:14]
        job_dir = WORK_BASE / job_id
        try:
            job_dir.mkdir(parents=True, exist_ok=False)
            voucher_path = _copy_to_job_dir(group.voucher_files[0].path, job_dir) if group.voucher_files else ""
            detail_path = _copy_to_job_dir(group.detail_files[0].path, job_dir) if group.detail_files else ""
            invoice_path = _copy_to_job_dir(group.invoice_files[0].path, job_dir) if group.invoice_files else ""
        except Exception as exc:
            with _LOCK:
                _JOBS[job_id] = JobRecord(
                    status="error",
                    progress=100,
                    profit_center=group.profit_center,
                    error_summary=str(exc),
                    error_detail=traceback.format_exc(),
                    files=group.to_payload(),
                )
            job_items.append((job_id, "", "", "", ""))
            continue

        with _LOCK:
            _JOBS[job_id] = JobRecord(
                status="queued",
                progress=0,
                profit_center=group.profit_center,
                files={
                    "voucher": [item.filename for item in group.voucher_files],
                    "detail": [item.filename for item in group.detail_files],
                    "invoice": [item.filename for item in group.invoice_files],
                    "other": [item.filename for item in group.other_files],
                },
            )
        job_items.append((job_id, str(job_dir), voucher_path, detail_path, invoice_path))

    thread = threading.Thread(
        target=_run_batch_jobs,
        args=(job_items,),
        daemon=True,
        name=f"web-batch-{batch_id}",
    )
    thread.start()

    job_ids = [item[0] for item in job_items]
    with _LOCK:
        batch = _BATCHES.get(batch_id)
        if batch:
            batch.job_ids = job_ids
        jobs_payload = [
            {"job_id": item[0], "profit_center": _JOBS[item[0]].profit_center}
            for item in job_items
            if item[0] in _JOBS
        ]

    return jsonify(
        {
            "batch_id": batch_id,
            "job_ids": job_ids,
            "jobs": jobs_payload,
        }
    )


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


@app.route("/api/batch_download/<batch_id>/<file_type>")
def batch_download(batch_id: str, file_type: str):
    if file_type not in ("delivery", "merged"):
        return jsonify({"error": "file_type 必须为 delivery 或 merged"}), 400

    with _LOCK:
        batch = _BATCHES.get(batch_id)
        if not batch:
            return jsonify({"error": "批次不存在或已失效"}), 404
        if not batch.started or not batch.job_ids:
            return jsonify({"error": "批次尚未开始处理"}), 400

    # 收集批次内所有已完成任务的路径
    matched_files = []
    with _LOCK:
        for job_id in batch.job_ids:
            job = _JOBS.get(job_id)
            if not job or job.status != "done":
                continue
            if file_type == "delivery" and job.delivery_path and os.path.exists(job.delivery_path):
                matched_files.append((job.profit_center or job_id, job.delivery_path))
            elif file_type == "merged" and job.merged_path and os.path.exists(job.merged_path):
                matched_files.append((job.profit_center or job_id, job.merged_path))

    if not matched_files:
        label = "交付版" if file_type == "delivery" else "合并表"
        return jsonify({"error": f"批次中尚无已完成的 {label} 文件"}), 400

    # 打包为 ZIP
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, filepath in matched_files:
            filename = f"{label}_{os.path.basename(filepath)}"
            zf.write(filepath, arcname=filename)
    buf.seek(0)
    label_zh = "交付版" if file_type == "delivery" else "合并表"
    return send_file(
        buf,
        mimetype="application/zip",
        as_attachment=True,
        download_name=f"batch_{batch_id}_{label_zh}.zip",
    )


@app.route("/bg.jpg")
def serve_bg():
    if BG_PATH.exists():
        return send_file(str(BG_PATH))
    return "", 404


def _run_batch_jobs(job_items: list[tuple[str, str, str, str, str]]) -> None:
    for job_id, job_dir, voucher_path, detail_path, invoice_path in job_items:
        with _LOCK:
            current = _JOBS.get(job_id)
            if current and current.status == "error":
                continue
            profit_center = current.profit_center if current else ""
            files = dict(current.files) if current else {}

        _start_job(job_id)
        _append_job_log(job_id, f"开始处理利润中心：{profit_center or '未识别'}")

        if not voucher_path or not detail_path:
            missing = []
            if not voucher_path:
                missing.append("凭证主表")
            if not detail_path:
                missing.append("辅助明细账")
            _finish_job_error(
                job_id,
                "、".join(missing) + "缺失",
                f"利润中心 {profit_center or '未识别'} 分组不完整，缺少：" + "、".join(missing),
            )
            continue

        if len(files.get("voucher", [])) > 1:
            _append_job_log(job_id, "⚠️ 该利润中心存在多个凭证主表，已使用列表中的第一个文件")
        if len(files.get("detail", [])) > 1:
            _append_job_log(job_id, "⚠️ 该利润中心存在多个辅助明细账，已使用列表中的第一个文件")
        if len(files.get("invoice", [])) > 1:
            _append_job_log(job_id, "⚠️ 该利润中心存在多个发票台账，已使用列表中的第一个文件")

        _pipeline(job_id, job_dir, voucher_path, detail_path, invoice_path)


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
