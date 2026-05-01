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
import json
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
    pc_name: str = ""  # <--- 新增：记忆利润中心名称描述
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
    if not args: return str(msg)
    try: return str(msg) % args
    except Exception: return " ".join([str(msg), *[str(arg) for arg in args]]).strip()


def _sanitize_filename(filename: str | None, fallback: str) -> str:
    name = os.path.basename((filename or "").strip())
    if not name: return fallback
    name = re.sub(r"[\x00-\x1f]", "_", name)
    return "".join("_" if ch in '<>:"/\\|?*' else ch for ch in name).strip(" .") or fallback


def _uploaded(file_storage) -> bool:
    return bool(file_storage and getattr(file_storage, "filename", "").strip())


def _elapsed_seconds(job: JobRecord) -> float:
    if job.elapsed > 0: return job.elapsed
    return max(0.0, datetime.now().timestamp() - job.start_time)


def _set_job_progress(job_id: str, progress: int) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if job: job.progress = max(job.progress, min(progress, 100))


def _append_job_log(job_id: str, message: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if job: job.logs.append(message)


def _start_job(job_id: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if job:
            job.status = "running"
            job.progress = max(job.progress, 5)
            job.start_time = datetime.now().timestamp()


def _finish_job_success(job_id: str, delivery_path: str, merged_path: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if job:
            job.status = "done"
            job.progress = 100
            job.delivery_path = delivery_path
            job.merged_path = merged_path
            job.elapsed = _elapsed_seconds(job)


def _finish_job_error(job_id: str, error_summary: str, error_detail: str) -> None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if job:
            job.status = "error"
            job.progress = 100
            job.error_summary = error_summary[:200]
            job.error_detail = error_detail
            job.elapsed = _elapsed_seconds(job)


def _job_payload(job_id: str) -> dict | None:
    with _LOCK:
        job = _JOBS.get(job_id)
        if not job: return None
        return {
            "job_id": job_id,
            "status": job.status,
            "progress": job.progress,
            "error_summary": job.error_summary,
            "error_detail": job.error_detail,
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
    file_storage.save(str(target))
    return str(target)


def _copy_to_job_dir(src: str, job_dir: Path) -> str:
    source = Path(src)
    target = job_dir / _sanitize_filename(source.name, source.name)
    if target.exists():
        target = job_dir / f"{target.stem}_{uuid.uuid4().hex[:6]}{target.suffix}"
    shutil.copy2(str(source), str(target))
    return str(target)


def _read_page_html() -> str:
    if HTML_PATH.exists():
        return HTML_PATH.read_text(encoding="utf-8")
    return "<h2 style='color:red;'>找不到网页界面文件 _page.html</h2>"


class JobLogger:
    def __init__(self, job_id: str): self.job_id = job_id
    def set_progress(self, progress: int) -> None: _set_job_progress(self.job_id, progress)
    def _emit(self, prefix: str, msg, *args) -> None: _append_job_log(self.job_id, prefix + _format_message(msg, *args))
    def info(self, msg, *args, **kwargs) -> None: self._emit("", msg, *args)
    def warning(self, msg, *args, **kwargs) -> None: self._emit("⚠️ ", msg, *args)
    def error(self, msg, *args, **kwargs) -> None: self._emit("❌ ", msg, *args)
    def exception(self, msg, *args, **kwargs) -> None: self.error(msg, *args)


@app.route("/")
def index():
    return render_template_string(_read_page_html())


@app.route("/api/submit_batch", methods=["POST"])
def submit_batch():
    action = (request.form.get("action") or "preview").strip().lower()
    if action == "start":
        return _start_batch_from_preview()

    upload_files = [item for item in request.files.getlist("files") if _uploaded(item)]
    if not upload_files: return jsonify({"error": "请上传文件"}), 400

    batch_id = uuid.uuid4().hex[:14]
    batch_dir = WORK_BASE / f"batch_{batch_id}"
    incoming_dir = batch_dir / "incoming"
    try:
        incoming_dir.mkdir(parents=True, exist_ok=False)
        saved_paths = [_save_upload_file(fs, incoming_dir, f"upload_{i}.xlsx") for i, fs in enumerate(upload_files, 1)]
        groups, rejected = build_profit_center_groups(saved_paths)
    except Exception as exc:
        return jsonify({"error": f"批量处理失败：{exc}"}), 400

    rejected_payload = [{"filename": r.filename, "profit_center": r.profit_center, "error": r.error} for r in rejected]

    with _LOCK:
        _BATCHES[batch_id] = BatchRecord(batch_id=batch_id, batch_dir=str(batch_dir), groups=groups, rejected=rejected_payload)

    return jsonify({"batch_id": batch_id, "groups": [_group_payload(g) for g in groups], "rejected": rejected_payload})


def _start_batch_from_preview():
    batch_id = (request.form.get("batch_id") or "").strip()
    with _LOCK:
        batch = _BATCHES.get(batch_id)
        if not batch: return jsonify({"error": "批次失效"}), 404
        if batch.started: return jsonify({"error": "已开始处理"}), 400
        batch.started = True

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
                # 【传递名称】
                _JOBS[job_id] = JobRecord(status="error", progress=100, profit_center=group.profit_center, pc_name=group.pc_name, error_summary=str(exc))
            job_items.append((job_id, "", "", "", ""))
            continue

        with _LOCK: 
            # 【传递名称】
            _JOBS[job_id] = JobRecord(status="queued", progress=0, profit_center=group.profit_center, pc_name=group.pc_name)
        job_items.append((job_id, str(job_dir), voucher_path, detail_path, invoice_path))

    threading.Thread(target=_run_batch_jobs, args=(job_items,), daemon=True).start()

    with _LOCK:
        if batch: batch.job_ids = [item[0] for item in job_items]
        jobs_payload = [{"job_id": i[0], "profit_center": _JOBS[i[0]].profit_center} for i in job_items if i[0] in _JOBS]

    return jsonify({"batch_id": batch_id, "job_ids": [i[0] for i in job_items], "jobs": jobs_payload})


@app.route("/api/status/<job_id>")
def status(job_id: str):
    payload = _job_payload(job_id)
    if not payload: return jsonify({"error": "任务不存在"}), 404
    return jsonify(payload)


@app.route("/api/download_selected")
def download_selected():
    file_type = request.args.get("type")
    ids_str = request.args.get("ids", "")
    if not ids_str: return "<h3 style='color:red;text-align:center;'>下载失败：没有勾选任何项目！</h3>", 400
        
    project_ids = [pid.strip() for pid in ids_str.split(",") if pid.strip()]
    matched_files = []
    seen_pc = set()

    with _LOCK:
        for job in reversed(list(_JOBS.values())):
            if job.profit_center in project_ids and job.status == "done" and job.profit_center not in seen_pc:
                if file_type == "benefit" and job.delivery_path and os.path.exists(job.delivery_path):
                    matched_files.append((job.profit_center, job.delivery_path))
                    seen_pc.add(job.profit_center)
                elif file_type == "merged" and job.merged_path and os.path.exists(job.merged_path):
                    matched_files.append((job.profit_center, job.merged_path))
                    seen_pc.add(job.profit_center)

    if not matched_files:
        return "<h3 style='color:red;text-align:center;margin-top:50px;'>下载失败：所选项目尚未生成成功，或底层文件已被删除！</h3>", 404

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for label, filepath in matched_files:
            filename = f"【{label}】_{os.path.basename(filepath)}"
            zf.write(filepath, arcname=filename)
    buf.seek(0)
    label_zh = "已填列效益表" if file_type == "benefit" else "合并凭证表"
    timestamp = datetime.now().strftime("%H%M%S")
    return send_file(buf, mimetype="application/zip", as_attachment=True, download_name=f"批量下载_{label_zh}_{timestamp}.zip")


@app.route("/api/handle_action", methods=["POST"])
def api_handle_action():
    data = request.get_json()
    action = data.get("action")
    if action == "push_audit": return jsonify({"status": "success", "message": "推送成功！"})
    return jsonify({"status": "success", "message": "未知动作"})


@app.route("/api/logic_audit_upload", methods=["POST"])
def logic_audit_upload():
    upload_files = [item for item in request.files.getlist("files") if _uploaded(item)]
    try: pushed_projects = json.loads(request.form.get("pushed_projects", "[]"))
    except: pushed_projects = []

    if not upload_files and not pushed_projects:
        return jsonify({"error": "没有收到任何上传的文件或推送的项目"}), 400

    from logic_auditor import LogicAuditor
    from batch_grouping import extract_profit_center_info
    
    auditor = LogicAuditor()
    audit_results = []
    temp_dir = WORK_BASE / "logic_temp"
    temp_dir.mkdir(parents=True, exist_ok=True)
    tasks = {}

    # 1. 抓取被推送过来的双表（同时提取记录好的 pc_name）
    for pc_code in pushed_projects:
        tasks[pc_code] = {'benefit': None, 'merged': None, 'pc_code': pc_code, 'pc_name': pc_code}
        for job in reversed(list(_JOBS.values())):
            if job.profit_center == pc_code and job.status == "done":
                if job.merged_path and os.path.exists(job.merged_path): tasks[pc_code]['merged'] = job.merged_path
                if job.delivery_path and os.path.exists(job.delivery_path): tasks[pc_code]['benefit'] = job.delivery_path
                if job.pc_name: tasks[pc_code]['pc_name'] = job.pc_name  # <--- 提取项目名
                break

    # 2. 抓取手工上传的双表
    for f in upload_files:
        saved_path = _save_upload_file(f, temp_dir, f.filename)
        try: pc_code, pc_name = extract_profit_center_info(saved_path)
        except: pc_code, pc_name = f.filename, f.filename
            
        if pc_code not in tasks:
            tasks[pc_code] = {'benefit': None, 'merged': None, 'pc_code': pc_code, 'pc_name': pc_name or pc_code}
        else:
            if pc_name and pc_name != pc_code: tasks[pc_code]['pc_name'] = pc_name
            
        if "合并" in f.filename: tasks[pc_code]['merged'] = saved_path
        else: tasks[pc_code]['benefit'] = saved_path

    # 3. 补漏机制
    for pc_code, paths in tasks.items():
        if paths['merged'] is None:
            for job in reversed(list(_JOBS.values())):
                if job.profit_center == pc_code and job.status == "done":
                    if job.merged_path and os.path.exists(job.merged_path): paths['merged'] = job.merged_path
                    break

    # 4. 执行核心审计
    for pc_code, paths in tasks.items():
        try:
            res = auditor.run_audit(benefit_path=paths['benefit'], merged_path=paths['merged'], project_name=paths['pc_name'])
            res['pc_code'] = paths['pc_code'] # 注入编码
            audit_results.append(res)
        except Exception as e:
            audit_results.append({
                "pc_code": paths['pc_code'],
                "project_name": paths['pc_name'],
                "is_pass": False,
                "errors": [{"type": "系统错误", "row": "-", "desc": f"解析失败: {str(e)}"}],
                "details": [{"rule": "核心异常", "status": "fail", "desc": str(e)}]
            })

    return jsonify({"status": "success", "data": audit_results})


def _run_batch_jobs(job_items: list[tuple[str, str, str, str, str]]) -> None:
    for job_id, job_dir, voucher_path, detail_path, invoice_path in job_items:
        with _LOCK:
            current = _JOBS.get(job_id)
            if current and current.status == "error": continue
        
        _start_job(job_id)
        if not voucher_path or not detail_path:
            _finish_job_error(job_id, "文件缺失", "缺少凭证主表或辅助明细账")
            continue
        _pipeline(job_id, job_dir, voucher_path, detail_path, invoice_path)


def _pipeline(job_id: str, job_dir: str, voucher_path: str, detail_path: str, invoice_path: str = ""):
    log = JobLogger(job_id)
    try:
        from audit_pipeline import AuditPipeline
        output_dir = Path(job_dir) / "output"
        pipeline = AuditPipeline(logger=log, template_path=str(TEMPLATE_XL))
        result = pipeline.run(
            voucher_path=voucher_path, detail_path=detail_path,
            output_base_dir=str(output_dir), invoice_path=invoice_path,
        )
        _finish_job_success(job_id, result.delivery_path, result.merged_path)
    except Exception as exc:
        _finish_job_error(job_id, str(exc), traceback.format_exc())


if __name__ == "__main__":
    print("\n  " + "─" * 58)
    print("  效益审核自动填报平台 已启动")
    print("  访问地址：http://localhost:5000")
    print("  " + "─" * 58 + "\n")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)