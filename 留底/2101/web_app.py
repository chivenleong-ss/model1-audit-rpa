# -*- coding: utf-8 -*-
"""
web_app.py — SAP 审计 · 效益审核自动填报平台 (优雅外部依赖版)
"""

import os, sys, uuid, threading, traceback
from datetime import datetime
from flask import Flask, request, jsonify, send_file, render_template_string, abort

# 🌟 核心路径防错补丁：让 exe 永远找得到它身边的配件 🌟
if getattr(sys, 'frozen', False):
    _DIR = os.path.dirname(sys.executable)
else:
    _DIR = os.path.dirname(os.path.abspath(__file__))

if _DIR not in sys.path:
    sys.path.insert(0, _DIR)

TEMPLATE_XL = os.path.join(_DIR, "项目效益审核表.xlsx")
HTML_PATH   = os.path.join(_DIR, "_page.html")  # 锁定旁边的 HTML 文件
WORK_BASE   = os.path.join(_DIR, "_jobs")
os.makedirs(WORK_BASE, exist_ok=True)

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 300 * 1024 * 1024

_JOBS: dict = {}
_LOCK = threading.Lock()

@app.route("/")
def index():
    # 🌟 动态热更新读取：每次刷新网页都会重新读取 _page.html
    if not os.path.exists(HTML_PATH):
        # 如果文件弄丢了，会在浏览器里给出友好的中文提示，而不是直接崩溃
        return f"""
        <h2 style='color:red; font-family:sans-serif; text-align:center; margin-top:50px;'>
            ❌ 找不到网页界面文件！
        </h2>
        <p style='text-align:center; font-family:sans-serif;'>
            请确保 <b>_page.html</b> 和 <b>web_app.exe</b> 放在同一个文件夹内。<br>
            <span style='color:gray; font-size:12px;'>当前查找路径: {HTML_PATH}</span>
        </p>
        """
    with open(HTML_PATH, encoding="utf-8") as f:
        return render_template_string(f.read())
    

@app.route("/api/submit", methods=["POST"])
def submit():
    voucher = request.files.get("voucher")
    detail  = request.files.get("detail")
    invoice = request.files.get("invoice")
    if not voucher or not detail:
        return jsonify({"error": "需要上传两个文件"}), 400
    job_id  = uuid.uuid4().hex[:14]
    job_dir = os.path.join(WORK_BASE, job_id)
    os.makedirs(job_dir)
    v_path = os.path.join(job_dir, voucher.filename)
    d_path = os.path.join(job_dir, detail.filename)
    voucher.save(v_path); detail.save(d_path)
    i_path = ""
    if invoice and invoice.filename:
        i_path = os.path.join(job_dir, invoice.filename)
        invoice.save(i_path)
    with _LOCK:
        _JOBS[job_id] = {
            "status": "running", "logs": [], "progress": 5,
            "error_summary": "", "error_detail": "",
            "delivery_path": "", "merged_path": "",
            "start_time": datetime.now().timestamp(), "elapsed": 0,
        }
    threading.Thread(target=_pipeline, args=(job_id, job_dir, v_path, d_path, i_path), daemon=True).start()
    return jsonify({"job_id": job_id})

@app.route("/api/status/<job_id>")
def status(job_id):
    job = _JOBS.get(job_id)
    if not job: return jsonify({"error": "任务不存在"}), 404
    with _LOCK:
        return jsonify({
            "status": job["status"], "progress": job["progress"],
            "all_logs": list(job["logs"]),
            "error_summary": job["error_summary"],
            "error_detail":  job["error_detail"],
            "elapsed": round(job["elapsed"], 1),
        })

@app.route("/api/dl/<job_id>/<which>")
def download(job_id, which):
    job = _JOBS.get(job_id)
    if not job: abort(404)
    path = job["delivery_path"] if which == "delivery" else job["merged_path"]
    if not path or not os.path.exists(path): abort(404)
    return send_file(path, as_attachment=True, download_name=os.path.basename(path))

# ▼▼▼ 新增：允许浏览器读取同一文件夹下的背景图片 ▼▼▼
@app.route("/bg.jpg")
def serve_bg():
    bg_path = os.path.join(_DIR, "bg.jpg")
    if os.path.exists(bg_path):
        return send_file(bg_path)
    return "", 404
# ▲▲▲ 新增结束 ▲▲▲

class _Log:
    def __init__(self, job_id):
        self.job_id = job_id
        self._steps = iter(range(20, 95, 3))
    def _push(self, msg):
        with _LOCK:
            job = _JOBS.get(self.job_id)
            if not job: return
            job["logs"].append(msg)
            try: job["progress"] = next(self._steps)
            except StopIteration: pass
    def info(self, msg, *a, **kw):      self._push(msg if not a else msg % a)
    def warning(self, msg, *a, **kw):   self._push("⚠️ " + (msg if not a else msg % a))
    def error(self, msg, *a, **kw):     self._push("❌ " + (msg if not a else msg % a))
    def exception(self, msg, *a, **kw): self.error(msg, *a, **kw)

def _pipeline(job_id, job_dir, v_path, d_path, i_path: str = ""):
    log = _Log(job_id)
    try:
        from sap_audit_core   import SAPAuditModel
        from benefit_reporter import BenefitReporter
        from excel_beautifier import ExcelBeautifier
        out_dir = os.path.join(job_dir, "output")
        log.info("🚀 启动数据合并与审计校验...")
        audit = SAPAuditModel(v_path, d_path, output_base_dir=out_dir)
        if not audit.execute_audit():
            raise RuntimeError("数据结构异常，审计模块返回失败")
        if not os.path.exists(TEMPLATE_XL):
            raise FileNotFoundError(f"模板文件不存在：{TEMPLATE_XL}")
        log.info("📝 启动效益审核表自动填报...")
        reporter = BenefitReporter(audit.result_dir, log)
        reporter.template_path = TEMPLATE_XL
        if i_path and os.path.exists(i_path):
            reporter.invoice_ledger_path = i_path
        if not reporter.execute_fill():
            raise RuntimeError("效益审核表填报失败")
        log.info("💅 启动格式美化引擎...")
        beautifier = ExcelBeautifier(audit.result_dir, log)
        if not beautifier.execute_beautify():
            raise RuntimeError("效益审核表美化失败")
        delivery = os.path.join(audit.result_dir, "最终完美交付版_效益审核表.xlsx")
        merged   = os.path.join(audit.result_dir, "4_合并表.xlsx")
        if not os.path.exists(delivery):
            raise FileNotFoundError("交付版文件未生成")
        elapsed = datetime.now().timestamp() - _JOBS[job_id]["start_time"]
        with _LOCK:
            _JOBS[job_id].update({"status":"done","progress":100,
                "delivery_path":delivery,"merged_path":merged,"elapsed":elapsed})
        log.info("🎉 全部完成！文件已就绪。")
    except Exception as exc:
        detail  = traceback.format_exc()
        elapsed = datetime.now().timestamp() - _JOBS[job_id]["start_time"]
        with _LOCK:
            _JOBS[job_id].update({"status":"error","progress":100,
                "error_summary":str(exc)[:200],"error_detail":detail,"elapsed":elapsed})
        log.error("流程异常：%s", str(exc))

if __name__ == "__main__":
    print("\n  ┌──────────────────────────────────────────────┐")
    print("  │   效益审核自动填报平台  已启动                │")
    print("  │   访问地址：http://localhost:5000             │")
    print("  └──────────────────────────────────────────────┘\n")
    app.run(host="0.0.0.0", port=5000, debug=False, threaded=True)
