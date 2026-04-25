"""
SAP 审计 RPA 机器人（auto_runner.py）v2 + 终极自动填报版
"""

import os
import sys
import glob
import time
import shutil
import logging
from datetime import datetime

# ── 确保任意工作目录下都能找到自定义模块 ──────────
_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
if _THIS_DIR not in sys.path:
    sys.path.insert(0, _THIS_DIR)

from sap_audit_core import SAPAuditModel, setup_logger   # noqa: E402
# ★ 引入最新开发的效益表填报机器人
from benefit_reporter import BenefitReporter             # noqa: E402
# ▼ 新增这一行 ▼
from excel_beautifier import ExcelBeautifier

# ── 核心目录定义 ────────────────────────────────────────────────
INPUT_DIR  = "1_上传原始数据"
OUTPUT_DIR = "2_输出审计底稿"
BACKUP_DIR = "3_已处理数据备份"
ERROR_DIR  = "4_错误文件隔离区"

for _d in [INPUT_DIR, OUTPUT_DIR, BACKUP_DIR, ERROR_DIR]:
    os.makedirs(_d, exist_ok=True)

log = setup_logger(log_dir=".")


# ── 辅助函数 ────────────────────────────────────────────────────

def _is_temp_file(filename: str) -> bool:
    return filename.startswith("~") or filename.startswith(".")


def _is_file_stable(path: str, wait_sec: float = 1.5, checks: int = 2) -> bool:
    """连续 checks 次采样文件大小不变，才认为写入已完成（防止读取正在复制的文件）"""
    try:
        sizes = []
        for _ in range(checks):
            sizes.append(os.path.getsize(path))
            time.sleep(wait_sec)
        return len(set(sizes)) == 1
    except OSError:
        return False


def _move_to_dir(files: list, dest_dir: str, label: str) -> None:
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for src in files:
        if src and os.path.exists(src):
            base = os.path.basename(src)
            name, ext = os.path.splitext(base)
            dst = os.path.join(dest_dir, f"{name}_{timestamp}{ext}")
            shutil.move(src, dst)
            log.info("   → %s 已移至 [%s]：%s", label, dest_dir, os.path.basename(dst))


def _scan_input_dir() -> tuple:
    """收集所有匹配文件到列表，防止单变量覆盖赋值。"""
    all_files = glob.glob(os.path.join(INPUT_DIR, "*.*"))
    voucher_files, detail_files = [], []
    for path in all_files:
        fname = os.path.basename(path)
        if _is_temp_file(fname):
            continue
        if any(k in fname for k in ["凭证", "综合查询", "主表"]):
            voucher_files.append(path)
        elif any(k in fname for k in ["明细", "辅助"]):
            detail_files.append(path)
    return voucher_files, detail_files


# ── 主循环 ──────────────────────────────────────────────────────

log.info("=" * 50)
log.info("         🤖 SAP 审计 RPA 机器人已启动 (稳健版)")
log.info("=" * 50)
log.info("👀 正在 24 小时监控文件夹：[%s]", INPUT_DIR)
log.info("💡 请将《综合查询》和《辅助明细》直接拖入该文件夹")
log.info("⚠️  按 Ctrl+C 可安全停止程序\n")

try:
    while True:
        voucher_files, detail_files = _scan_input_dir()

        if len(voucher_files) > 1:
            log.warning("⚠️  检测到 %d 个凭证主表文件，请保证只放 1 个。跳过本轮…", len(voucher_files))
            time.sleep(3)
            continue

        if len(detail_files) > 1:
            log.warning("⚠️  检测到 %d 个明细账表文件，请保证只放 1 个。跳过本轮…", len(detail_files))
            time.sleep(3)
            continue

        voucher_file = voucher_files[0] if voucher_files else None
        detail_file  = detail_files[0]  if detail_files  else None

        if voucher_file and detail_file:
            log.info("\n🎉 叮！检测到新数据上传，RPA 开始执行任务...")
            log.info("📄 识别到主表：%s", os.path.basename(voucher_file))
            log.info("📄 识别到明细：%s", os.path.basename(detail_file))

            for fpath in [voucher_file, detail_file]:
                if not _is_file_stable(fpath):
                    log.warning("⏳  文件 [%s] 仍在写入，等待稳定…", os.path.basename(fpath))

            try:
                # 1. 触发核心大脑：底稿生成与 ETL 数据清洗
                audit_job = SAPAuditModel(voucher_file, detail_file, output_base_dir=OUTPUT_DIR)
                success = audit_job.execute_audit()

                if success:
                    # 2. ★ 触发效益表自动填报引擎 ★
                    log.info("🚀 底稿生成完毕，准备执行 Excel 自动填报...")
                    reporter = BenefitReporter(audit_job.result_dir, log)
                    reporter.execute_fill()

                    # ▼ 新增：触发美化引擎 ▼
                    beautifier = ExcelBeautifier(audit_job.result_dir, log)
                    beautifier.execute_beautify()

                    # 3. 任务成功，安全转移原文件
                    _move_to_dir([voucher_file, detail_file], BACKUP_DIR, "原始文件")
                    log.info("✅  任务完美结束！底稿已生成，原文件已备份。继续监控中…\n")
                else:
                    log.error("❌  数据结构严重错误，任务中止。")
                    _move_to_dir([voucher_file, detail_file], ERROR_DIR, "问题文件")
                    log.error("   问题文件已移至 [%s]，请检查后重新上传。\n", ERROR_DIR)

            except Exception as exc:
                log.exception("❌  程序发生未知异常：%s", exc)
                _move_to_dir([voucher_file, detail_file], ERROR_DIR, "异常文件")
                log.error("   异常文件已移至 [%s]，机器人继续监控。\n", ERROR_DIR)

        time.sleep(3)

except KeyboardInterrupt:
    log.info("\n🛑  收到退出指令，RPA 机器人已安全停止。")