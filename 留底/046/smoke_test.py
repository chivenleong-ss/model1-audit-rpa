# -*- coding: utf-8 -*-
"""
端到端冒烟测试脚本。

默认行为：
1. 从 `3_已处理数据备份` 自动挑选一对最近的“凭证 + 明细”样本；
2. 调用 AuditPipeline 跑完整主链；
3. 检查关键输出文件是否生成；
4. 以清晰的控制台结果返回 0/1。

用法：
  python smoke_test.py
  python smoke_test.py --voucher <路径> --detail <路径> [--invoice <路径>]
  python smoke_test.py --sample-dir <目录> --output-dir <目录>
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

from audit_pipeline import AuditPipeline
from sap_audit_core import setup_logger


ROOT = Path(__file__).resolve().parent
DEFAULT_SAMPLE_DIR = ROOT / "3_已处理数据备份"
DEFAULT_OUTPUT_DIR = ROOT / "_smoke_output"

VOUCHER_HINTS = ("凭证", "综合查询", "主表")
DETAIL_HINTS = ("明细", "辅助")
INVOICE_HINTS = ("发票", "收票", "台账", "台帳", "已认证")

EXPECTED_RESULT_FILES = (
    "4_合并表.xlsx",
    "5_中间汇总表_效益审核数据源.xlsx",
)
EXPECTED_DELIVERY_FILES = (
    "最终完美交付版_效益审核表.xlsx",
    "交付版_效益审核表.xlsx",
    "自动填报完成_效益审核表.xlsx",
)


def _is_match(name: str, hints: tuple[str, ...]) -> bool:
    return any(hint in name for hint in hints)


def _latest_files(sample_dir: Path, hints: tuple[str, ...]) -> list[Path]:
    files = [p for p in sample_dir.iterdir() if p.is_file() and _is_match(p.name, hints)]
    return sorted(files, key=lambda p: p.stat().st_mtime, reverse=True)


def auto_select_inputs(sample_dir: Path) -> tuple[Path, Path, Path | None]:
    voucher_files = _latest_files(sample_dir, VOUCHER_HINTS)
    detail_files = _latest_files(sample_dir, DETAIL_HINTS)
    invoice_files = _latest_files(sample_dir, INVOICE_HINTS)

    if not voucher_files:
        raise FileNotFoundError(f"样本目录中找不到凭证文件：{sample_dir}")
    if not detail_files:
        raise FileNotFoundError(f"样本目录中找不到明细文件：{sample_dir}")

    voucher = voucher_files[0]
    detail = detail_files[0]
    invoice = invoice_files[0] if invoice_files else None
    return voucher, detail, invoice


def resolve_inputs(args) -> tuple[Path, Path, Path | None]:
    if args.voucher and args.detail:
        voucher = Path(args.voucher).resolve()
        detail = Path(args.detail).resolve()
        invoice = Path(args.invoice).resolve() if args.invoice else None
        return voucher, detail, invoice
    return auto_select_inputs(Path(args.sample_dir).resolve())


def assert_exists(path: Path, label: str) -> None:
    if not path.exists():
        raise AssertionError(f"{label} 未生成：{path}")


def assert_any_exists(base_dir: Path, filenames: tuple[str, ...], label: str) -> Path:
    for filename in filenames:
        candidate = base_dir / filename
        if candidate.exists():
            return candidate
    raise AssertionError(f"{label} 未生成，候选文件：{', '.join(filenames)}")


def build_parser():
    parser = argparse.ArgumentParser(description="运行审计流程端到端冒烟测试")
    parser.add_argument("--voucher", help="凭证文件路径")
    parser.add_argument("--detail", help="明细文件路径")
    parser.add_argument("--invoice", help="发票台账文件路径")
    parser.add_argument("--sample-dir", default=str(DEFAULT_SAMPLE_DIR), help="自动选样本时的样本目录")
    parser.add_argument("--output-dir", default=str(DEFAULT_OUTPUT_DIR), help="冒烟测试输出目录")
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    try:
        voucher, detail, invoice = resolve_inputs(args)
        output_dir = Path(args.output_dir).resolve()
        output_dir.mkdir(parents=True, exist_ok=True)

        log = setup_logger()
        log.info("== 冒烟测试开始 ==")
        log.info("样本凭证：%s", voucher.name)
        log.info("样本明细：%s", detail.name)
        if invoice:
            log.info("样本发票：%s", invoice.name)
        else:
            log.info("样本发票：未提供，按无发票路径测试")

        pipeline = AuditPipeline(logger=log, template_path=str(ROOT / "项目效益审核表.xlsx"))
        result = pipeline.run(
            voucher_path=str(voucher),
            detail_path=str(detail),
            output_base_dir=str(output_dir),
            invoice_path=str(invoice) if invoice else "",
        )

        result_dir = Path(result.result_dir)
        assert_exists(result_dir, "结果目录")
        for filename in EXPECTED_RESULT_FILES:
            assert_exists(result_dir / filename, filename)
        delivery_file = assert_any_exists(result_dir, EXPECTED_DELIVERY_FILES, "交付文件")
        assert_exists(Path(result.merged_path), "合并表下载路径")
        assert_exists(Path(result.delivery_path), "交付文件下载路径")

        print("\n冒烟测试通过")
        print(f"结果目录: {result_dir}")
        print(f"交付文件: {delivery_file.name}")
        print(f"合并表: {Path(result.merged_path).name}")
        return 0

    except Exception as exc:
        print(f"\n冒烟测试失败: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
