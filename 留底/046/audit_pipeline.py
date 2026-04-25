# -*- coding: utf-8 -*-
"""
主流程编排层。

统一串联：
1. 审计底稿生成
2. 效益审核表填报
3. 交付版美化

这样网页入口和命令行入口都只负责触发，不再各自维护一套流程。
"""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from benefit_reporter import BenefitReporter
from excel_beautifier import ExcelBeautifier
from sap_audit_core import SAPAuditModel, setup_logger


DELIVERY_CANDIDATES = (
    "最终完美交付版_效益审核表.xlsx",
    "交付版_效益审核表.xlsx",
    "自动填报完成_效益审核表.xlsx",
)
MERGED_CANDIDATES = ("4_合并表.xlsx",)


@dataclass
class PipelineResult:
    result_dir: str
    delivery_path: str
    merged_path: str


class AuditPipeline:
    def __init__(self, logger=None, template_path: str | None = None):
        self.log = logger or setup_logger()
        self.template_path = template_path or str(Path(__file__).resolve().parent / "项目效益审核表.xlsx")

    def _set_progress(self, progress: int) -> None:
        setter = getattr(self.log, "set_progress", None)
        if callable(setter):
            setter(progress)

    @staticmethod
    def _find_output_file(result_dir: str, filenames: tuple[str, ...]) -> str:
        base = Path(result_dir)
        for filename in filenames:
            path = base / filename
            if path.exists():
                return str(path)
        return ""

    def _validate_template(self) -> None:
        if not os.path.exists(self.template_path):
            raise FileNotFoundError(f"模板文件不存在：{self.template_path}")

    def _run_audit(self, voucher_path: str, detail_path: str, output_base_dir: str) -> SAPAuditModel:
        self._set_progress(10)
        self.log.info("🚀 启动数据合并与审计校验...")
        audit = SAPAuditModel(voucher_path, detail_path, output_base_dir=output_base_dir, logger=self.log)
        if not audit.execute_audit():
            raise RuntimeError("数据结构异常，审计模块返回失败")
        return audit

    def _run_fill(self, result_dir: str, invoice_path: str = "") -> BenefitReporter:
        self._set_progress(45)
        self._validate_template()
        self.log.info("📝 启动效益审核表自动填报...")
        reporter = BenefitReporter(result_dir, self.log)
        reporter.template_path = self.template_path
        if invoice_path and os.path.exists(invoice_path):
            reporter.invoice_ledger_path = invoice_path
        if not reporter.execute_fill():
            raise RuntimeError("效益审核表填报失败")
        return reporter

    def _run_beautify(self, result_dir: str) -> ExcelBeautifier:
        self._set_progress(80)
        self.log.info("💅 启动格式美化引擎...")
        beautifier = ExcelBeautifier(result_dir, self.log)
        if not beautifier.execute_beautify():
            raise RuntimeError("效益审核表美化失败")
        return beautifier

    def _collect_outputs(self, result_dir: str) -> PipelineResult:
        self._set_progress(96)
        delivery_path = self._find_output_file(result_dir, DELIVERY_CANDIDATES)
        merged_path = self._find_output_file(result_dir, MERGED_CANDIDATES)
        if not delivery_path:
            raise FileNotFoundError("交付版文件未生成")
        return PipelineResult(
            result_dir=result_dir,
            delivery_path=delivery_path,
            merged_path=merged_path,
        )

    def run(
        self,
        voucher_path: str,
        detail_path: str,
        output_base_dir: str,
        invoice_path: str = "",
    ) -> PipelineResult:
        audit = self._run_audit(voucher_path, detail_path, output_base_dir)
        self._run_fill(audit.result_dir, invoice_path=invoice_path)
        self._run_beautify(audit.result_dir)
        result = self._collect_outputs(audit.result_dir)
        self.log.info("🎉 全部完成！文件已就绪。")
        return result
