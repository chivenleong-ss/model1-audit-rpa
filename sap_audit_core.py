# -*- coding: utf-8 -*-
"""
SAPAuditModel 对外入口层。

具体职责已拆到：
1. sap_audit_reader.py   读数与预处理
2. sap_audit_rules.py    对账规则
3. sap_audit_exporter.py 结果导出
"""

from __future__ import annotations

import logging
import os
from datetime import datetime

from sap_audit_exporter import DEFAULT_AUXILIARY_COL_ORDER, DEFAULT_RAW_COLS_TO_DROP, SAPAuditExporter
from sap_audit_reader import DEFAULT_ALIAS_MAP, DEFAULT_JOIN_KEYS, SAPAuditReader
from sap_audit_rules import SAPAuditRules


def setup_logger(log_dir: str = ".") -> logging.Logger:
    logger = logging.getLogger("sap_audit")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    fmt = logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    return logger


class SAPAuditModel:
    def __init__(self, voucher_file, detail_file, output_base_dir='2_输出审计底稿', tolerance=0.01, logger=None):
        self.voucher_path = voucher_file
        self.detail_path = detail_file
        self.tolerance = tolerance
        self.log = logger or setup_logger()

        self.join_keys = list(DEFAULT_JOIN_KEYS)
        self.alias_map = dict(DEFAULT_ALIAS_MAP)
        self.auxiliary_col_order = list(DEFAULT_AUXILIARY_COL_ORDER)
        self.raw_cols_to_drop = list(DEFAULT_RAW_COLS_TO_DROP)

        time_str = datetime.now().strftime('%Y%m%d_%H%M%S')
        self.process_dir = os.path.join(output_base_dir, f"过程表单_{time_str}")
        os.makedirs(self.process_dir, exist_ok=True)
        self.result_dir = os.path.join(output_base_dir, f"审计核对结果_{time_str}")
        os.makedirs(self.result_dir, exist_ok=True)

        self.reader = SAPAuditReader(self.log, join_keys=self.join_keys, alias_map=self.alias_map)
        self.rules = SAPAuditRules(self.log, tolerance=self.tolerance)
        self.exporter = SAPAuditExporter(
            self.log,
            process_dir=self.process_dir,
            result_dir=self.result_dir,
            auxiliary_col_order=self.auxiliary_col_order,
            raw_cols_to_drop=self.raw_cols_to_drop,
        )

    def execute_audit(self):
        self.log.info("=" * 55)
        self.log.info("🚀 启动自动化审计校验模块 (基础比对)")
        self.log.info("=" * 55)

        voucher_df, detail_df = self.reader.load_source_data(self.voucher_path, self.detail_path)
        merged = self.reader.merge_source_data(voucher_df, detail_df)
        merged = self.reader.normalize_amount_columns(merged)
        merged = self.reader.broadcast_group_fields(merged)

        merged = self.rules.apply_audit_rules(merged)
        merged = self.rules.apply_exemption_rules(merged)

        self.exporter.export_process_outputs(merged)
        self.exporter.build_intermediate_tables(merged)
        self.exporter.export_result_outputs(merged)
        return True
