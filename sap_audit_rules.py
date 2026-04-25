# -*- coding: utf-8 -*-
"""
SAP 审计对账规则模块。
"""

from __future__ import annotations

import numpy as np
import pandas as pd


class SAPAuditRules:
    def __init__(self, logger, tolerance: float = 0.01):
        self.log = logger
        self.tolerance = tolerance

    def apply_audit_rules(self, merged):
        voucher_amt = merged["本位币金额"].abs()
        detail_amt_signed = merged["借方本位币金额"].fillna(0) - merged["贷方本位币金额"].fillna(0)
        detail_amt = detail_amt_signed.abs()

        merged["发生额差额"] = (voucher_amt - detail_amt).round(2).astype(str)
        merged["发生额差额_含符号"] = (merged["本位币金额"] - detail_amt_signed).round(2).astype(str)
        merged["金额校验"] = np.where(
            (voucher_amt - detail_amt).round(2).abs() <= self.tolerance,
            "✅ 金额正确",
            "❌ 金额异常",
        )
        merged["方向校验"] = np.where(
            (merged["金额校验"] == "✅ 金额正确")
            & ((merged["本位币金额"] - detail_amt_signed).round(2).abs() > self.tolerance),
            "⚠️ 借贷方向疑似反向",
            "",
        )
        return merged

    @staticmethod
    def _exemption_mask(merged):
        mask = pd.Series(False, index=merged.index)
        if "文本" in merged.columns:
            mask = mask | (merged["文本"] == "自动清账剩余项目")
        if "反记帐" in merged.columns:
            mask = mask | (merged["反记帐"].notna() & (merged["反记帐"] != ""))
        return mask

    def apply_exemption_rules(self, merged):
        ex_mask = self._exemption_mask(merged)
        merged.loc[ex_mask, "发生额差额"] = "✅ 清账/反记账豁免"
        merged.loc[ex_mask, "发生额差额_含符号"] = "✅ 清账/反记账豁免"
        merged.loc[ex_mask, "方向校验"] = ""
        return merged
