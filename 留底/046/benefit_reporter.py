# benefit_reporter.py（第三层收口后）
# -*- coding: utf-8 -*-
"""
BenefitReporter 现在只负责填报编排。

具体职责已拆到：
1. benefit_reporter_template.py   模板定位/表头定位
2. benefit_reporter_aggregator.py 聚合规则
3. benefit_reporter_writer.py     写表/补写/公式修复
"""

from __future__ import annotations

import os
import traceback

import pandas as pd

from benefit_reporter_aggregator import BenefitDataAggregator
from benefit_reporter_template import MATERIAL_ROW_ALIASES, BenefitTemplateLocator
from benefit_reporter_writer import BenefitSheetWriter
from project_config import MATERIAL_FIXED_SUBCATS


class BenefitReporter:
    def __init__(self, result_dir, logger):
        self.result_dir = result_dir
        self.log = logger
        self.template_path = "项目效益审核表.xlsx"
        self.invoice_ledger_path = None

        self.locator = BenefitTemplateLocator(logger)
        self.aggregator = BenefitDataAggregator()
        self.writer = BenefitSheetWriter(logger, self.locator)

    @staticmethod
    def _source_paths(result_dir: str) -> tuple[str, str]:
        return (
            os.path.join(result_dir, "5_中间汇总表_效益审核数据源.xlsx"),
            os.path.join(result_dir, "4_合并表.xlsx"),
        )

    def _validate_inputs(self, source_path: str) -> bool:
        if not os.path.exists(source_path):
            self.log.error("⚠️ 找不到中间汇总表：%s", source_path)
            return False
        if not os.path.exists(self.template_path):
            self.log.error("⚠️ 找不到模板：%s", self.template_path)
            return False
        return True

    def _load_fill_data(self, source_path: str, merged_path: str):
        df_raw = pd.read_excel(source_path)
        df = df_raw[df_raw['利润中心'].astype(str).str.strip() != '筛选合计：'].copy()
        df['最终发生额'] = pd.to_numeric(df['最终发生额'], errors='coerce').fillna(0)
        for col in ['合同编码', '细分科目', '客商名称', '中台单据号']:
            if col in df.columns:
                df[col] = df[col].fillna('').astype(str).str.strip().replace('nan', '')
        df4 = pd.read_excel(merged_path) if os.path.exists(merged_path) else pd.DataFrame()
        return df, df4

    def _prepare_lookup_tables(self, df4):
        if df4.empty:
            return {}, {}, {}
        return self.writer.build_lookups(df4)

    @staticmethod
    def _insert_plan():
        return [
            ('(一)人工费', 'labor', '人工费小计', 'above', False),
            ('(二)分包工程', 'vendor', '分包工程小计', 'above', False),
            ('(三)材料费', 'material', '材料费小计', 'above', False),
            ('(四)机械租赁费', 'vendor', '机械费小计', 'above', True),
            ('(五)其他直接费', 'sub', '其他直接费小计', 'above', True),
            ('(六)间接费', 'sub', '间接费小计', 'above', True),
            ('(七)安全费', 'sub', '安全费小计', 'above', True),
            ('研发支出', 'sub', '八、研发费用', 'below', False),
        ]

    @staticmethod
    def _header_keyword_map():
        return {
            '(一)人工费': '（一）人工费',
            '(二)分包工程': '（二）分包工程',
            '(三)材料费': '(三)材料费',
            '(四)机械租赁费': '(四）机械租赁费',
            '(五)其他直接费': '（五）其他直接费',
            '(六)间接费': '（六）间接费',
            '(七)安全费': '（七）安全费',
        }

    def _build_rows_for_category(self, cat_df, mode):
        if mode == 'labor':
            return self.aggregator.aggregate_labor(cat_df)
        if mode == 'vendor':
            return self.aggregator.aggregate_vendor(cat_df)
        if mode == 'material':
            return self.aggregator.aggregate_material(cat_df)
        return self.aggregator.aggregate_sub(cat_df)

    def _locate_insert_anchor(self, ws, cat, anchor_kw, action):
        if action == 'above':
            hdr_kw = self._header_keyword_map().get(cat, '')
            hdr_row = self.locator.find_row(ws, hdr_kw) if hdr_kw else None
            anchor = self.locator.find_anchor(ws, anchor_kw, min_row=hdr_row or 1)
            if anchor is None:
                anchor = self.locator.find_row(ws, anchor_kw)
            return hdr_row, anchor
        return None, self.locator.find_row(ws, anchor_kw)

    def _insert_category_rows(self, ws, df):
        last_extended = False
        for cat, mode, anchor_kw, action, extended in self._insert_plan():
            cat_df = df[df['成本_财务大类'] == cat].copy()
            if cat_df.empty:
                self.log.info("   ℹ️  [%s] 无数据", cat)
                continue

            rows_data = self._build_rows_for_category(cat_df, mode)
            if not rows_data:
                continue

            hdr_row, anchor = self._locate_insert_anchor(ws, cat, anchor_kw, action)
            if anchor is None:
                self.log.warning("   ⚠️  [%s] 锚点未找到（'%s'）", cat, anchor_kw)
                continue

            if action == 'above' and hdr_row:
                self.writer.delete_stubs(ws, hdr_row, anchor)
                anchor = self.locator.find_anchor(ws, anchor_kw, min_row=hdr_row) or self.locator.find_row(ws, anchor_kw)

            self.log.info("   📍 [%s] 锚点=行%-4d  插入 %d 行", cat, anchor, len(rows_data))
            self.writer.insert_rows(ws, anchor, action, rows_data, extended, is_rd=(cat == '研发支出'))
            last_extended = extended

        self.writer.ensure_material_fixed_rows(ws, last_extended)

    def _repair_formulas(self, ws):
        self.writer.fix_stale_self_refs(ws)
        self.writer.fix_all_sums(ws)
        self.writer.fix_aggregate_rows(ws)

    def _fill_single_fixed_item(self, ws, df, cat_or_filter, kw, label):
        if isinstance(cat_or_filter, str):
            amount = df[df['成本_财务大类'] == cat_or_filter]['最终发生额'].sum()
        else:
            mask = pd.Series(True, index=df.index)
            for key, value in cat_or_filter.items():
                mask &= (df[key] == value)
            amount = df[mask]['最终发生额'].sum()

        row = self.locator.find_material_row(ws, kw) if kw in MATERIAL_ROW_ALIASES else self.locator.find_row(ws, kw)
        if row and amount != 0:
            ws.cell(row, 13).value = round(float(amount), 2)
            ws.cell(row, 13).number_format = '#,##0.00'
            self.log.info("   ✅ %-20s 行%d  M=%.2f", label, row, amount)
        elif not row:
            self.log.warning("   ⚠️  找不到关键词 '%s'", kw)

    def _fill_fixed_items(self, ws, df):
        self._fill_single_fixed_item(ws, df, '六、资金占用费用', '六、资金占用费用', '资金占用费用')
        self._fill_single_fixed_item(ws, df, '七、局投资收益（局投资项目选填）', '七、局投资收益', '局投资收益')

        for sub in MATERIAL_FIXED_SUBCATS:
            self._fill_single_fixed_item(
                ws,
                df,
                {'成本_财务大类': '(三)材料费', '细分科目': sub},
                sub,
                sub,
            )

        tax_amt = df[df['成本_财务大类'].fillna('').str.contains('税金及附加')]['最终发生额'].sum()
        tax_row = self.locator.find_row(ws, '十一、税金及附加')
        if tax_row and tax_amt:
            ws.cell(tax_row, 13).value = round(float(tax_amt), 2)
            ws.cell(tax_row, 13).number_format = '#,##0.00'
            self.log.info("   ✅ 税金及附加 行%d M=%.2f", tax_row, tax_amt)

    def _fill_lookup_columns(self, ws, df4, vendor_lkp, pay_lkp, vat_lkp):
        if not df4.empty:
            self.writer.fill_per_row_cols(ws, vendor_lkp, pay_lkp, vat_lkp)
        self.log.info("   ✅ B/N/T 列匹配填写完成")

    def _finalize_workbook(self, wb, ws, output_path: str):
        self.writer.apply_invoice_match(ws, self.result_dir, self.invoice_ledger_path)
        self.writer.renumber(ws)
        self.log.info("   ✅ A列序号重排完成")
        wb.save(output_path)
        self.log.info("🎉 填报完成：%s", output_path)

    def execute_fill(self):
        source_path, merged_path = self._source_paths(self.result_dir)
        if not self._validate_inputs(source_path):
            return False

        self.log.info("📝 填报引擎 v6 启动（第三层拆分版）...")
        try:
            df, df4 = self._load_fill_data(source_path, merged_path)
            vendor_lkp, pay_lkp, vat_lkp = self._prepare_lookup_tables(df4)
            wb, ws = self.locator.load_template_sheet(self.template_path)

            self.locator.fill_header_fields(ws, df, df4)
            self._insert_category_rows(ws, df)
            self._repair_formulas(ws)
            self._fill_fixed_items(ws, df)
            self._fill_lookup_columns(ws, df4, vendor_lkp, pay_lkp, vat_lkp)

            out = os.path.join(self.result_dir, "自动填报完成_效益审核表.xlsx")
            self._finalize_workbook(wb, ws, out)
            return True

        except Exception:
            self.log.error("❌ 填报异常:\n%s", traceback.format_exc())
            return False
