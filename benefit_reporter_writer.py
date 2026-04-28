# -*- coding: utf-8 -*-
"""
效益表写表动作与后处理动作。
"""

from __future__ import annotations

import copy as py_copy
import os
import re

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, Side
from openpyxl.utils import get_column_letter

from benefit_reporter_template import (
    COL_A,
    COL_B,
    COL_C,
    COL_D,
    COL_H,
    COL_J,
    COL_L,
    COL_M,
    COL_N,
    COL_O,
    COL_P,
    COL_Q,
    COL_R,
    COL_S,
    COL_T,
    SUBTOTAL_MAP,
)
from project_config import MATERIAL_FIXED_SUBCATS, VENDOR_SHORTNAME_MAP


class BenefitSheetWriter:
    _KEEP_ROW_KEYWORDS = {
        '（1）集中采购', '(1)集中采购', '（2）自行采购', '(2)自行采购',
        '材料库存', '材料调入', '材料调出', '材料调入+CFK',
        '研发支出', '研发费用'
    }

    def __init__(self, logger, locator):
        self.log = logger
        self.locator = locator

    @staticmethod
    def copy_fill(src_cell, dst_cell):
        try:
            fill = src_cell.fill
            if fill and fill.fill_type not in (None, "none"):
                dst_cell.fill = py_copy.copy(fill)
        except Exception:
            pass

    @staticmethod
    def row_formulas(row_num: int, extended: bool) -> dict:
        formulas = {
            COL_O: f"=L{row_num}-M{row_num}",
            COL_P: f"=K{row_num}-N{row_num}",
            COL_Q: f"=M{row_num}+O{row_num}",
            COL_R: f"=M{row_num}+N{row_num}+O{row_num}+P{row_num}",
        }
        if extended:
            formulas[COL_J] = f"=H{row_num}+I{row_num}"
            formulas[COL_L] = f"=J{row_num}-K{row_num}"
        return formulas

    def style_row(self, ws, row_num, ncol=22, skip_fill=False):
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        font = Font(name="宋体", size=10)
        for col in range(1, ncol + 1):
            cell = ws.cell(row_num, col)
            cell.border = border
            cell.font = font
            cell.alignment = Alignment(
                horizontal="left" if col in (COL_C, COL_D) else "center",
                vertical="center",
                wrap_text=(col in (COL_C, COL_D)),
            )
            if row_num > 1 and not skip_fill:
                self.copy_fill(ws.cell(row_num - 1, col), cell)

    def write_row(self, ws, row_num, d_val, c_val, m_val, extended, skip_fill=False):
        self.style_row(ws, row_num, skip_fill=skip_fill)
        if d_val:
            ws.cell(row_num, COL_D).value = str(d_val)
        if c_val:
            ws.cell(row_num, COL_C).value = str(c_val)
        if m_val is not None:
            cell = ws.cell(row_num, COL_M)
            cell.value = round(float(m_val), 2)
            cell.number_format = '#,##0.00'
        for col, formula in self.row_formulas(row_num, extended).items():
            ws.cell(row_num, col).value = formula

    def delete_stubs(self, ws, hdr_row: int, anchor_row: int):
        protected = {self.locator.norm(k) for k in self._KEEP_ROW_KEYWORDS}
        to_delete = []
        for row in range(hdr_row + 1, anchor_row):
            d_val = ws.cell(row, COL_D).value
            m_val = ws.cell(row, COL_M).value
            if isinstance(m_val, (int, float)):
                continue
            d_norm = self.locator.norm(str(d_val or ''))
            if any(key in d_norm for key in protected):
                continue
            to_delete.append(row)
        for row in reversed(to_delete):
            ws.delete_rows(row)
        if to_delete:
            self.locator.clear_cache()
            self.log.info("      🗑 清除 %d 行（旧数据/公式占位行）", len(to_delete))

    def insert_rows(self, ws, anchor, action, rows_data, extended, is_rd=False):
        for index, row_data in enumerate(rows_data):
            insert_at = anchor + index if action == 'above' else anchor + 1 + index
            ws.insert_rows(insert_at)
            self.write_row(
                ws,
                insert_at,
                row_data.get('d'),
                row_data.get('c'),
                row_data.get('p', 0),
                extended,
                skip_fill=is_rd,
            )
        if rows_data:
            self.locator.clear_cache()

    def ensure_material_fixed_rows(self, ws, extended) -> None:
        hdr_row = self.locator.find_row(ws, '(三)材料费')
        subtotal_row = self.locator.find_row(ws, '材料费小计')
        if not subtotal_row:
            return

        missing = [sub for sub in MATERIAL_FIXED_SUBCATS if self.locator.find_material_row(ws, sub) is None]
        if not missing:
            return

        anchor_row = self.locator.find_anchor(ws, '材料费小计', min_row=hdr_row or 1)
        insert_at = anchor_row or subtotal_row
        for sub in missing:
            ws.insert_rows(insert_at)
            self.write_row(ws, insert_at, sub, None, None, extended)
            insert_at += 1

        self.locator.clear_cache()
        self.log.info("   已自动补插材料费固定行: %s", "、".join(missing))

    def fix_stale_self_refs(self, ws):
        skip_words = {'成本合计', '成本及费用合计', '利润', '增值税实际税负', '税金及附加'}
        target_cols = {COL_J, COL_L, COL_O, COL_P, COL_Q, COL_R, COL_S}

        for row in range(10, ws.max_row + 1):
            d_val = self.locator.norm(str(ws.cell(row, COL_D).value or ''))
            if any(word in d_val for word in skip_words):
                continue
            for col in target_cols:
                cell = ws.cell(row, col)
                if not isinstance(cell.value, str) or not cell.value.startswith('='):
                    continue
                formula = cell.value
                if ':' in formula:
                    matched = re.fullmatch(r'=SUM\(T(\d+):V\1\)', formula, re.IGNORECASE)
                    if matched:
                        old_row = int(matched.group(1))
                        if old_row != row:
                            cell.value = f'=SUM(T{row}:V{row})'
                    continue
                nums = [int(num) for num in re.findall(r'(?<=[A-Za-z])(\d+)', formula)]
                if not nums:
                    continue
                unique = set(nums)
                if len(unique) == 1:
                    old_row = unique.pop()
                    if old_row != row and old_row > 5:
                        cell.value = re.sub(r'(?<=[A-Za-z])' + str(old_row) + r'(?!\d)', str(row), formula)


    def fix_all_sums(self, ws):
        """
        对所有成本节小计行，强制重写全部列的 SUM 公式，
        并强制重写锚点行和自身小计行的行内运算公式（O/P/Q/R/J/L 列）。
        """
        for hdr_kw, sub_kw in SUBTOTAL_MAP.items():
            hdr_row = self.locator.find_row(ws, hdr_kw)
            sub_row = self.locator.find_row(ws, sub_kw)
            if not hdr_row or not sub_row:
                continue

            anchor = None
            for row in range(sub_row - 1, max(hdr_row, sub_row - 400), -1):
                for col in range(2, 6):
                    if '财务未列部分需增列' in self.locator.norm(ws.cell(row, col).value):
                        anchor = row
                        break
                if anchor:
                    break
            end_row = anchor if anchor else sub_row - 1

            for col in range(COL_H, ws.max_column + 1):
                cell = ws.cell(sub_row, col)
                col_letter = get_column_letter(col)
                cell.value = f'=SUM({col_letter}{hdr_row}:{col_letter}{end_row})'
                if col == COL_M and cell.number_format == 'General':
                    cell.number_format = '#,##0.00'

            ws.cell(sub_row, COL_J).value = f'=H{sub_row}+I{sub_row}'
            ws.cell(sub_row, COL_L).value = f'=J{sub_row}-K{sub_row}'
            ws.cell(sub_row, COL_O).value = f'=L{sub_row}-M{sub_row}'
            ws.cell(sub_row, COL_P).value = f'=K{sub_row}-N{sub_row}'
            ws.cell(sub_row, COL_Q).value = f'=M{sub_row}+O{sub_row}'
            ws.cell(sub_row, COL_R).value = f'=M{sub_row}+N{sub_row}+O{sub_row}+P{sub_row}'

            if anchor:
                ws.cell(anchor, COL_J).value = f'=H{anchor}+I{anchor}'
                ws.cell(anchor, COL_L).value = f'=J{anchor}-K{anchor}'
                ws.cell(anchor, COL_O).value = f'=L{anchor}-M{anchor}'
                ws.cell(anchor, COL_P).value = f'=K{anchor}-N{anchor}'
                ws.cell(anchor, COL_Q).value = f'=M{anchor}+O{anchor}'
                ws.cell(anchor, COL_R).value = f'=M{anchor}+N{anchor}+O{anchor}+P{anchor}'

            self.log.info("   🧮 %-14s SUM 全部列重写 行%d→%d；O/P/Q/R 行内运算重写",
                         sub_kw, hdr_row, end_row)

    def fix_aggregate_rows(self, ws):
        """
        强制重写汇总行的所有列公式。
        覆盖行：三、成本合计，四~七各固定项，八、研发费用，八/九、成本及费用合计
        """
        cost_total_row = self.locator.find_row(ws, '三、成本合计')
        if cost_total_row:
            subtotal_keywords = ['人工费小计', '分包工程小计', '材料费小计', '机械费小计', '其他直接费小计', '间接费小计', '安全费小计']
            refs = []
            for kw in subtotal_keywords:
                r = self.locator.find_row(ws, kw)
                if r and r != cost_total_row:
                    refs.append(r)
            if len(refs) == 7:
                for col in range(COL_H, ws.max_column + 1):
                    cl = get_column_letter(col)
                    ws.cell(cost_total_row, col).value = '=' + '+'.join(f'{cl}{r}' for r in refs)
                self.log.info("   🧮 三、成本合计 强制重写 行%d", cost_total_row)

        r_yanfa = self.locator.find_row(ws, '八、研发费用')
        if r_yanfa:
            r_nine = self.locator.find_row(ws, '九、成本及费用合计') or self.locator.find_row(ws, '八、成本及费用合计')
            if r_nine and r_nine > r_yanfa + 1:
                start_row = r_yanfa + 1
                end_row = r_nine - 1
                if end_row < start_row:
                    end_row = start_row
                for col in range(COL_H, ws.max_column + 1):
                    cl = get_column_letter(col)
                    ws.cell(r_yanfa, col).value = f'=SUM({cl}{start_row}:{cl}{end_row})'
                ws.cell(r_yanfa, COL_J).value = f'=H{r_yanfa}+I{r_yanfa}'
                ws.cell(r_yanfa, COL_L).value = f'=J{r_yanfa}-K{r_yanfa}'
                ws.cell(r_yanfa, COL_O).value = f'=L{r_yanfa}-M{r_yanfa}'
                ws.cell(r_yanfa, COL_P).value = f'=K{r_yanfa}-N{r_yanfa}'
                ws.cell(r_yanfa, COL_Q).value = f'=M{r_yanfa}+O{r_yanfa}'
                ws.cell(r_yanfa, COL_R).value = f'=M{r_yanfa}+N{r_yanfa}+O{r_yanfa}+P{r_yanfa}'
                self.log.info("   🧮 八、研发费用 强制重写 行%d", r_yanfa)

        for kw in ['四、计提保修金', '五、过程节点奖金', '六、资金占用费用', '七、局投资收益']:
            r = self.locator.find_row(ws, kw)
            if r:
                ws.cell(r, COL_J).value = f'=H{r}+I{r}'
                ws.cell(r, COL_L).value = f'=J{r}-K{r}'
                ws.cell(r, COL_O).value = f'=L{r}-M{r}'
                ws.cell(r, COL_P).value = f'=K{r}-N{r}'
                ws.cell(r, COL_Q).value = f'=M{r}+O{r}'
                ws.cell(r, COL_R).value = f'=M{r}+N{r}+O{r}+P{r}'

        total_row = self.locator.find_row(ws, '八、成本及费用合计') or self.locator.find_row(ws, '九、成本及费用合计')
        if total_row and cost_total_row:
            related_kws = ['三、成本合计', '四、计提保修金', '五、过程节点奖金', '六、资金占用费用', '七、局投资收益']
            related_rows = []
            for kw in related_kws:
                r = self.locator.find_row(ws, kw)
                if r and r != total_row:
                    related_rows.append(r)
            if related_rows:
                for col in range(COL_H, ws.max_column + 1):
                    cl = get_column_letter(col)
                    ws.cell(total_row, col).value = '=' + '+'.join(f'{cl}{r}' for r in related_rows)
                self.log.info("   🧮 成本及费用合计 强制重写 行%d", total_row)

    def build_lookups(self, df4):
        required_cols = {'总账科目长文本', '供应商', '借方本位币金额'}
        missing = required_cols - set(df4.columns)
        if missing:
            self.log.warning("   ⚠️ 合并表缺少字段，跳过 B/N/T 匹配：%s", "、".join(sorted(missing)))
            return {}, {}, {}

        ven_col = '供应商名称' if '供应商名称' in df4.columns else ('供应商描述' if '供应商描述' in df4.columns else None)

        vendor_lkp = {}
        if ven_col:
            subset = df4[df4[ven_col].notna() & df4['供应商'].notna()]
            dedup_cols = [ven_col, '合同'] if '合同' in subset.columns else [ven_col]
            value_cols = [ven_col, '供应商'] + (['合同'] if '合同' in subset.columns else [])
            for row in subset.drop_duplicates(dedup_cols)[value_cols].itertuples(index=False, name=None):
                ven_raw = row[0]
                code_raw = row[1]
                con_raw = row[2] if len(row) > 2 else ''
                ven = self.locator.norm(str(ven_raw))
                con = self.locator.norm(str(con_raw or ''))
                code = str(int(float(code_raw))) if pd.notna(code_raw) else ''
                if ven:
                    vendor_lkp[(ven, con)] = code
                    if (ven, '') not in vendor_lkp:
                        vendor_lkp[(ven, '')] = code

        ap_mask = (
            df4['总账科目长文本'].fillna('').str.contains('应付账款')
            & ~df4['总账科目长文本'].fillna('').str.contains('待确认进项税额')
        )
        ap = df4[ap_mask].copy()
        ap['_v'] = ap['供应商'].fillna(0).apply(lambda value: str(int(float(value))) if pd.notna(value) and value else '')
        ap['_c'] = ap['合同'].fillna('').astype(str).str.strip()
        pay_lkp = dict(ap.groupby(['_v', '_c'])['借方本位币金额'].sum())

        vat = df4[
            df4['总账科目长文本'].fillna('').str.contains('其他应收款')
            & df4['总账科目长文本'].fillna('').str.contains('待确认进项税额')
        ].copy()
        vat['_v'] = vat['供应商'].fillna(0).apply(lambda value: str(int(float(value))) if pd.notna(value) and value else '')
        vat['_c'] = vat['合同'].fillna('').astype(str).str.strip()
        vat_lkp = dict(vat.groupby(['_v', '_c'])['借方本位币金额'].sum())
        return vendor_lkp, pay_lkp, vat_lkp

    def fill_per_row_cols(self, ws, vendor_lkp, pay_lkp, vat_lkp):
        skip_words = {
            '人工费小计', '分包工程小计', '材料费小计', '机械费小计', '其他直接费小计',
            '间接费小计', '安全费小计', '财务未列部分需增列', '成本合计', '项目效益',
            '成本及费用合计', '增值税', '税金及附加', '利润', '保修金', '节点奖'
        }
        for row in range(10, ws.max_row + 1):
            d_val = ws.cell(row, COL_D).value
            m_val = ws.cell(row, COL_M).value
            c_val = ws.cell(row, COL_C).value
            if not d_val or not isinstance(m_val, (int, float)):
                continue
            if any(word in self.locator.norm(str(d_val)) for word in skip_words):
                continue

            d_norm = self.locator.norm(str(d_val))
            c_norm = self.locator.norm(str(c_val or ''))
            code = vendor_lkp.get((d_norm, c_norm)) or vendor_lkp.get((d_norm, ''), '')

            if not code and d_norm in VENDOR_SHORTNAME_MAP:
                mapped = VENDOR_SHORTNAME_MAP[d_norm]
                code = vendor_lkp.get((self.locator.norm(mapped), c_norm)) or vendor_lkp.get((self.locator.norm(mapped), ''))
            if not code and d_norm:
                log_debug = getattr(self.log, 'debug', None)
                if callable(log_debug):
                    log_debug("未匹配到供应商编码：行%d D='%s' C='%s'", row, d_val, c_val)

            if code:
                ws.cell(row, COL_B).value = code
                pay = pay_lkp.get((code, c_norm), 0)
                if pay:
                    cell = ws.cell(row, COL_T)
                    cell.value = round(float(pay), 2)
                    cell.number_format = '#,##0.00'
                vat = vat_lkp.get((code, c_norm), 0)
                if vat:
                    cell = ws.cell(row, COL_N)
                    cell.value = round(float(vat), 2)
                    cell.number_format = '#,##0.00'

    def apply_invoice_match(self, ws, result_dir: str, invoice_ledger_path: str | None):
        invoice_path = invoice_ledger_path
        if not invoice_path:
            try:
                for filename in os.listdir(result_dir):
                    if filename.lower().endswith((".xlsx", ".xls")) and any(key in filename for key in ["发票", "收票", "台账", "台帳"]):
                        invoice_path = os.path.join(result_dir, filename)
                        break
            except Exception:
                invoice_path = None

        if not (invoice_path and os.path.exists(invoice_path)):
            self.log.info("   ⓘ 未提供发票台账，跳过 AB 列发票匹配")
            return

        from invoice_matcher import InvoiceMatcher

        matcher = InvoiceMatcher(self.log)
        matcher.apply_to_worksheet(ws, invoice_path)

    def renumber(self, ws):
        sections = [
            ('（一）人工费', '人工费小计'),
            ('（二）分包工程', '分包工程小计'),
            ('(三)材料费', '材料费小计'),
            ('(四）机械租赁费', '机械费小计'),
        ]
        for hdr_kw, sub_kw in sections:
            hdr = self.locator.find_row(ws, hdr_kw)
            sub = self.locator.find_row(ws, sub_kw)
            if not hdr or not sub:
                continue
            anchor = None
            for row in range(sub - 1, hdr, -1):
                for col in range(2, 6):
                    if '财务未列部分需增列' in self.locator.norm(ws.cell(row, col).value):
                        anchor = row
                        break
                if anchor:
                    break
            end = (anchor - 1) if anchor else (sub - 1)
            for row in range(hdr + 1, sub):
                ws.cell(row, COL_A).value = None
            seq = 1
            for row in range(hdr + 1, end + 1):
                d_val = ws.cell(row, COL_D).value
                m_val = ws.cell(row, COL_M).value
                if d_val and isinstance(m_val, (int, float)) and m_val != 0:
                    ws.cell(row, COL_A).value = seq
                    seq += 1
