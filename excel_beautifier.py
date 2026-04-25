# -*- coding: utf-8 -*-
"""
excel_beautifier.py v3.1 - 修复循环引用

修复清单：
  1. L列填充色与J列保持一致（逐行复制）
  2. 删除研发费用节内的空公式占位行（无D、无M值的行）
  3. A列序号：补全研发费用节及其他各节的序号
  4. 修复「九、成本及费用合计」公式（#REF! -> 正确行号）
  5. 修复「十二、利润」「十三、利润率」公式中失效的旧行引用
  6. 全局边框补齐（AL列）
  7. 汇总行格式统一（四~八 与三统一样式）
  8. 通用循环引用检测与清理（_clean_circular_refs）
  9. 所有 SUM 公式强制确保引用范围不含自身行
"""

import os, re, traceback
import openpyxl
from openpyxl.utils import get_column_letter, column_index_from_string
from openpyxl.styles import Border, Side, PatternFill
import copy as py_copy

GREY_SET = {"FFBFBFBF", "FFD9D9D9", "FFC0C0C0", "FFE0E0E0","FFD3D3D3", "FFCCCCCC", "FFB8B8B8", "FFAEAAAA"}
WHITE_FILL = PatternFill(start_color="FFFFFFFF", end_color="FFFFFFFF", fill_type="solid")

class ExcelBeautifier:

    def __init__(self, result_dir, logger):
        self.result_dir  = result_dir
        self.log         = logger
        self.input_path  = os.path.join(result_dir, "自动填报完成_效益审核表.xlsx")
        self.output_path = os.path.join(result_dir, "最终完美交付版_效益审核表.xlsx")

    @staticmethod
    def _norm(val) -> str:
        if val is None: return ""
        return (str(val).replace(" ", "").replace("\u3000", "").replace("（", "(").replace("）", ")").replace("：", ":").replace("*", "").replace(":", ""))

    def _find_row(self, ws, keyword, cols=(2, 3, 4, 5)):
        kw = self._norm(keyword)
        for r in range(1, ws.max_row + 1):
            for c in cols:
                val = ws.cell(r, c).value
                if val is not None and kw in self._norm(val):
                    return r
        return None

    @staticmethod
    def _get_fill_rgb(cell) -> str:
        try:
            f = cell.fill
            if f and f.fill_type not in (None, "none"):
                return f.fgColor.rgb
        except Exception:
            pass
        return ""

    # ====================================================================
    #  通用：检测并修复循环引用
    # ====================================================================
    def _clean_circular_refs(self, ws):
        """
        遍历所有包含公式的单元格，检查公式中是否引用了自身单元格。
        如果公式只包含自身引用（如 =A1 且当前在 A1），则清空该单元格。
        如果公式的 SUM/+/ 等范围中包含自身行，则修正范围使其排除当前行。
        """
        fixed = 0
        cleared = 0
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value is None or not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                cl = get_column_letter(c)
                formula = cell.value.upper()
                self_ref = f"{cl}{r}"
                
                # 情况1: 公式只包含自身引用，如 =A1
                if formula == f"={self_ref}":
                    cell.value = 0
                    cleared += 1
                    continue
                
                # 情况2: SUM 范围包含自身行
                if "SUM(" in formula:
                    new_formula = cell.value  # 保留原始大小写
                    # 找到所有 SUM(X:Y) 模式
                    pattern = re.compile(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", re.IGNORECASE)
                    changed = False
                    for match in pattern.finditer(new_formula):
                        col1, row1_str, col2, row2_str = match.group(1), match.group(2), match.group(3), match.group(4)
                        row1 = int(row1_str)
                        row2 = int(row2_str)
                        start_col_letter = col1.upper()
                        end_col_letter = col2.upper()
                        # 检查范围是否包含当前行且当前列在范围内
                        col_idx = c
                        try:
                            start_col = column_index_from_string(start_col_letter)
                            end_col = column_index_from_string(end_col_letter)
                        except Exception:
                            continue
                        if row1 <= r <= row2 and start_col <= col_idx <= end_col:
                            # 需要修正：如果当前行等于某一边界，缩小范围
                            if r == row1 and r < row2:
                                new_row1 = r + 1
                                old_part = f"{col1}{row1}:{col2}{row2}"
                                new_part = f"{col1}{new_row1}:{col2}{row2}"
                                new_formula = new_formula.replace(old_part, new_part, 1)
                                changed = True
                            elif r == row2 and r > row1:
                                new_row2 = r - 1
                                old_part = f"{col1}{row1}:{col2}{row2}"
                                new_part = f"{col1}{row1}:{col2}{new_row2}"
                                new_formula = new_formula.replace(old_part, new_part, 1)
                                changed = True
                            elif row1 < r < row2:
                                # 当前行在范围中间，拆成两个 SUM 或缩小范围
                                # 简单处理：缩小范围到 r-1
                                new_row2 = r - 1
                                old_part = f"{col1}{row1}:{col2}{row2}"
                                new_part = f"{col1}{row1}:{col2}{new_row2}"
                                new_formula = new_formula.replace(old_part, new_part, 1)
                                changed = True
                            elif row1 == r and row2 == r:
                                # 只引用了自身一行，清空
                                cell.value = 0
                                cleared += 1
                                changed = True
                                break
                    if changed and isinstance(cell.value, str) and cell.value.startswith("="):
                        cell.value = new_formula
                        fixed += 1
        
        total = cleared + fixed
        if total:
            self.log.info("   循环引用修复: 清空 %d 个自引用, 修正 %d 个范围", cleared, fixed)

    def _sync_l_fill_to_j(self, ws):
        NO_FILL = PatternFill(fill_type=None)
        fixed = 0
        for r in range(7, ws.max_row + 1):
            j_cell = ws.cell(r, 10)
            l_cell = ws.cell(r, 12)
            j_rgb  = self._get_fill_rgb(j_cell)
            l_rgb  = self._get_fill_rgb(l_cell)
            if j_rgb == l_rgb:
                continue
            try:
                l_cell.fill = py_copy.copy(j_cell.fill) if j_rgb else NO_FILL
                fixed += 1
            except Exception:
                pass
        self.log.info("   L列填充色与J列同步 (%d 个单元格)", fixed)

    def _delete_yanfa_ghost_rows(self, ws):
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if not r_yanfa or not r_nine:
            return
        to_del = []
        for r in range(r_yanfa + 1, r_nine):
            d = ws.cell(r, 4).value
            m = ws.cell(r, 13).value
            d_empty = (d is None) or (isinstance(d, str) and not d.strip())
            if d_empty and not isinstance(m, (int, float)):
                to_del.append(r)
        for r in reversed(to_del):
            ws.delete_rows(r)
        if to_del:
            self.log.info("   删除研发费用节空占位行 %d 行", len(to_del))

    def _delete_cost_zero_m_rows(self, ws):
        SECTIONS = [
            ('（一）人工费',      '人工费小计'),
            ('（二）分包工程',    '分包工程小计'),
            ('(三)材料费',       '材料费小计'),
            ('(四）机械租赁费',   '机械费小计'),
            ('（五）其他直接费',  '其他直接费小计'),
            ('（六）间接费',      '间接费小计'),
            ('（七）安全费',      '安全费小计'),
        ]
        total_deleted = 0
        for hdr_kw, sub_kw in SECTIONS:
            hdr = self._find_row(ws, hdr_kw)
            sub = self._find_row(ws, sub_kw)
            if not hdr or not sub or sub <= hdr:
                continue
            anchor = None
            for r in range(sub - 1, hdr, -1):
                for c in range(2, 6):
                    if '财务未列部分需增列' in self._norm(ws.cell(r, c).value):
                        anchor = r
                        break
                if anchor:
                    break
            end_row = (anchor - 1) if anchor else (sub - 1)
            to_del = []
            for r in range(hdr + 1, end_row + 1):
                d_val = ws.cell(r, 4).value
                m_val = ws.cell(r, 13).value
                has_d = d_val is not None and (not isinstance(d_val, str) or d_val.strip())
                m_empty = (m_val is None) or (isinstance(m_val, (int, float)) and m_val == 0)
                if has_d and m_empty:
                    to_del.append(r)
            for r in reversed(to_del):
                ws.delete_rows(r)
            total_deleted += len(to_del)
            if to_del:
                self.log.info("   清除 %s 节无效行 %d 行（含材料调出、废旧物资等）", hdr_kw, len(to_del))
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if r_yanfa and r_nine and r_nine > r_yanfa + 1:
            to_del = []
            for r in range(r_yanfa + 1, r_nine):
                d_val = ws.cell(r, 4).value
                m_val = ws.cell(r, 13).value
                has_d = d_val is not None and (not isinstance(d_val, str) or d_val.strip())
                m_empty = (m_val is None) or (isinstance(m_val, (int, float)) and m_val == 0)
                if has_d and m_empty:
                    to_del.append(r)
            for r in reversed(to_del):
                ws.delete_rows(r)
            if to_del:
                self.log.info("   清除 八、研发费用 节无效行 %d 行", len(to_del))
                total_deleted += len(to_del)
        if total_deleted:
            self.log.info("   总计清除成本节内空金额行 %d 行", total_deleted)

    def _fix_yanfa_subtotal_formula(self, ws):
        """
        修复"八、研发费用"行的 SUM 引用范围，确保不包含自身行。
        """
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if not r_yanfa or not r_nine or r_nine <= r_yanfa:
            return
        start_row = r_yanfa + 1
        end_row = min(start_row, r_nine - 1)
        if end_row < start_row:
            end_row = start_row
        for col in range(1, ws.max_column + 1):
            cell = ws.cell(r_yanfa, col)
            if cell.value and isinstance(cell.value, str) and 'SUM(' in cell.value.upper():
                cl = get_column_letter(col)
                cell.value = f"=SUM({cl}{start_row}:{cl}{end_row})"
        self.log.info("   研发费用小计公式修正（行 %d，范围 %d-%d）", r_yanfa, start_row, end_row)

    def _fill_sequence_numbers(self, ws):
        SECTIONS = [('（一）人工费','人工费小计'),('（二）分包工程','分包工程小计'),('(三)材料费','材料费小计'),('(四）机械租赁费','机械费小计'),('（五）其他直接费','其他直接费小计'),('（六）间接费','间接费小计'),('（七）安全费','安全费小计')]
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        def _number_section(hdr, end_excl):
            for r in range(hdr + 1, end_excl):
                ws.cell(r, 1).value = None
            seq = 1
            for r in range(hdr + 1, end_excl):
                d = ws.cell(r, 4).value
                m = ws.cell(r, 13).value
                if d and isinstance(m, (int, float)):
                    ws.cell(r, 1).value = seq
                    seq += 1
        for hdr_kw, sub_kw in SECTIONS:
            hdr = self._find_row(ws, hdr_kw)
            sub = self._find_row(ws, sub_kw)
            if not hdr or not sub:
                continue
            anchor = None
            for r in range(sub - 1, hdr, -1):
                for c in range(2, 6):
                    if '财务未列部分需增列' in self._norm(ws.cell(r, c).value):
                        anchor = r; break
                if anchor: break
            end = (anchor - 1) if anchor else (sub - 1)
            _number_section(hdr, end + 1)
        if r_yanfa and r_nine:
            _number_section(r_yanfa, r_nine)
        self.log.info("   A列序号全节补全完成")

    def _fix_yanfa_detail_fill(self, ws):
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if not r_yanfa or not r_nine:
            return
        no_fill = PatternFill(fill_type=None)
        ref_fill_c = ref_fill_d = None
        for r in range(r_yanfa + 1, r_nine):
            d = ws.cell(r, 4).value
            m = ws.cell(r, 13).value
            if d and isinstance(m, (int, float)):
                ref_fill_c = py_copy.copy(ws.cell(r, 3).fill)
                ref_fill_d = py_copy.copy(ws.cell(r, 4).fill)
                break
        fixed = 0
        for r in range(r_yanfa + 1, r_nine):
            d = ws.cell(r, 4).value
            m = ws.cell(r, 13).value
            if d and isinstance(m, (int, float)):
                ws.cell(r, 3).fill = py_copy.copy(ref_fill_c) if ref_fill_c else py_copy.copy(no_fill)
                ws.cell(r, 4).fill = py_copy.copy(ref_fill_d) if ref_fill_d else py_copy.copy(no_fill)
                fixed += 1
        if fixed:
            self.log.info("   研发费用明细 C/D 列填充修正 %d 行", fixed)

    def _clear_yanfa_section_s_fill(self, ws):
        r_yanfa = self._find_row(ws, '八、研发费用')
        r_nine = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if not r_yanfa or not r_nine:
            return
        no_fill = PatternFill(fill_type=None)
        fixed = 0
        for r in range(r_yanfa + 1, r_nine):
            ws.cell(r, 19).fill = py_copy.copy(no_fill)
            fixed += 1
        if fixed:
            self.log.info("   研发费用区间 S 列填充清空 %d 行", fixed)

    def _fix_all_sums(self, ws):
        """
        遍历全表所有包含 SUM 公式的单元格，
        确保引用范围不包含公式所在行本身。
        """
        fixed = 0
        for r in range(1, ws.max_row + 1):
            for c in range(1, ws.max_column + 1):
                cell = ws.cell(r, c)
                if cell.value is None or not isinstance(cell.value, str) or not cell.value.startswith("="):
                    continue
                formula = cell.value
                if "SUM(" not in formula.upper():
                    continue
                cl = get_column_letter(c)
                new_formula = formula
                changed = False
                # 找到所有 SUM(Xnn:Xnn) 模式
                pattern = re.compile(r"SUM\(([A-Z]+)(\d+):([A-Z]+)(\d+)\)", re.IGNORECASE)
                for match in pattern.finditer(formula):
                    col1, rs1, col2, rs2 = match.group(1), match.group(2), match.group(3), match.group(4)
                    row_start = int(rs1)
                    row_end = int(rs2)
                    col1_upper = col1.upper()
                    col2_upper = col2.upper()
                    try:
                        ci1 = column_index_from_string(col1_upper)
                        ci2 = column_index_from_string(col2_upper)
                    except Exception:
                        continue
                    # 检查：当前行 r 是否在 [row_start, row_end] 范围内且当前列 c 在 [ci1, ci2] 内
                    in_col_range = (ci1 <= c <= ci2) or (ci2 <= c <= ci1)
                    if not (row_start <= r <= row_end and in_col_range):
                        continue
                    old_part = match.group(0)
                    # 排除自身行
                    if row_start == r == row_end:
                        # 只引用自身一行，清空
                        cell.value = 0
                        changed = True
                        break
                    elif r == row_start:
                        new_range = f"{col1}{r + 1}:{col2}{row_end}"
                        new_part = f"SUM({new_range})"
                        new_formula = new_formula.replace(old_part, new_part, 1)
                        changed = True
                    elif r == row_end:
                        new_range = f"{col1}{row_start}:{col2}{r - 1}"
                        new_part = f"SUM({new_range})"
                        new_formula = new_formula.replace(old_part, new_part, 1)
                        changed = True
                    else:
                        # r 在中间，分成两段或直接缩小
                        new_range = f"{col1}{row_start}:{col2}{r - 1}"
                        new_part = f"SUM({new_range})"
                        new_formula = new_formula.replace(old_part, new_part, 1)
                        changed = True
                if changed and isinstance(cell.value, str) and cell.value.startswith("="):
                    cell.value = new_formula
                    fixed += 1
        if fixed:
            self.log.info("   全局 SUM 范围修正 %d 个单元格（排除自引用）", fixed)

    def _fix_nine_formula(self, ws):
        """
        修复「九、成本及费用合计」公式，确保引用的行号中不包含自身行。
        """
        ROW_KWS = ['三、成本合计','四、计提保修金','五、过程节点奖金','六、资金占用费用','七、局投资收益','八、研发费用']
        row_nums = []
        for kw in ROW_KWS:
            r = self._find_row(ws, kw)
            if r:
                row_nums.append(r)
        r_nine = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        if not r_nine or not row_nums:
            self.log.warning("   未能定位九、成本及费用合计，跳过公式修复")
            return
        # 安全过滤：排除等于 r_nine 自身行的引用
        safe_rows = [r for r in row_nums if r != r_nine]
        if not safe_rows:
            self.log.warning("   九、成本及费用合计 所有引用行均等于自身行，跳过")
            return
        for col in range(8, ws.max_column + 1):
            cell = ws.cell(r_nine, col)
            if cell.value is None:
                continue
            cl = get_column_letter(col)
            cell.value = '=' + '+'.join(f'{cl}{r}' for r in safe_rows)
        self.log.info("   九、成本及费用合计 公式修正（参考行 %s）", '+'.join(str(r) for r in safe_rows))

    def _fix_profit_formulas(self, ws):
        """
        修复十二、利润 / 十三、利润率 公式，确保不引用自身行。
        """
        r6 = 6
        r_nine  = self._find_row(ws, '九、成本及费用合计') or self._find_row(ws, '八、成本及费用合计')
        r_vat   = self._find_row(ws, '十、增值税实际税负')
        r_tax   = self._find_row(ws, '十一、税金及附加')
        r_profit = self._find_row(ws, '十二、利润')
        r_rate   = self._find_row(ws, '十三、利润率')
        if r_profit and r_nine:
            deduct = [r for r in [r_nine, r_vat, r_tax] if r and r != r_profit]
            for col in range(8, ws.max_column + 1):
                cell = ws.cell(r_profit, col)
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    cl = get_column_letter(col)
                    cell.value = f'={cl}{r6}' + ''.join(f'-{cl}{r}' for r in deduct)
            self.log.info("   十二、利润 公式修正（行 %d）", r_profit)
        if r_rate and r_profit:
            for col in range(8, ws.max_column + 1):
                cell = ws.cell(r_rate, col)
                if cell.value and isinstance(cell.value, str) and cell.value.startswith('='):
                    cl = get_column_letter(col)
                    cell.value = f'={cl}{r_profit}/{cl}{r6}'
            self.log.info("   十三、利润率 公式修正（行 %d）", r_rate)

    def _delete_trailing_empty_rows(self, ws):
        deleted = 0
        for r in range(ws.max_row, 10, -1):
            if all(ws.cell(r, c).value is None for c in range(1, 40)):
                ws.delete_rows(r)
                deleted += 1
            else:
                break
        if deleted:
            self.log.info("   删除尾部空行 %d 行", deleted)

    def _unify_summary_rows(self, ws, max_col=38):
        r_base = self._find_row(ws, '三、成本合计')
        if not r_base: return
        KEYWORDS = ["四、计提", "五、过程", "六、资金", "七、局投资", "八、研发"]
        for r in range(r_base + 1, ws.max_row + 1):
            val = self._norm(ws.cell(r, 4).value or "")
            if any(k in val for k in KEYWORDS):
                for c in range(1, max_col + 1):
                    src = ws.cell(r_base, c)
                    tgt = ws.cell(r, c)
                    try:
                        if src.font:      tgt.font      = py_copy.copy(src.font)
                        if src.fill:      tgt.fill      = py_copy.copy(src.fill)
                        if src.alignment: tgt.alignment = py_copy.copy(src.alignment)
                    except Exception:
                        pass

    def _apply_borders(self, ws, max_col=38):
        thin = Side(border_style="thin", color="000000")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)
        for r in range(7, ws.max_row + 1):
            if not ws.cell(r, 4).value and not ws.cell(r, 13).value:
                continue
            for c in range(1, max_col + 1):
                ws.cell(r, c).border = border

    def _load_workbook(self):
        wb = openpyxl.load_workbook(self.input_path)
        ws = wb.active
        for sn in wb.sheetnames:
            if any(k in sn for k in ["效益审核", "附表"]):
                ws = wb[sn]
                break
        return wb, ws

    def _run_structure_repairs(self, ws):
        self._delete_trailing_empty_rows(ws)
        self._delete_yanfa_ghost_rows(ws)
        self._delete_cost_zero_m_rows(ws)

    def _run_formula_repairs(self, ws):
        self._clean_circular_refs(ws)
        self._fix_all_sums(ws)
        self._fix_yanfa_subtotal_formula(ws)
        self._fix_nine_formula(ws)
        self._fix_profit_formulas(ws)
        self._fill_sequence_numbers(ws)

    def _run_style_repairs(self, ws):
        self._fix_yanfa_detail_fill(ws)
        self._clear_yanfa_section_s_fill(ws)
        self._unify_summary_rows(ws)
        self._sync_l_fill_to_j(ws)
        self._apply_borders(ws)

    def _save_workbook(self, wb):
        wb.save(self.output_path)
        self.log.info("交付版生成完成：%s", self.output_path)

    def execute_beautify(self):
        if not os.path.exists(self.input_path):
            self.log.error("找不到输入文件：%s", self.input_path)
            return False
        self.log.info("美化引擎 v3.1 启动（含循环引用修复）...")
        try:
            wb, ws = self._load_workbook()
            self._run_structure_repairs(ws)
            self._run_formula_repairs(ws)
            self._run_style_repairs(ws)
            self._save_workbook(wb)
            return True
        except Exception as e:
            self.log.error("美化异常:\n%s", traceback.format_exc())
            return False
