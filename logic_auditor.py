# -*- coding: utf-8 -*-
import pandas as pd
import re
import os

# --- 容差与列号配置 ---
ALLOWED_DIFF = 0.01  # 允许的最大正负差异（0.01元）
COL = {
    'B': 1, 'C': 2, 'D': 3, 'H': 7, 'L': 11, 'M': 12, 'O': 14, 'T': 19, 'AB': 27
}

class LogicAuditor:
    def __init__(self):
        self.merged_subject_col = "总账科目长文本"
        self.merged_debit_col = "借方本位币金额"
        self.merged_credit_col = "贷方本位币金额"

    def _to_float(self, val):
        """将不规则单元格转为浮点数，空值/文字转化为 0"""
        try:
            if pd.isna(val) or val is None or str(val).strip() == '':
                return 0.0
            return float(str(val).replace(',', '').strip())
        except (ValueError, TypeError):
            return 0.0

    def _clean_text(self, val):
        """清洗文本：去空格、去冒号"""
        if pd.isna(val) or val is None:
            return ""
        return re.sub(r'[\s:：]+', '', str(val))

    def _get_merged_sum(self, df_merged, subject_keyword, direction='借方', exact=False):
        """根据科目关键字从合并表提取金额求和"""
        if df_merged is None or df_merged.empty:
            return 0.0
        sub_col = self.merged_subject_col
        if sub_col not in df_merged.columns:
            return 0.0
        val_col = self.merged_debit_col if direction == '借方' else self.merged_credit_col
        if val_col not in df_merged.columns:
            return 0.0

        if exact:
            mask = df_merged[sub_col].str.strip() == subject_keyword
        else:
            mask = df_merged[sub_col].str.contains(subject_keyword.replace('\\', '\\\\'), na=False, regex=True)
            
        return df_merged.loc[mask, val_col].apply(self._to_float).sum()

    def _dynamic_sum(self, df, keyword, idx):
        """
        🚀 核心黑科技：当读取到未被 Excel 计算的公式(0.00)时，
        Python 会智能向上逆推，自己把明细行的 M 列加起来。
        """
        total = 0.0
        if "小计" in keyword:
            for i in range(idx - 1, -1, -1):
                d_text = str(df.iloc[i][COL['D']]).strip()
                if d_text.startswith("（") or "小计" in d_text or "合计" in d_text:
                    break
                total += self._to_float(df.iloc[i][COL['M']])
        elif "研发费用" in keyword:
            for i in range(idx + 1, len(df)):
                d_text = str(df.iloc[i][COL['D']]).strip()
                if d_text.startswith("九、") or "合计" in d_text or "小计" in d_text:
                    break
                total += self._to_float(df.iloc[i][COL['M']])
        elif "合计" in keyword:
            start_idx = 0
            for i in range(idx):
                if "（一）" in str(df.iloc[i][COL['D']]):
                    start_idx = i
                    break
            for i in range(start_idx, idx):
                d_text = str(df.iloc[i][COL['D']]).strip()
                if "小计" not in d_text and "合计" not in d_text and not d_text.startswith("（") and d_text:
                    total += self._to_float(df.iloc[i][COL['M']])
        return total

    def run_audit(self, benefit_path=None, merged_path=None, project_name="未知项目"):
        """执行审核，返回详细的逐条步骤(details)和最终错误(errors)"""
        errors = []
        details = [] 
        
        try:
            df_ben = None
            has_business_data = False 
            
            if benefit_path and os.path.exists(benefit_path):
                df_ben = pd.read_excel(benefit_path, header=None)
                for idx, row in df_ben.iterrows():
                    if len(row) <= COL['H']: continue 
                    text_d = self._clean_text(row[COL['D']])
                    if "三、成本合计" in text_d:
                        val_h = self._to_float(row[COL['H']])
                        # 精度修复：忽略极其微小的浮点数尾数
                        if round(abs(val_h), 2) > ALLOWED_DIFF:
                            has_business_data = True
                        break 
                
            df_mer = None
            if merged_path and os.path.exists(merged_path):
                df_mer = pd.read_excel(merged_path)
                df_mer.columns = [str(c).strip() for c in df_mer.columns]

            # ==========================================
            # 0. 完整性检查步骤
            # ==========================================
            if df_ben is None:
                details.append({"rule": "提取商务表", "status": "warning", "desc": "未检测到效益表，跳过所有比对规则"})
                errors.append({"type": "业务提示", "row": "-", "desc": "当前项目缺乏《商务数据表》（效益表），无法执行校验。"})
            else:
                if has_business_data:
                    details.append({"rule": "提取商务表", "status": "pass", "desc": "已成功提取完整版《效益审核表》（含手工商务数据）"})
                else:
                    details.append({"rule": "提取商务表", "status": "warning", "desc": "智能探测生效：检测到商务成本汇总(H列)为空置，已自动跳过财商对比规则"})
                
            if df_mer is None:
                details.append({"rule": "提取合并表", "status": "warning", "desc": "未找到匹配的《合并表》，跳过底层财务比对"})
                errors.append({"type": "缺失提示", "row": "-", "desc": "缺乏《合并大表》凭证库，无法执行财务账表核对。"})
            else:
                details.append({"rule": "提取合并表", "status": "pass", "desc": f"已成功接入底层凭证库 (读取到 {len(df_mer)} 行明细)"})

            # ==========================================
            # 1. 财务填列校验 (商务表 vs 合并表)
            # ==========================================
            if df_ben is not None and df_mer is not None:
                sums = {
                    "成本合计": -self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\结转", '贷方', exact=False),
                    "人工费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\直接人工费", '借方', exact=False),
                    "分包工程小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\分包工程支出", '借方', exact=False),
                    "材料费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\直接材料费", '借方', exact=False),
                    "机械费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\机械使用费", '借方', exact=False),
                    "其他直接费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\其他直接费用", '借方', exact=False), 
                    "间接费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\间接费用", '借方', exact=False),     
                    "安全费小计": self._get_merged_sum(df_mer, "合同履约成本\工程施工成本\安全生产费", '借方', exact=False),
                    "研发费用": self._get_merged_sum(df_mer, "研发支出", '借方', exact=False)                            
                }
                
                check_items = [
                    ("三、成本合计", "成本合计"), ("人工费小计", "人工费小计"), ("分包工程小计", "分包工程小计"),
                    ("材料费小计", "材料费小计"), ("机械费小计", "机械费小计"), ("其他直接费小计", "其他直接费小计"),
                    ("间接费小计", "间接费小计"), ("安全费小计", "安全费小计"), ("八、研发费用", "研发费用")
                ]
                
                for keyword, sum_key in check_items:
                    rule_err_count = 0
                    for idx, row in df_ben.iterrows():
                        if len(row) <= COL['M']: continue
                        text_d = self._clean_text(row[COL['D']])
                        val_m = self._to_float(row[COL['M']])
                        
                        if self._clean_text(keyword) in text_d: 
                            target_sum = sums[sum_key]
                            
                            # 精度修复：补算判断也加入四舍五入
                            if round(abs(val_m), 2) <= ALLOWED_DIFF and round(abs(target_sum), 2) > ALLOWED_DIFF:
                                val_m = self._dynamic_sum(df_ben, keyword, idx)
                                print(f"🔄 触发补算: 【{keyword}】 | 逆推明细结果: {val_m:,.2f} | 目标: {target_sum:,.2f}")
                            else:
                                print(f"📊 直接比对: 【{keyword}】 | 表内读取: {val_m:,.2f} | 目标: {target_sum:,.2f}")
                            
                            # 🚀 精度修复核心：计算差额后，先保留两位小数，再与 0.01 进行比较
                            diff = round(abs(val_m - target_sum), 2)
                            if diff > ALLOWED_DIFF:
                                rule_err_count += 1
                                errors.append({
                                    "type": "财务填列异常",
                                    "row": idx + 1,
                                    "desc": f"【{keyword}】效益表列报 [{val_m:,.2f}] 与合并表实际发生额 [{target_sum:,.2f}] 不符 (差额: {diff:,.2f})"
                                })
                                
                    if rule_err_count > 0:
                        details.append({"rule": f"表间核对：{keyword}", "status": "fail", "desc": f"发现 {rule_err_count} 处挂账不匹配"})
                    else:
                        details.append({"rule": f"表间核对：{keyword}", "status": "pass", "desc": "账表金额严丝合缝"})

            # ==========================================
            # 2. 财商对比填列校验 (商务表内逻辑) -> 仅当有商务数据时执行！
            # ==========================================
            if df_ben is not None and has_business_data:
                err_o_count = 0
                err_t_count = 0
                err_ab_count = 0
                
                for idx, row in df_ben.iterrows():
                    if len(row) <= COL['AB']: continue 
                    val_b = str(row[COL['B']]).strip()
                    val_c = str(row[COL['C']]).strip()

                    if (val_b and val_b != 'nan') or (val_c and val_c != 'nan'):
                        row_num = idx + 1
                        val_l = self._to_float(row[COL['L']])
                        val_o = self._to_float(row[COL['O']])
                        val_t = self._to_float(row[COL['T']])
                        val_ab = self._to_float(row[COL['AB']])

                        # 精度修复：全部保留两位小数后，看是否严格小于 -0.01
                        if round(val_o, 2) < -ALLOWED_DIFF:
                            err_o_count += 1
                            errors.append({"type": "财商疑问", "row": row_num, "desc": f"可能存在超结/重复入账 (O列金额为 {val_o:,.2f})"})
                        if round(val_t - val_l, 2) < -ALLOWED_DIFF:
                            err_t_count += 1
                            errors.append({"type": "财商疑问", "row": row_num, "desc": f"可能存在超付风险 (T列与L列差额为 {(val_t - val_l):,.2f})"})
                        if round(val_ab - val_l, 2) < -ALLOWED_DIFF:
                            err_ab_count += 1
                            errors.append({"type": "财商疑问", "row": row_num, "desc": f"可能存在无发票付款 (AB列与L列差额为 {(val_ab - val_l):,.2f})"})

                details.append({"rule": "财商：超结重入风险监测", "status": "fail" if err_o_count else "pass", "desc": f"锁定 {err_o_count} 处疑问点" if err_o_count else "未见异常"})
                details.append({"rule": "财商：超付资金流失监测", "status": "fail" if err_t_count else "pass", "desc": f"锁定 {err_t_count} 处疑问点" if err_t_count else "未见异常"})
                details.append({"rule": "财商：无票付款税务风险", "status": "fail" if err_ab_count else "pass", "desc": f"锁定 {err_ab_count} 处疑问点" if err_ab_count else "未见异常"})

        except Exception as e:
            details.append({"rule": "系统异常", "status": "fail", "desc": "读取文件时发生底层异常"})
            errors.append({"type": "解析错误", "row": "-", "desc": f"系统错误：{str(e)}"})

        hard_errors = [e for e in errors if "提示" not in e["type"]]
        is_pass = (len(hard_errors) == 0)

        return {
            "project_name": project_name,
            "errors": errors,
            "details": details,
            "is_pass": is_pass
        }