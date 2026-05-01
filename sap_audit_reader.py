# -*- coding: utf-8 -*-
"""
SAP 审计读数与预处理模块。
"""

from __future__ import annotations

import os

import numpy as np
import pandas as pd


DEFAULT_JOIN_KEYS = ['公司代码', '财年', '凭证编号', '行项目']
DEFAULT_ALIAS_MAP = {
    '公司代码': ['公司', 'Company Code', 'BUKRS'],
    '财年': ['会计年度', '年度', 'Fiscal Year', 'GJAHR', '年度/期间'],
    '凭证编号': ['凭证号', '会计凭证', 'Document Number', 'BELNR'],
    '行项目': ['项目', '行号', 'Item No', 'BUZEI'],
    '本位币金额': ['金额', '主表金额', 'Amount in LC', 'DMBTR'],
    '借方本位币金额': ['借方金额', '借方', 'Debit'],
    '贷方本位币金额': ['贷方金额', '贷方', 'Credit'],
    '文本': ['摘要', 'Item Text', 'SGTXT', 'Text'],
    '反记帐': ['Reverse posting', 'XNEGP', '反记账标识', '反记账'],
    '供应商名称': ['供应商描述'],
    '总账科目': ['科目号'],
    '利润中心文本描述': ['利润中心描述'],
}


class SAPAuditReader:
    def __init__(self, logger, join_keys=None, alias_map=None):
        self.log = logger
        self.join_keys = join_keys or list(DEFAULT_JOIN_KEYS)
        self.alias_map = alias_map or dict(DEFAULT_ALIAS_MAP)

    def smart_read(self, path):
        self.log.info("读取文件: %s", os.path.basename(path))
        is_csv = path.lower().endswith('.csv')

        # --- 🚀 智能探针：寻找真实表头，跳过前导废话行 ---
        try:
            if is_csv:
                try:
                    df_probe = pd.read_csv(path, nrows=20, header=None, low_memory=False)
                except Exception:
                    df_probe = pd.read_csv(path, encoding='gbk', nrows=20, header=None, low_memory=False)
            else:
                df_probe = pd.read_excel(path, nrows=20, header=None)

            header_row = 0
            # 探测关键字：只要行内包含以下核心字段，即判定为真实表头
            keywords = ['公司代码', '公司', '凭证编号', '凭证号', '会计凭证']
            for idx, row in df_probe.iterrows():
                row_strs = [str(x).strip() for x in row.values if pd.notna(x)]
                if any(k in row_strs for k in keywords):
                    header_row = idx
                    break

            # 使用找到的真实行号正式读取
            if is_csv:
                try:
                    df = pd.read_csv(path, header=header_row, low_memory=False)
                except Exception:
                    df = pd.read_csv(path, encoding='gbk', header=header_row, low_memory=False)
            else:
                df = pd.read_excel(path, header=header_row)

            # 清理全空行及表头前后看不见的空格
            df = df.dropna(how='all')
            df.columns = [str(c).strip() for c in df.columns]
            return df

        except Exception as e:
            self.log.warning("智能探针探测表头异常，退回传统读取模式: %s", str(e))
            # 兜底：如果探针失败，退回最原始的盲读方式
            if is_csv:
                try:
                    return pd.read_csv(path, low_memory=False)
                except Exception:
                    return pd.read_csv(path, encoding='gbk', low_memory=False)
            return pd.read_excel(path).dropna(how='all')

    def preprocess(self, df):
        for std_name, aliases in self.alias_map.items():
            for alias in aliases:
                if alias in df.columns and std_name not in df.columns:
                    df = df.rename(columns={alias: std_name})

        for key in self.join_keys:
            if key in df.columns:
                df[key] = (
                    df[key]
                    .astype(str)
                    .str.strip()
                    .str.replace(r'\.0$', '', regex=True)
                    .replace(['nan', 'None', ''], np.nan)
                )

        key_cols = [key for key in self.join_keys if key in df.columns]
        if key_cols:
            before_len = len(df)
            df = df.dropna(subset=key_cols, how='any')
            removed = before_len - len(df)
            if removed > 0:
                self.log.info("   🗑️ 清洗拦截了 %d 行无效数据。", removed)
        return df

    def load_source_data(self, voucher_path: str, detail_path: str):
        voucher_df = self.preprocess(self.smart_read(voucher_path))
        detail_df = self.preprocess(self.smart_read(detail_path))
        return voucher_df, detail_df

    def merge_source_data(self, voucher_df, detail_df):
        merged = pd.merge(
            voucher_df,
            detail_df,
            on=self.join_keys,
            how="outer",
            indicator="匹配状态",
            suffixes=("", "_重复待删"),
        )
        merged["匹配状态"] = merged["匹配状态"].astype(str).replace({
            "both": "✅ 完全匹配",
            "left_only": "❌ 仅主表有(缺明细)",
            "right_only": "❌ 仅明细有(单边账)",
        })
        return merged

    @staticmethod
    def numeric_columns():
        return ["本位币金额", "借方本位币金额", "贷方本位币金额"]

    def normalize_amount_columns(self, merged):
        for col in self.numeric_columns():
            if col in merged.columns:
                merged[col] = pd.to_numeric(merged[col], errors="coerce").fillna(0)
        return merged

    def broadcast_group_fields(self, merged):
        self.log.info("🔄 正在执行同凭证内信息智能广播补齐 (客商、合同)...")
        fill_cols = ['供应商', '供应商名称', '合同', '合同文本描述', '客户', '客户描述']
        existing_fill_cols = [col for col in fill_cols if col in merged.columns]
        group_keys = [key for key in ['公司代码', '财年', '凭证编号'] if key in merged.columns]
        if not (group_keys and existing_fill_cols):
            return merged

        for col in existing_fill_cols:
            merged[col] = merged[col].replace(r'^\s*$', np.nan, regex=True)
        merged[existing_fill_cols] = merged.groupby(group_keys, dropna=False)[existing_fill_cols].transform(
            lambda series: series.ffill().bfill()
        )
        return merged