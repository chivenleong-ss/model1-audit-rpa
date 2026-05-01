# -*- coding: utf-8 -*-
"""
批量上传分组工具。

只负责识别上传文件的利润中心和文件类型，不触碰审计、填报、美化核心流程。
"""

from __future__ import annotations

import os
import re
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


PROFIT_CENTER_ALIASES = {
    "利润中心",
    "利润中心编号",
    "利润中心编码",
    "利润中心代码",
    "利润中心号",
    "利润中心ID",
    "利润中心id",
    "PRCTR",
    "Profit Center",
    "ProfitCenter",
    "profit_center",
}

# --- 新增：用于识别项目名称表头的关键字 ---
PROFIT_CENTER_NAME_ALIASES = {
    "利润中心文本描述",
    "利润中心名称",
    "项目名称",
    "利润中心描述",
    "描述",
}

PROFIT_CENTER_CODE_RE = re.compile(r"\b[A-Z]\d{6,}\b|\b\d{6,}\b", re.IGNORECASE)
STRICT_PROFIT_CENTER_CODE_RE = re.compile(r"\b[A-Z]\d{6,}\b", re.IGNORECASE)

VOUCHER_NAME_KEYS = ("凭证", "综合查询", "主表")
DETAIL_NAME_KEYS = ("明细", "辅助")
INVOICE_NAME_KEYS = ("发票", "收票", "台账", "台帐", "已认证")


@dataclass
class BatchFile:
    path: str
    filename: str
    profit_center: str
    file_type: str
    error: str = ""


@dataclass
class ProfitCenterGroup:
    profit_center: str
    pc_name: str = ""  # --- 新增：利润中心名称描述 ---
    voucher_files: list[BatchFile] = field(default_factory=list)
    detail_files: list[BatchFile] = field(default_factory=list)
    invoice_files: list[BatchFile] = field(default_factory=list)
    other_files: list[BatchFile] = field(default_factory=list)

    @property
    def ready(self) -> bool:
        return bool(self.voucher_files and self.detail_files)

    @property
    def error(self) -> str:
        missing = []
        if not self.voucher_files:
            missing.append("凭证主表")
        if not self.detail_files:
            missing.append("辅助明细账")
        if missing:
            return "缺少：" + "、".join(missing)
        return ""

    def to_payload(self) -> dict:
        return {
            "profit_center": self.profit_center,
            "pc_name": self.pc_name,  # --- 新增：推送给前端的 JSON 字段 ---
            "ready": self.ready,
            "error": self.error,
            "voucher_files": [item.filename for item in self.voucher_files],
            "detail_files": [item.filename for item in self.detail_files],
            "invoice_files": [item.filename for item in self.invoice_files],
            "other_files": [item.filename for item in self.other_files],
        }


def _clean_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.lower() in {"nan", "none", "nat"}:
        return ""
    return text


def _normalize_alias(value) -> str:
    return re.sub(r"[\s:_：/\\（）()\[\]【】-]+", "", _clean_text(value)).lower()


def _is_profit_center_alias(value) -> bool:
    normalized = _normalize_alias(value)
    if not normalized:
        return False
    return any(normalized == _normalize_alias(alias) for alias in PROFIT_CENTER_ALIASES)


# --- 新增：判断是否是名称表头 ---
def _is_profit_center_name_alias(value) -> bool:
    normalized = _normalize_alias(value)
    if not normalized:
        return False
    return any(normalized == _normalize_alias(alias) for alias in PROFIT_CENTER_NAME_ALIASES)


def _read_preview(file_path: str) -> pd.DataFrame:
    path = str(file_path)
    suffix = Path(path).suffix.lower()
    try:
        if suffix == ".csv":
            try:
                return pd.read_csv(path, header=None, nrows=10, dtype=str, low_memory=False)
            except UnicodeDecodeError:
                return pd.read_csv(path, header=None, nrows=10, dtype=str, encoding="gbk", low_memory=False)
        if suffix in {".xlsx", ".xls", ".xlsm"}:
            return pd.read_excel(path, header=None, nrows=10, dtype=str)
    except Exception as exc:
        raise RuntimeError(f"读取文件前10行失败：{os.path.basename(path)}，原因：{exc}") from exc
    raise RuntimeError(f"不支持的文件格式：{os.path.basename(path)}")


def _extract_code_from_text(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    match = PROFIT_CENTER_CODE_RE.search(text)
    if not match:
        return ""
    return match.group(0).upper()


def _extract_strict_code_from_text(value: str) -> str:
    text = _clean_text(value)
    if not text:
        return ""
    match = STRICT_PROFIT_CENTER_CODE_RE.search(text)
    if not match:
        return ""
    return match.group(0).upper()


# --- 修改：升级版提取函数，同时提取 Code 和 Name ---
def extract_profit_center_info(file_path: str) -> tuple[str, str]:
    """
    读取文件前10行，提取利润中心唯一编码 和 文本描述（如果有）。
    返回: (pc_code, pc_name)
    """
    preview = _read_preview(file_path)
    if preview.empty:
        raise RuntimeError(f"文件为空，无法提取利润中心：{os.path.basename(file_path)}")

    found_codes: list[str] = []
    found_names: list[str] = []
    row_count, col_count = preview.shape

    for row_idx in range(row_count):
        for col_idx in range(col_count):
            cell = _clean_text(preview.iat[row_idx, col_idx])
            
            # 1. 抓取利润中心代码
            if _is_profit_center_alias(cell):
                for next_row in range(row_idx + 1, row_count):
                    code = _extract_code_from_text(preview.iat[next_row, col_idx])
                    if code:
                        found_codes.append(code)

                if col_idx + 1 < col_count:
                    code = _extract_code_from_text(preview.iat[row_idx, col_idx + 1])
                    if code:
                        found_codes.append(code)

            # 2. 抓取利润中心名称 (新增逻辑)
            if _is_profit_center_name_alias(cell):
                for next_row in range(row_idx + 1, row_count):
                    name_val = _clean_text(preview.iat[next_row, col_idx])
                    # 确保抓到的名字不是纯粹的一串数字代码
                    if name_val and not PROFIT_CENTER_CODE_RE.fullmatch(name_val):
                        found_names.append(name_val)
                        break

    if not found_codes:
        for value in preview.to_numpy().ravel():
            code = _extract_strict_code_from_text(value)
            if code:
                found_codes.append(code)

    unique_codes = []
    for code in found_codes:
        if code not in unique_codes:
            unique_codes.append(code)

    if not unique_codes:
        raise RuntimeError(f"前10行未识别到利润中心编码：{os.path.basename(file_path)}")
    if len(unique_codes) > 1:
        raise RuntimeError(
            f"前10行识别到多个利润中心编码：{os.path.basename(file_path)}（{', '.join(unique_codes)}）"
        )
        
    final_code = unique_codes[0]
    final_name = found_names[0] if found_names else ""
    return final_code, final_name


def detect_file_type(file_path: str) -> str:
    filename = os.path.basename(file_path)
    if any(key in filename for key in INVOICE_NAME_KEYS):
        return "invoice"
    if any(key in filename for key in DETAIL_NAME_KEYS):
        return "detail"
    if any(key in filename for key in VOUCHER_NAME_KEYS):
        return "voucher"

    preview = _read_preview(file_path)
    text = "|".join(_clean_text(value) for value in preview.to_numpy().ravel())
    if any(key in text for key in INVOICE_NAME_KEYS) or "发票号码" in text:
        return "invoice"
    if "辅助" in text or "利润中心文本描述" in text:
        return "detail"
    if "凭证编号" in text and "总账科目" in text:
        return "voucher"
    return "other"


def build_profit_center_groups(file_paths: list[str]) -> tuple[list[ProfitCenterGroup], list[BatchFile]]:
    groups: dict[str, ProfitCenterGroup] = {}
    rejected: list[BatchFile] = []

    for file_path in file_paths:
        filename = os.path.basename(file_path)
        try:
            # --- 修改：解包获取 code 和 name ---
            profit_center, pc_name = extract_profit_center_info(file_path)
            file_type = detect_file_type(file_path)
            item = BatchFile(
                path=file_path,
                filename=filename,
                profit_center=profit_center,
                file_type=file_type,
            )
            
            # --- 修改：如果是第一次创建 group，带上 pc_name ---
            if profit_center not in groups:
                groups[profit_center] = ProfitCenterGroup(profit_center=profit_center, pc_name=pc_name)
            else:
                # 如果这个项目已经建了组，但是之前的表没扫出名字，现在的表扫出了名字，就补充进去
                if pc_name and not groups[profit_center].pc_name:
                    groups[profit_center].pc_name = pc_name
                    
            group = groups[profit_center]
            
            if file_type == "voucher":
                group.voucher_files.append(item)
            elif file_type == "detail":
                group.detail_files.append(item)
            elif file_type == "invoice":
                group.invoice_files.append(item)
            else:
                group.other_files.append(item)
        except Exception as exc:
            rejected.append(
                BatchFile(
                    path=file_path,
                    filename=filename,
                    profit_center="未识别",
                    file_type="other",
                    error=str(exc),
                )
            )

    return sorted(groups.values(), key=lambda group: group.profit_center), rejected