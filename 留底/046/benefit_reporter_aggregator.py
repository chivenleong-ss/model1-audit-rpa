# -*- coding: utf-8 -*-
"""
效益表填报用的数据聚合规则。
"""

from __future__ import annotations

from project_config import MATERIAL_FIXED_SUBCATS


class BenefitDataAggregator:
    @staticmethod
    def sort_rows(rows):
        return sorted(
            rows,
            key=lambda item: (
                0 if item.get("is_v", False) else 1,
                0 if item.get("c") and str(item.get("c", "")).strip() not in ("", "nan", "None") else 1,
            ),
        )

    def aggregate_labor(self, df):
        df = df.copy()
        df['_pfx'] = df['中台单据号'].fillna('').astype(str).str[:3].str.upper()
        df['_ven'] = df['客商名称'].fillna('').astype(str).str.strip().replace('nan', '')
        df['_con'] = df['合同编码'].fillna('').astype(str).str.strip().replace('nan', '')
        rows = []

        fgd = df[df['_pfx'] == 'FGD']
        if not fgd.empty:
            rows.append({"d": "FGD", "c": None, "p": round(fgd['最终发生额'].sum(), 2), "is_v": True})

        for ven, grp in df[df['_pfx'] == 'CFK'].groupby('_ven', sort=False):
            label = f"CFK+{ven}" if ven else "CFK"
            rows.append({"d": label, "c": None, "p": round(grp['最终发生额'].sum(), 2), "is_v": True})

        others = df[~df['_pfx'].isin(['FGD', 'CFK'])].copy()
        has_vc = others[(others['_ven'] != '') | (others['_con'] != '')]
        for (ven, con), grp in has_vc.groupby(['_ven', '_con'], sort=False):
            rows.append({"d": ven or None, "c": con or None, "p": round(grp['最终发生额'].sum(), 2), "is_v": True})

        return self.sort_rows(rows)

    def aggregate_vendor(self, df):
        rows = []
        for (ven, con), grp in df.groupby([df['客商名称'].fillna(''), df['合同编码'].fillna('')], sort=False):
            rows.append({"d": ven or None, "c": con or None, "p": round(grp['最终发生额'].sum(), 2), "is_v": True})
        return self.sort_rows(rows)

    @staticmethod
    def aggregate_sub(self_df):
        rows = []
        for sub, grp in self_df.groupby(self_df['细分科目'].fillna(''), sort=False):
            rows.append({"d": sub or None, "c": None, "p": round(grp['最终发生额'].sum(), 2), "is_v": False})
        return rows

    def aggregate_material(self, df):
        excl_set = set(MATERIAL_FIXED_SUBCATS)
        df = df[~df['细分科目'].fillna('').astype(str).str.strip().isin(excl_set)].copy()
        if df.empty:
            return []

        has_contract = df[df['合同编码'].fillna('').astype(str).str.strip().replace('nan', '') != '']
        no_contract = df[~df.index.isin(has_contract.index)]
        rows = self.sort_rows(self.aggregate_vendor(has_contract) + self.aggregate_sub(no_contract))

        for row in rows:
            d_val = str(row.get('d', ''))
            if d_val.startswith('研发支出-'):
                row['d'] = '减：研发费用-材料费'
        return rows
