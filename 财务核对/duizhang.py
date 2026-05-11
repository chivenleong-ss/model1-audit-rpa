import pandas as pd
import numpy as np
import os

print("🚀 启动【审计员逻辑复刻版】终极对账...")
current_dir = os.path.dirname(os.path.abspath(__file__))
file_hebing = os.path.join(current_dir, '4_合并表.xlsx')
file_zhongjian = os.path.join(current_dir, '5_中间汇总表_效益审核数据源.xlsx')

df_hebing = pd.read_excel(file_hebing, dtype={'中台单据号': str})
df_zhongjian = pd.read_excel(file_zhongjian, dtype={'中台单据号': str})

# 核心：完全按照 logic_auditor.py 的清洗与识别方法
df_hebing['借方'] = pd.to_numeric(df_hebing['借方本位币金额'], errors='coerce').fillna(0)
df_zhongjian['最终发生额'] = pd.to_numeric(df_zhongjian['最终发生额'], errors='coerce').fillna(0)

def clean_bill_no(val):
    if pd.isna(val) or str(val).strip().lower() in ['', '0', '0.0', 'nan', 'null']:
        return '【无单据号/纯手工凭证】'
    return str(val).strip()

df_hebing['单据号'] = df_hebing['中台单据号'].apply(clean_bill_no)
df_zhongjian['单据号'] = df_zhongjian['中台单据号'].apply(clean_bill_no)

# 审计员 (LogicAuditor) 的粗暴关键词字典
AUDIT_RULES = {
    '人工费': {
        'etl_cat': '(一)人工费',
        'sap_keyword': r'合同履约成本\\工程施工成本\\直接人工费'
    },
    '分包工程': {
        'etl_cat': '(二)分包工程',
        'sap_keyword': r'合同履约成本\\工程施工成本\\分包工程支出'
    },
    '材料费': {
        'etl_cat': '(三)材料费',
        'sap_keyword': r'合同履约成本\\工程施工成本\\直接材料费'
    }
}

all_diffs = []

for name, rule in AUDIT_RULES.items():
    print(f"\n{'='*40}\n正在按照 logic_auditor 规则核对: 【{name}】\n{'='*40}")
    
    # 1. 运动员 (ETL表) 的金额
    sub_zhongjian = df_zhongjian[df_zhongjian['成本_财务大类'] == rule['etl_cat']]
    grp_zhongjian = sub_zhongjian.groupby('单据号')['最终发生额'].sum().reset_index()
    total_etl = grp_zhongjian['最终发生额'].sum()
    
    # 2. 裁判员 (Auditor) 的瞎眼抓取法（只看包含关键词的借方，不看贷方！）
    # 注意这里一定要用 regex=True 并且处理好反斜杠转义
    mask_sap = df_hebing['总账科目长文本'].str.contains(rule['sap_keyword'], na=False, regex=True)
    sub_sap = df_hebing[mask_sap]
    grp_sap = sub_sap.groupby('单据号')['借方'].sum().reset_index()
    total_sap = grp_sap['借方'].sum()
    
    print(f" -> [效益表列报]:  {total_etl:,.2f} 元")
    print(f" -> [合并表审计数]: {total_sap:,.2f} 元")
    
    # 合并对比
    comp = pd.merge(grp_sap, grp_zhongjian, on='单据号', how='outer').fillna(0)
    # Auditor的代码报错逻辑：效益表(M列) - 合并表(抓取值)
    comp['差额'] = comp['最终发生额'] - comp['借方']
    
    diff_df = comp[abs(comp['差额']) > 0.01].sort_values('差额', ascending=False)
    
    if not diff_df.empty:
        total_diff = diff_df['差额'].sum()
        print(f" [!] 锁定总差额: {total_diff:,.2f} 元")
        
        diff_df.rename(columns={'借方': 'Auditor系统抓取额(仅看借方)', '最终发生额': '效益表填报额'}, inplace=True)
        diff_df.insert(0, '排查科目', name)
        
        # 附带单据性质
        tags = sub_zhongjian.groupby('单据号')['细分科目'].apply(lambda x: ' / '.join(set(x.dropna()))).reset_index()
        diff_df = pd.merge(diff_df, tags, on='单据号', how='left')
        
        all_diffs.append(diff_df)

if all_diffs:
    final_df = pd.concat(all_diffs, ignore_index=True)
    out_file = os.path.join(current_dir, 'LogicAuditor_终极漏洞追踪表.xlsx')
    final_df.to_excel(out_file, index=False)
    print(f"\n🎉 完美复盘！导致 {diff_df['差额'].sum():,.2f} 差额的全部底细已生成，请看：{out_file}")