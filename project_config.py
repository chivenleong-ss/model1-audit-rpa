# 项目共享配置
MATERIAL_FIXED_SUBCATS = [
    '材料库存', '材料调入', '材料调出', '其他应收+CFK','其他','内部存款/内部往来+CFK',
    '材料调出+XSD', '材料调出+SQD', '废旧物资消耗', '废旧物资消耗（SQD）'
]

# 安全生产费/研发费用统一收口规则：
# 1. M 列金额以中间汇总表为准，不再在写表阶段按 B/C 回加。
# 2. 仅以下细分科目允许生成“减：”抵减行。
# 3. SQD 单据、ZGD-* 专项行保留单列展示，不并入通用减项。
DRAFT_DEDUCTION_EXCLUDE_DOC_PREFIXES = ('SQD',)
DRAFT_DEDUCTION_EXCLUDE_SUBCAT_PREFIXES = ('ZGD-',)

SAFETY_DEDUCTION_RULES = {
    '安全生产-人工费': ('(一)人工费', '减：安全生产-人工费'),
    '安全生产-分包工程': ('(二)分包工程', '减：安全生产-分包工程'),
    '安全生产-机械使用': ('(四)机械租赁费', '减：安全生产-机械使用'),
}

RESEARCH_DEDUCTION_RULES = {
    '研发费用-机械租赁': ('(四)机械租赁费', '减：研发费用-机械租赁'),
}

RESEARCH_OFFSET_MAP = {
    '材料费': ('(三)材料费', '研发支出-材料费'),
}

ZGD_PREFIX = 'ZGD'
ZGD_SUBCAT_PREFIX = 'ZGD-'

VENDOR_SHORTNAME_MAP = {}
