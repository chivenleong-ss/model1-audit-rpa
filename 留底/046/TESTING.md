# 测试说明

## 1. 端到端冒烟测试

已提供脚本：

`smoke_test.py`

默认会：

1. 从 `3_已处理数据备份` 自动挑选最近的一对凭证/明细样本
2. 如存在发票台账，也会自动带上
3. 调用 `AuditPipeline` 跑完整主链
4. 在 `_smoke_output` 下检查关键输出是否生成

运行方式：

```powershell
& 'C:\Users\sasa\AppData\Local\Python\pythoncore-3.14-64\python.exe' smoke_test.py
```

如需手工指定输入：

```powershell
& 'C:\Users\sasa\AppData\Local\Python\pythoncore-3.14-64\python.exe' smoke_test.py `
  --voucher '你的凭证文件.xlsx' `
  --detail '你的明细文件.xlsx' `
  --invoice '你的发票文件.xlsx'
```

## 2. 最小回归测试

已提供脚本：

`regression_test.py`

默认会：

1. 不重复跑完整主链
2. 自动定位 `_smoke_output` 下最近一次成功结果目录
3. 校验关键输出文件是否齐全
4. 校验 `4_合并表.xlsx` 的表头、`SUBTOTAL` 合计行与首行数据
5. 校验 `5_中间汇总表_效益审核数据源.xlsx` 的关键列与分类数据
6. 校验 `自动填报完成_效益审核表.xlsx` / `最终完美交付版_效益审核表.xlsx` 的标题、关键行、关键公式

运行方式：

```powershell
& 'C:\Users\sasa\AppData\Local\Python\pythoncore-3.14-64\python.exe' regression_test.py
```

如需指定某次结果目录：

```powershell
& 'C:\Users\sasa\AppData\Local\Python\pythoncore-3.14-64\python.exe' regression_test.py `
  --result-dir '_smoke_output\审计核对结果_20260425_001713'
```

如需先刷新一次最新产物，再校验：

```powershell
& 'C:\Users\sasa\AppData\Local\Python\pythoncore-3.14-64\python.exe' regression_test.py --run-smoke
```

## 3. 输出位置

默认输出到：

`_smoke_output`

关键校验文件：

- `4_合并表.xlsx`
- `5_中间汇总表_效益审核数据源.xlsx`
- `最终完美交付版_效益审核表.xlsx`
  或 `交付版_效益审核表.xlsx`
  或 `自动填报完成_效益审核表.xlsx`

## 4. 时长说明

真实 Excel 样本较大时，完整冒烟测试通常需要几分钟。

默认回归测试不会重跑主链，只读取现有产物，因此通常明显快于冒烟测试。

本次已实际验证通过，最近一次成功输出目录为：

`_smoke_output\审计核对结果_20260425_001713`
