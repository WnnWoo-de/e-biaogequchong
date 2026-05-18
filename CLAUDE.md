# Excel 合并去重工具

作者: WnnW

## 项目概述

处理报名表、签到表等 Excel 名单数据，支持多文件合并和按指定列去重。

## 文件结构

- `remove.py` — **主脚本**，统一的合并去重工具，支持 CLI 参数、模板格式保留
- `deduplicate_excel.py` — 单文件去重工具，自动识别表头行，适合处理问卷星导出的带标题表格

## 常用命令

```bash
# 合并多个文件并去重（默认按姓名+学号）
python remove.py -i 名单1.xlsx 名单2.xlsx -o 最终名单.xlsx -c 姓名 学号

# 保留最后一次出现的数据
python remove.py -i 名单1.xlsx 名单2.xlsx --keep last

# 保留原始 Excel 格式（使用模板）
python remove.py -i 名单1.xlsx 名单2.xlsx -o 最终名单.xlsx -t 模板文件.xlsx

# 同时导出重复人员名单
python remove.py -i 名单1.xlsx 名单2.xlsx -o 最终名单.xlsx -d 重复人员.xlsx

# 单文件去重（自动识别问卷星表头）
python deduplicate_excel.py -i 问卷星导出.xlsx -o 去重结果.xlsx

```

## 依赖

```bash
pip install pandas openpyxl
```

## 注意事项

- 参与合并的表格应保持相同表头
- 去重字段名必须与表格列名一致
- `deduplicate_excel.py` 可自动跳过问卷星导出的标题行
