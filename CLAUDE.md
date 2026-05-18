# Excel 合并去重工具

作者: WnnW

## 项目概述

处理报名表、签到表等 Excel 名单数据，支持多文件合并和按指定列去重。

## 文件结构

- `remove.py` — **主脚本**，统一的合并去重工具，支持 CLI 参数
- `Finnish.py` — 合并去重脚本（保留 Excel 模板格式版），适合需要保留原始表格样式的场景
- `deduplicate_excel.py` — 单文件去重工具，自动识别表头行，适合处理问卷星导出的带标题表格
- `find_duplicates.py` — 两表对比工具，筛选出两个表格中同时存在的人员
- `fire_read.py` — 简单的文本文件读取工具

## 常用命令

```bash
# 合并多个文件并去重（默认按姓名+学号）
python remove.py -i 名单1.xlsx 名单2.xlsx -o 最终名单.xlsx -c 姓名 学号

# 保留最后一次出现的数据
python remove.py -i 名单1.xlsx 名单2.xlsx --keep last

# 单文件去重（自动识别问卷星表头）
python deduplicate_excel.py -i 问卷星导出.xlsx -o 去重结果.xlsx

# 两表对比找重复人员
python find_duplicates.py
```

## 依赖

```bash
pip install pandas openpyxl
```

## 注意事项

- 参与合并的表格应保持相同表头
- 去重字段名必须与表格列名一致
- `deduplicate_excel.py` 可自动跳过问卷星导出的标题行
