# Excel 合并去重工具

作者: WnnW

这个项目用于处理报名表、签到表等 Excel 名单数据，支持**多文件合并**和**按指定列去重**。

现在统一使用 `remove.py`，一个脚本完成读取、合并、清洗和去重。

## 功能说明

- 支持一次输入多个 `.xlsx` 文件并合并
- 支持按多个字段去重（默认：`姓名 学号`）
- 自动清理判重字段中的前后空格
- 支持重复保留策略：`first` 或 `last`
- 输出去重后的新 Excel 文件

## 环境依赖

```bash
pip install pandas openpyxl
```

## 快速开始

### 1) 默认方式（单文件）

```bash
python remove.py
```

默认会读取 `参与人员加分表.xlsx`，输出 `合并去重后的最终名单.xlsx`。

### 2) 多文件合并去重

```bash
python remove.py -i 名单1.xlsx 名单2.xlsx 名单3.xlsx -o 最终名单.xlsx -c 姓名 学号
```

### 3) 保留最后一次出现的数据

```bash
python remove.py -i 名单1.xlsx 名单2.xlsx --keep last
```

## 参数说明

- `-i, --inputs`：输入文件列表（可多个）
- `-o, --output`：输出文件名
- `-c, --columns`：去重依据列（可多个）
- `--keep`：重复保留策略（`first` / `last`）

## 示例

```bash
python remove.py -i 第二次周例活动问卷星收集.xlsx 英语角第一次周例活动问卷星.xlsx -o 合并去重后的最终名单.xlsx -c 姓名 学号
```

## 注意事项

- 参与合并的表格应尽量保持相同表头
- 去重字段名必须与表格列名一致
- 建议处理前备份原始文件
