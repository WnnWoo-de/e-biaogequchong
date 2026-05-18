import argparse
import os
import re
import sys
from typing import List

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")


def normalize_column_name(column_name: object) -> str:
    """统一表头名称，兼容"1、姓名"这类问卷星导出格式。"""
    normalized = str(column_name).strip()
    normalized = re.sub(r"^\d+\s*[、.．)\]】）]?\s*", "", normalized)
    return normalized.strip()


def normalize_cell_text(value: object) -> str:
    """清理普通文本字段，避免空值被转成 nan 参与去重。"""
    if pd.isna(value):
        return ""
    return str(value).replace("　", " ").strip()


def normalize_student_id(value: object) -> str:
    """统一学号格式，处理 25215150414.0 或科学计数法等情况。"""
    text = normalize_cell_text(value).replace(" ", "")
    if text.lower() in {"nan", "none", "null"}:
        return ""
    if re.fullmatch(r"\d+\.0+", text):
        return text.split(".")[0]
    if re.fullmatch(r"\d+(?:\.\d+)?[eE][+-]?\d+", text):
        try:
            return format(float(text), ".0f")
        except ValueError:
            return text
    return re.sub(r"\.0+$", "", text)


def normalize_dedup_columns(df: pd.DataFrame, dedup_columns: List[str]) -> pd.DataFrame:
    """清理去重字段，学号走专用逻辑，其他字段清理空格和全角空格。"""
    copied = df.copy()
    for col in dedup_columns:
        if col == "学号":
            copied[col] = copied[col].apply(normalize_student_id)
        else:
            copied[col] = copied[col].apply(normalize_cell_text)
    return copied


def find_header_row(input_file: str, dedup_columns: List[str]) -> int:
    """自动寻找真正的表头行，例如第一行是大标题、第二行才是姓名/学号表头。"""
    preview_df = pd.read_excel(input_file, header=None, dtype=str, nrows=20)
    normalized_dedup_columns = {normalize_column_name(col) for col in dedup_columns}

    for row_index, row in preview_df.iterrows():
        row_values = {normalize_column_name(value) for value in row.dropna().tolist()}
        if normalized_dedup_columns.issubset(row_values):
            return int(row_index)

    return 0


def validate_columns(df: pd.DataFrame, dedup_columns: List[str], filename: str) -> bool:
    """检查去重列是否完整存在。"""
    missing = [col for col in dedup_columns if col not in df.columns]
    if missing:
        print(f"[错误] 文件 '{filename}' 缺少必要列: {missing}")
        print(f"[提示] 当前列名: {list(df.columns)}")
        return False
    return True


def save_dataframe_with_template(df: pd.DataFrame, output_file: str, template_file: str, description: str) -> None:
    """使用模板文件导出 DataFrame，保留原始 Excel 格式（字体、列宽等）。"""
    print(f"[导出] 正在导出{description} ({output_file})...")
    try:
        wb = load_workbook(template_file)
        ws = wb.active

        # 清空原有数据（保留表头）
        if ws.max_row > 1:
            ws.delete_rows(2, amount=ws.max_row - 1)

        # 从第 2 行开始写入数据，保留模板中的表头
        rows = dataframe_to_rows(df, index=False, header=False)
        for r_idx, row in enumerate(rows, start=2):
            for c_idx, value in enumerate(row, start=1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        wb.save(output_file)
        print(f"[完成] {description}已保存至: {output_file}")

    except Exception as e:
        print(f"[警告] {description}格式保存失败（{e}），尝试直接保存为普通 Excel...")
        df.to_excel(output_file, index=False)
        print(f"[完成] {description}已保存(无格式版本): {output_file}")


def remove_duplicates_from_excel(
    input_file: str,
    output_file: str,
    dedup_columns: List[str],
    duplicate_output_file: str = None,
    template_file: str = None,
) -> None:
    """遍历一个 Excel 表格，删除后面重复出现的数据，保留前面第一次出现的数据。"""
    if not os.path.exists(input_file):
        print(f"[错误] 找不到输入文件: {input_file}")
        return

    try:
        print(f"[读取] 正在读取: {input_file}")
        header_row = find_header_row(input_file, dedup_columns)
        if header_row > 0:
            print(f"[识别] 检测到真正表头在第 {header_row + 1} 行，已自动跳过前面的标题行")
        df = pd.read_excel(input_file, header=header_row, dtype=str)
    except Exception as exc:
        print(f"[错误] 读取失败: {exc}")
        return

    df.columns = [normalize_column_name(col) for col in df.columns]
    if not validate_columns(df, dedup_columns, input_file):
        return

    df = normalize_dedup_columns(df, dedup_columns)

    # 去掉去重字段全为空的行
    present_cols = [col for col in dedup_columns if col in df.columns]
    if present_cols:
        before_drop = len(df)
        df = df[~df[present_cols].eq("").any(axis=1)].copy()
        dropped = before_drop - len(df)
        if dropped:
            print(f"   - 已跳过 {dropped} 行空数据")

    before_count = len(df)

    # 提取重复人员（保留首次出现的那份）
    duplicate_people_df = None
    if duplicate_output_file:
        duplicate_people_df = df[df.duplicated(subset=dedup_columns, keep=False)].copy()
        duplicate_people_df = duplicate_people_df.drop_duplicates(subset=dedup_columns, keep="first")

    duplicate_rows = df[df.duplicated(subset=dedup_columns, keep="first")]
    deduped_df = df.drop_duplicates(subset=dedup_columns, keep="first")
    removed_count = before_count - len(deduped_df)

    # 导出去重结果
    tpl = template_file if template_file else input_file
    save_dataframe_with_template(deduped_df, output_file, tpl, "去重后的表格")

    print(f"\n[统计] 表格去重完成")
    print(f"   - 原始行数: {before_count}")
    print(f"   - 去重后行数: {len(deduped_df)}")
    print(f"   - 删除重复行数: {removed_count}")

    # 导出重复人员名单
    if duplicate_output_file:
        if duplicate_people_df is not None and not duplicate_people_df.empty:
            save_dataframe_with_template(duplicate_people_df, duplicate_output_file, tpl, "重复名单")
            print(f"   - 重复人员: {len(duplicate_people_df)} 人")
        else:
            print("[提示] 未发现重复人员。")

    if removed_count:
        print(f"\n[重复数据] 以下是被删除的后续重复数据:")
        print(duplicate_rows[dedup_columns].to_string(index=False))
    else:
        print(f"\n[提示] 没有发现重复数据。")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="单个 Excel 表格去重工具 — 作者: WnnW")
    parser.add_argument(
        "-i",
        "--input",
        default="2026上半年周例活动加分名单.xlsx",
        help="需要去重的 Excel 文件",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="去重后的表格.xlsx",
        help="去重后的输出文件名",
    )
    parser.add_argument(
        "-c",
        "--columns",
        nargs="+",
        default=["姓名", "学号"],
        help="判断重复的列名，可传多个；默认按 姓名 学号 判断",
    )
    parser.add_argument(
        "-d",
        "--duplicates",
        default=None,
        help="重复人员名单输出文件名（不指定则不单独导出）",
    )
    parser.add_argument(
        "-t",
        "--template",
        default=None,
        help="Excel 模板文件，用于保留原始格式（默认使用输入文件自身）",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    remove_duplicates_from_excel(args.input, args.output, args.columns, args.duplicates, args.template)
