import argparse
import os
from typing import List

import pandas as pd


def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """清理列名空格，避免列名不一致导致找不到字段。"""
    copied = df.copy()
    copied.columns = [str(col).strip() for col in copied.columns]
    return copied


def normalize_dedup_columns(df: pd.DataFrame, dedup_columns: List[str]) -> pd.DataFrame:
    """仅清理去重字段（转字符串并去空格），避免误匹配。"""
    copied = df.copy()
    for col in dedup_columns:
        copied[col] = copied[col].astype(str).str.strip()
    return copied


def validate_columns(df: pd.DataFrame, dedup_columns: List[str], filename: str) -> bool:
    """检查去重列是否完整存在。"""
    missing = [col for col in dedup_columns if col not in df.columns]
    if missing:
        print(f"❌ 文件 '{filename}' 缺少必要列: {missing}")
        print(f"   当前列名: {list(df.columns)}")
        return False
    return True


def merge_and_deduplicate(
    input_files: List[str],
    output_file: str,
    dedup_columns: List[str],
    keep: str = "first",
) -> None:
    """
    合并多个 Excel 文件并按指定列去重。
    """
    if not input_files:
        print("❌ 未提供任何输入文件。")
        return

    valid_dfs = []
    total_read_rows = 0

    for path in input_files:
        if not os.path.exists(path):
            print(f"⚠️ 找不到文件，已跳过: {path}")
            continue

        try:
            df = pd.read_excel(path)
            df = normalize_columns(df)
            if not validate_columns(df, dedup_columns, path):
                return
            df = normalize_dedup_columns(df, dedup_columns)
            valid_dfs.append(df)
            total_read_rows += len(df)
            print(f"📖 已读取: {path} ({len(df)} 行)")
        except Exception as exc:
            print(f"❌ 读取失败: {path}，错误: {exc}")
            return

    if not valid_dfs:
        print("❌ 没有可用数据，处理结束。")
        return

    merged_df = pd.concat(valid_dfs, ignore_index=True)
    before_count = len(merged_df)
    deduped_df = merged_df.drop_duplicates(subset=dedup_columns, keep=keep)
    after_count = len(deduped_df)

    deduped_df.to_excel(output_file, index=False)

    print("\n✅ 处理完成")
    print(f"   输入文件数: {len(valid_dfs)}")
    print(f"   累计读取行数: {total_read_rows}")
    print(f"   合并后行数: {before_count}")
    print(f"   去重后行数: {after_count}")
    print(f"   删除重复数: {before_count - after_count}")
    print(f"   输出文件: {output_file}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Excel 合并去重工具")
    parser.add_argument(
        "-i",
        "--inputs",
        nargs="+",
        required=False,
        default=["参与人员加分表.xlsx"],
        help="待处理 Excel 文件（可传多个）",
    )
    parser.add_argument(
        "-o",
        "--output",
        default="合并去重后的最终名单.xlsx",
        help="输出文件名",
    )
    parser.add_argument(
        "-c",
        "--columns",
        nargs="+",
        default=["姓名", "学号"],
        help="用于判重的列名（可多个）",
    )
    parser.add_argument(
        "--keep",
        choices=["first", "last"],
        default="first",
        help="重复项保留策略：first 或 last",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = parse_args()
    merge_and_deduplicate(args.inputs, args.output, args.columns, args.keep)