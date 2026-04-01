import pandas as pd
import os


def remove_duplicates_and_save(file_path, output_file='去重后的最终名单.xlsx'):
    # --- 配置区域 ---
    # 这里填写你Excel中实际的列名，用于判断“谁是同一个人”
    col_id = '学号'
    col_name = '姓名'
    # ----------------

    print(f"正在读取文件: {file_path} ...")

    if not os.path.exists(file_path):
        print(f"❌ 错误: 找不到文件 '{file_path}'，请检查路径。")
        return

    try:
        # 读取 Excel 文件
        df = pd.read_excel(file_path)

        # 检查列名是否存在
        if col_id not in df.columns or col_name not in df.columns:
            print(f"❌ 错误: 表格中找不到列名 '{col_id}' 或 '{col_name}'。")
            print(f"当前表格的所有列名: {list(df.columns)}")
            return

        # 记录去重前的行数
        original_count = len(df)
        print(f"📊 原始数据共有: {original_count} 条")

        # 【核心代码】删除重复项
        # subset: 依据学号和姓名判断
        # keep='first': 保留第一次出现的记录，删除后面重复的
        # (如果想保留最后一次出现的，把 'first' 改为 'last')
        df_cleaned = df.drop_duplicates(subset=[col_id, col_name], keep='first')

        # 记录去重后的行数
        cleaned_count = len(df_cleaned)
        removed_count = original_count - cleaned_count

        if removed_count == 0:
            print("✅ 未发现重复项，无需清理。")
            # 即使没有重复，也建议保存一份或者提示用户
        else:
            print(f"✂️ 删除了 {removed_count} 条重复记录。")
            print(f"✅ 剩余有效数据: {cleaned_count} 条")

            # 保存结果到新的 Excel
            df_cleaned.to_excel(output_file, index=False)
            print("-" * 30)
            print(f"📄 处理完成！结果已保存为: {output_file}")
            print("-" * 30)

    except Exception as e:
        print(f"❌ 发生未知错误: {e}")


# --- 执行部分 ---
if __name__ == "__main__":
    # 替换为你的文件名
    input_filename = '参与人员加分表.xlsx'

    remove_duplicates_and_save(input_filename)