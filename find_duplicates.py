import pandas as pd
import os


def find_common_names(file1, file2, output_file):
    """
    筛选出同时存在于两个 Excel 表格中的名单。
    逻辑：只有【姓名】和【学号】都一致，才算重复。
    """
    print("--- 正在启动重复筛选工具 ---")

    # 1. 检查文件
    if not os.path.exists(file1) or not os.path.exists(file2):
        print(f"❌ 错误：找不到文件，请检查路径：\n  {file1}\n  {file2}")
        return

    try:
        # 2. 读取数据
        print("正在读取表格...")
        df1 = pd.read_excel(file1)
        df2 = pd.read_excel(file2)

        # ==========================================
        # 核心配置：用来判断“是否重复”的关键列
        # 建议同时使用 '姓名' 和 '学号'，确保万无一失
        # ==========================================
        check_columns = ['姓名', '学号']

        # 3. 检查列名是否存在
        for col in check_columns:
            if col not in df1.columns or col not in df2.columns:
                print(f"❌ 错误：表格中缺少列名 '{col}'。")
                print(f"  表1列名: {list(df1.columns)}")
                print(f"  表2列名: {list(df2.columns)}")
                return

        # 4. 数据清洗 (防止因为看不见的空格导致匹配失败)
        print("正在比对数据...")
        # 将用于对比的列转换为字符串并去除空格
        for col in check_columns:
            df1[col] = df1[col].astype(str).str.strip()
            df2[col] = df2[col].astype(str).str.strip()

        # 5. 执行筛选 (Inner Join)
        # 逻辑：取两个表的“交集”。how='inner' 表示只保留两个表都有的数据
        duplicates_df = pd.merge(df1, df2, on=check_columns, how='inner', suffixes=('', '_表2'))

        # 6. 整理结果
        # merge 操作可能会产生重复的辅助列（比如 '班级' 和 '班级_表2'）
        # 这里我们只保留表1里的原始列，让结果更干净
        final_columns = df1.columns
        result_df = duplicates_df[final_columns]

        # 去重（防止如果表2里同一个名字出现了两次，导致结果里也显示两次）
        result_df = result_df.drop_duplicates(subset=check_columns)

        if result_df.empty:
            print("✅ 结果：两个表格没有发现重复人员。")
        else:
            print(f"✅ 发现 {len(result_df)} 个重复人员！")

            # 保存结果
            result_df.to_excel(output_file, index=False)
            print(f"   结果已保存至: {output_file}")
            print("   (保存格式保留了表1的完整信息，如学院、班级等)")

    except Exception as e:
        print(f"❌ 发生未知错误: {e}")


# ==========================================
# 👇 请在此处修改文件名 👇
# ==========================================

FILE_A = '名单1.xlsx'  # 这个文件将被用作“格式模板”
FILE_B = '名单2.xlsx'

# 结果保存的文件名
OUTPUT_FILE = '筛选出的重复名单.xlsx'

# ==========================================
# 运行程序
# ==========================================
if __name__ == "__main__":
    find_common_names(FILE_A, FILE_B, OUTPUT_FILE)