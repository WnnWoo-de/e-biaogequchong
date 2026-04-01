import pandas as pd
import os
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def merge_and_deduplicate(input_files, output_file, subset_cols=['姓名', '学号'], template_file=None):
    """
    合并多个Excel文件并去重，保留第一个文件的格式。
    
    :param input_files: 文件路径列表
    :param output_file: 输出文件名
    :param subset_cols: 去重依据的列名列表
    :param template_file: 模板文件（用于保留格式），默认使用列表中的第一个文件
    """
    print(f"--- 🚀 开始处理表格 (共 {len(input_files)} 个文件) ---")
    
    all_dfs = []
    
    # 1. 循环读取所有文件
    for file in input_files:
        if not os.path.exists(file):
            print(f"⚠️ 警告: 找不到文件 {file}，已跳过")
            continue
            
        print(f"📖 正在读取: {file}...")
        try:
            df = pd.read_excel(file)
            # 数据预处理：去除列名和内容中的空格，防止匹配失败
            df.columns = [str(c).strip() for c in df.columns]
            for col in subset_cols:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()
            
            all_dfs.append(df)
        except Exception as e:
            print(f"❌ 读取 {file} 失败: {e}")

    if not all_dfs:
        print("❌ 错误: 没有有效的数据可以合并")
        return

    # 2. 合并数据
    print("🔄 正在合并数据...")
    combined_df = pd.concat(all_dfs, ignore_index=True)
    original_count = len(combined_df)

    # 3. 执行去重
    # 检查去重列是否存在
    missing_cols = [c for c in subset_cols if c not in combined_df.columns]
    if missing_cols:
        print(f"❌ 错误: 表格中缺少去重必要的列 {missing_cols}")
        print(f"💡 当前表头包含: {list(combined_df.columns)}")
        return

    print(f"✂️ 正在根据 {subset_cols} 进行去重...")
    clean_df = combined_df.drop_duplicates(subset=subset_cols, keep='first')
    
    final_count = len(clean_df)
    print(f"📊 统计结果:")
    print(f"   - 原始总行数: {original_count}")
    print(f"   - 去重后行数: {final_count}")
    print(f"   - 剔除重复项: {original_count - final_count} 条")

    # 4. 写入文件并保留格式
    # 默认使用第一个有效文件作为样式模板
    if not template_file:
        template_file = input_files[0]

    print(f"🎨 正在应用模板格式 ({template_file})...")
    try:
        wb = load_workbook(template_file)
        ws = wb.active
        
        # 清空原有数据（保留表头）
        if ws.max_row > 1:
            ws.delete_rows(2, amount=ws.max_row)

        # 写入清洗后的数据
        rows = dataframe_to_rows(clean_df, index=False, header=False)
        for r_idx, row in enumerate(rows, 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx + 1, column=c_idx, value=value)

        wb.save(output_file)
        print(f"✅ 处理完成！结果已保存至: {output_file}")
        
    except Exception as e:
        print(f"⚠️ 格式保存失败（可能是模板不匹配），尝试直接保存为普通Excel...")
        clean_df.to_excel(output_file, index=False)
        print(f"✅ 已保存(无格式版本): {output_file}")

# ==========================================
# 🔧 配置区域
# ==========================================
if __name__ == "__main__":
    # 1. 待合并的文件列表 (请确保这些文件在当前目录下)
    FILES_TO_MERGE = [
        '第二次周例活动问卷星收集.xlsx',
        '英语角第一次周例活动问卷星.xlsx',
        # '可以继续添加更多文件名.xlsx'
    ]
    
    # 2. 输出文件名
    OUTPUT_NAME = '合并去重后的最终名单.xlsx'
    
    # 3. 判定重复的依据（这两列都一样才算重复）
    DEDUP_COLUMNS = ['姓名', '学号']

    merge_and_deduplicate(FILES_TO_MERGE, OUTPUT_NAME, DEDUP_COLUMNS)
