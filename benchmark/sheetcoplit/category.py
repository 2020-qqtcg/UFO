import os
import shutil
import pandas as pd

# 路径配置
excel_path = r'C:\Users\v-yuhangxie\UFO\benchmark\sheetcoplit\SheetCopilot\dataset\dataset.xlsx'  # 替换为你的Excel文件路径
source_dir = r'C:\Users\v-yuhangxie\UFO\logs\20250530_copilot'
target_base_dir = r'C:\Users\v-yuhangxie\UFO\logs\20250530_copilot_category'

# 读取Excel中的所有Sheet
df = pd.read_excel(excel_path, sheet_name=None)

# 遍历每个Sheet和对应数据
for sheet_name, sheet_df in df.items():
    for idx, row in sheet_df.iterrows():
        try:
            print()
            no = str(row['No.'])
            raw_category = str(row['Categories']).strip()

            # 跳过无效分类
            if raw_category.lower() == 'nan' or raw_category == '':
                continue

            # 拆分类别
            categories = [cat.strip() for cat in raw_category.split(',') if cat.strip()]
            folder_name = f"{no}_{sheet_name}"
            src_folder_path = os.path.join(source_dir, folder_name)

            if not os.path.exists(src_folder_path):
                print(f"❌ 找不到源文件夹: {src_folder_path}")
                continue

            for category in categories:
                dst_folder_path = os.path.join(target_base_dir, category, folder_name)

                if os.path.exists(dst_folder_path):
                    print(f"⚠️ 已存在目标文件夹: {dst_folder_path}，跳过复制")
                    continue

                os.makedirs(os.path.dirname(dst_folder_path), exist_ok=True)
                shutil.copytree(src_folder_path, dst_folder_path)
                print(f"✅ 已复制: {src_folder_path} -> {dst_folder_path}")

        except Exception as e:
            print(f"⚠️ 出错 (第 {idx+1} 行): {e}")
