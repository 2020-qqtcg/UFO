import pandas as pd
import os
import shutil
import json
import re # argparse 已移除

def clean_filename(name):
    """移除或替换不适合文件名的字符"""
    name = str(name)
    # 移除非字母、数字、下划线、短横线、点的字符
    name = re.sub(r'[^\w\-\.]', '_', name)
    # 避免多个连续下划线
    name = re.sub(r'_+', '_', name)
    # 避免开头或结尾是下划线
    name = name.strip('_')
    if not name: # 如果清理后为空，给一个默认名
        name = "untitled"
    return name

def process_tasks(master_excel_path, task_sheets_dir, output_base_dir):
    """
    读取主Excel文件，处理每一行数据，复制工作表，并生成JSON。

    Args:
        master_excel_path (str): 主任务Excel文件的路径。
        task_sheets_dir (str): 包含原始工作表Excel文件的目录路径。
        output_base_dir (str): 保存处理后文件和JSON的基础目录路径。
    """
    if not os.path.exists(master_excel_path):
        print(f"错误：主Excel文件未找到: {master_excel_path}")
        return

    if not os.path.isdir(task_sheets_dir):
        print(f"错误：task_sheets目录未找到: {task_sheets_dir}")
        return

    # 创建输出子目录
    prepared_excel_output_dir = os.path.join(output_base_dir, "prepared_excel_files")
    json_output_dir = os.path.join(output_base_dir, "json_files")
    agent_save_base_dir = os.path.join(output_base_dir, "agent_task_outputs") # 代理保存文件的基础路径

    os.makedirs(prepared_excel_output_dir, exist_ok=True)
    os.makedirs(json_output_dir, exist_ok=True)
    os.makedirs(agent_save_base_dir, exist_ok=True) # 确保这个目录也存在

    try:
        df = pd.read_excel(master_excel_path)
    except FileNotFoundError:
        print(f"错误: 无法读取主Excel文件 {master_excel_path}")
        return
    except Exception as e:
        print(f"读取主Excel文件时发生错误: {e}")
        return

    # 检查必需的列是否存在
    required_columns = ['Sheet Name', 'No.', 'Instructions']
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        print(f"错误：主Excel文件中缺少以下必需列: {', '.join(missing_columns)}")
        return

    print(f"开始处理主Excel文件: {master_excel_path}")
    processed_count = 0
    skipped_count = 0

    for index, row in df.iterrows():
        try:
            # 确保从Excel读取的原始数据转换为字符串并去除首尾空格
            sheet_name_original = str(row['Sheet Name']).strip()
            task_no_original = str(row['No.']).strip()
            instructions_text = str(row['Instructions']).strip()
            # context_text = str(row.get('Context', '')).strip() # 如果需要Context

            if not sheet_name_original or not task_no_original:
                print(f"警告: 第 {index + 2} 行缺少 'Sheet Name' 或 'No.'，跳过此行。")
                skipped_count += 1
                continue

            # 清理工作表名和任务编号，用于文件名和路径
            sheet_name_cleaned = clean_filename(sheet_name_original)
            task_no_cleaned = clean_filename(task_no_original)

            # 1. 定位并复制Excel文件
            # 假设task_sheets下的文件直接是Sheet Name.xlsx (原始Sheet Name)
            source_excel_filename = f"{sheet_name_original}.xlsx"
            source_excel_path = os.path.join(task_sheets_dir, source_excel_filename)

            if not os.path.exists(source_excel_path):
                print(f"警告: 在 {task_sheets_dir} 中未找到源Excel文件 '{source_excel_filename}' (对应原始任务No. {task_no_original})，跳过此任务。")
                skipped_count += 1
                continue

            # 准备的目标Excel文件名和路径 (使用清理后的名称)
            prepared_excel_filename = f"{task_no_cleaned}_{sheet_name_cleaned}.xlsx"
            prepared_excel_full_path = os.path.join(prepared_excel_output_dir, prepared_excel_filename)

            # 复制文件
            shutil.copy2(source_excel_path, prepared_excel_full_path)
            print(f"  已复制: '{source_excel_filename}' 到 '{prepared_excel_full_path}'")

            # 2. 准备 'save_as' 路径 (使用清理后的名称)
            agent_specific_save_dir = os.path.join(agent_save_base_dir, f"{task_no_cleaned}_{sheet_name_cleaned}")
            os.makedirs(agent_specific_save_dir, exist_ok=True)
            save_as_path_for_json = os.path.join(agent_specific_save_dir, f"{task_no_cleaned}_{sheet_name_cleaned}.xlsx")

            # 3. 构建 output_data 字典
            output_data = {
                "task": instructions_text,
                "object": os.path.normpath(prepared_excel_full_path),
                "close": "True",
                "save_as": os.path.normpath(save_as_path_for_json)
            }

            # 4. 存储为单独的JSON文件 (使用清理后的名称)
            json_filename = f"{task_no_cleaned}_{sheet_name_cleaned}.json"
            json_full_path = os.path.join(json_output_dir, json_filename)

            with open(json_full_path, 'w', encoding='utf-8') as f:
                json.dump(output_data, f, ensure_ascii=False, indent=4)
            print(f"  已生成JSON: '{json_full_path}'")
            processed_count += 1

        except KeyError as e:
            print(f"警告: 第 {index + 2} 行缺少必需的列名 {e}，跳过此行。")
            skipped_count += 1
        except Exception as e:
            print(f"处理第 {index + 2} 行 (原始任务No. {row.get('No.', '未知')}) 时发生意外错误: {e}")
            skipped_count += 1

    print(f"\n处理完成。成功处理 {processed_count} 个任务，跳过 {skipped_count} 个任务。")
    print(f"准备好的Excel文件位于: {prepared_excel_output_dir}")
    print(f"生成的JSON文件位于: {json_output_dir}")
    print(f"代理预期保存文件的基础目录是: {agent_save_base_dir}")


if __name__ == "__main__":
    # --- 在这里直接修改为你需要的路径 ---
    master_excel_path = r"D:\code\SheetCopilot\dataset\dataset.xlsx"  # 例如: "C:/Users/YourUser/Documents/tasks_main.xlsx"
    task_sheets_dir = r"D:\code\SheetCopilot\dataset\task_sheets"         # 例如: "C:/Users/YourUser/Documents/excel_sources"
    output_base_dir = r"D:\code\UFO\benchmark\sheetcoplit\tasks" # 例如: "C:/Users/YourUser/Documents/processed_tasks"
    # --- 修改结束 ---

    # 确保你提供的路径是有效的，并且脚本有权限读写这些位置
    # 例如，对于Windows用户，路径可能像 "C:\\path\\to\\file.xlsx" 或使用正斜杠 "C:/path/to/file.xlsx"
    # 对于Linux/macOS用户, 路径像 "/home/user/path/to/file.xlsx"

    print("--- 配置的路径 ---")
    print(f"主Excel文件: {master_excel_path}")
    print(f"任务工作表目录: {task_sheets_dir}")
    print(f"输出基础目录: {output_base_dir}")
    print("--------------------")

    if not all([master_excel_path != "path/to/your/master_tasks.xlsx",
                task_sheets_dir != "path/to/your/task_sheets",
                output_base_dir != "path/to/your/output_data_processed"]):
        print("\n警告：请在脚本中更新 master_excel_path, task_sheets_dir, 和 output_base_dir 的占位符路径！")
    else:
        process_tasks(master_excel_path, task_sheets_dir, output_base_dir)