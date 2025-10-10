import os
import json


# --- 辅助函数：用于计算成本和步骤 ---

def get_step_num(subfolder_path: str) -> int:
    """
    从指定的子文件夹中获取操作步骤的数量。

    Args:
        subfolder_path (str): 目标子文件夹的完整路径。

    Returns:
        int: 'steps' 字典中的项目数。如果文件不存在则返回 0，如果出错则返回 -1。
    """
    json_path = os.path.join(subfolder_path, 'document', 'document_demo_step.json')

    if not os.path.exists(json_path):
        return 0  # 文件不存在，视为 0 个步骤

    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            # 确保 'steps' 存在并且是一个字典
            if 'steps' in data and isinstance(data.get('steps'), dict):
                return len(data['steps'])
            else:
                return 0  # 'steps' 键不存在或格式不正确
    except (json.JSONDecodeError, KeyError) as e:
        print(f"  ⚠️  警告: 处理 {json_path} 时出错. 错误: {e}")
        return -1  # 返回 -1 表示处理时发生错误


def calculate_money_cost(subfolder_path: str) -> dict:
    """
    计算单个子文件夹的金钱成本。

    Args:
        subfolder_path (str): 目标子文件夹的路径。

    Returns:
        dict: 包含各项及总金钱成本的字典。
    """
    result = {
        "gen_step": 0.0,
        "gen_video": 0.0,
        "gen_document": 0.0,
        "total_money": 0.0
    }

    # 1. 处理 response.log
    gen_step_total_cost = 0.0
    response_log_path = os.path.join(subfolder_path, 'response.log')
    try:
        with open(response_log_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not line.strip(): continue
                try:
                    data = json.loads(line)
                    gen_step_total_cost += data.get('Cost', 0.0)
                except json.JSONDecodeError:
                    pass  # 静默处理格式错误的行
        result["gen_step"] = gen_step_total_cost
    except FileNotFoundError:
        pass  # 文件不存在则跳过

    # 2. 处理 video_cost/video_demo_cost.json
    video_cost_path = os.path.join(subfolder_path, 'video_cost', 'video_demo_cost.json')
    try:
        with open(video_cost_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            result["gen_video"] = data.get('llm_request', {}).get('cost', 0.0)
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        pass

    # 3. 处理 document_cost/document_demo_cost.json
    doc_cost_path = os.path.join(subfolder_path, 'document_cost', 'document_demo_cost.json')
    try:
        with open(doc_cost_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            result["gen_document"] = data.get('llm_request', {}).get('cost', 0.0)
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        pass

    # 4. 计算总费用
    result["total_money"] = result["gen_step"] + result["gen_video"] + result["gen_document"]
    return result


def calculate_time_cost(subfolder_path: str) -> dict:
    """
    计算单个子文件夹的时间成本。

    Args:
        subfolder_path (str): 目标子文件夹的路径。

    Returns:
        dict: 包含各项及总时间成本的字典。
    """
    result = {
        "gen_step": 0.0,
        "gen_video": 0.0,
        "gen_document": 0.0,
        "total_time": 0.0
    }

    # 1. 处理 response.log
    gen_step_total_diff = 0.0
    response_log_path = os.path.join(subfolder_path, 'response.log')
    try:
        with open(response_log_path, 'r', encoding='utf-8') as f:
            for line in f:
                if not line.strip(): continue
                try:
                    data = json.loads(line)
                    total_time_cost = data.get('total_time_cost', 0.0)
                    response_time = data.get('get_response_time_true', data.get('get_response_time', 0.0))
                    get_response_api_time = data.get('time_cost', {}).get('get_response', 0.0)
                    diff = total_time_cost + response_time - get_response_api_time
                    gen_step_total_diff += diff
                except json.JSONDecodeError:
                    pass
        result["gen_step"] = gen_step_total_diff
    except FileNotFoundError:
        pass

    # 2. 处理 video_cost/video_demo_cost.json
    video_cost_path = os.path.join(subfolder_path, 'video_cost', 'video_demo_cost.json')
    try:
        with open(video_cost_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            time_taken = data.get('llm_request', {}).get('time_taken_seconds', 0.0)
            gen_time = data.get('gen_document_time', 0.0)
            result["gen_video"] = time_taken + gen_time
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        pass

    # 3. 处理 document_cost/document_demo_cost.json
    doc_cost_path = os.path.join(subfolder_path, 'document_cost', 'document_demo_cost.json')
    try:
        with open(doc_cost_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            time_taken = data.get('llm_request', {}).get('time_taken_seconds', 0.0)
            gen_time = data.get('gen_document_time', 0.0)
            result["gen_document"] = time_taken + gen_time
    except (FileNotFoundError, KeyError, json.JSONDecodeError):
        pass

    # 4. 计算总时间
    result["total_time"] = result["gen_step"] + result["gen_video"] + result["gen_document"]
    return result


# --- 主处理函数 ---

def process_judged_data(input_path: str, output_path: str, source_map: dict):
    """
    读取JSONL文件，处理judge=true的记录，并添加新字段后写入新文件。

    Args:
        input_path (str): 输入的 JSONL 文件路径。
        output_path (str): 输出的 JSONL 文件路径。
        source_map (dict): 'source' 到其根文件夹路径的映射。
    """
    print(f"🚀 开始处理文件: {input_path}")
    processed_count = 0
    skipped_count = 0

    try:
        with open(input_path, 'r', encoding='utf-8') as infile, \
                open(output_path, 'w', encoding='utf-8') as outfile:

            for line in infile:
                if not line.strip():
                    continue

                try:
                    data = json.loads(line)
                except json.JSONDecodeError:
                    print(f"  ⚠️  警告: 跳过一个格式错误的JSON行。")
                    continue

                # 核心逻辑：只处理 judge 为 true 的记录
                if data.get("judge") is not True:
                    skipped_count += 1
                    continue

                source = data.get("source")
                file_name = data.get("file_name")

                if source in source_map and file_name:
                    base_folder = source_map[source]
                    subfolder_path = os.path.join(base_folder, file_name)

                    if os.path.isdir(subfolder_path):
                        # 1. 添加 path 字段
                        data['path'] = subfolder_path

                        # 2. 添加 step_num 字段
                        data['step_num'] = get_step_num(subfolder_path)

                        # 3. 添加 money_cost 字段
                        data['money_cost'] = calculate_money_cost(subfolder_path)

                        # 4. 添加 time_cost 字段
                        data['time_cost'] = calculate_time_cost(subfolder_path)

                        processed_count += 1
                    else:
                        print(f"  ℹ️  提示: 找不到对应的文件夹，跳过: {subfolder_path}")
                        skipped_count += 1
                        continue  # 如果找不到文件夹，跳过此记录
                else:
                    skipped_count += 1
                    continue  # 如果 source 或 file_name 无效，跳过

                # 将更新后的字典写回新文件
                outfile.write(json.dumps(data, ensure_ascii=False) + '\n')

    except FileNotFoundError:
        print(f"❌ 错误: 输入文件未找到 '{input_path}'")
        return
    except Exception as e:
        print(f"❌ 处理过程中发生意外错误: {e}")
        return

    print("\n✅ 处理完成!")
    print(f"  - {processed_count} 条记录已处理并写入到: {output_path}")
    print(f"  - {skipped_count} 条记录被跳过 (judge不为true或信息不完整)。")


# --- 主程序入口 ---
if __name__ == "__main__":
    # --- 配置区域 ---

    # 1. 定义 'source' 到根文件夹的映射
    # 请确保这些路径是正确的
    SOURCE_TO_FOLDER_MAP = {
        "bing_search": r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_bing_4.1_cost_complete_double',
        "m365": r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_m365_4.1_cost_complete_double',
        "qabench": r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_qabench_4.1_cost_complete_double'
    }

    # SOURCE_TO_FOLDER_MAP = {
    #     "bing_search":  r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\bing_completion_double',
    #     "m365":   r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\m365_completion_double',
    #     "qabench":r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\qabench_completion_double',
    # }



    # 2. 定义输入和输出文件的路径
    # 假设你的输入文件名为 'output_with_judge.jsonl' (上一步脚本的输出)
    input_jsonl_file = 'output_with_judge.jsonl'
    output_jsonl_file = 'final_processed_data.jsonl'

    # 3. 执行主处理函数
    process_judged_data(input_jsonl_file, output_jsonl_file, SOURCE_TO_FOLDER_MAP)
