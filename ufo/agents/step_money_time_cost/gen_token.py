import os
import json
import tiktoken
import base64
import math
from PIL import Image
import io


def calculate_image_tokens(base64_string):
    """
    根据图片的base64编码计算其在高细节模式下的token数。
    需要 Pillow 库。
    """
    try:
        header, encoded = base64_string.split(",", 1)
        image_data = base64.b64decode(encoded)
        image = Image.open(io.BytesIO(image_data))
        width, height = image.size

        if width > 2048 or height > 2048:
            if width > height:
                new_width = 2048
                new_height = int(2048 * height / width)
            else:
                new_height = 2048
                new_width = int(2048 * width / height)
            width, height = new_width, new_height

        if min(width, height) > 768:
            if width < height:
                new_width = 768
                new_height = int(768 * height / width)
            else:
                new_height = 768
                new_width = int(768 * width / height)
            width, height = new_width, new_height

        tiles_width = math.ceil(width / 512)
        tiles_height = math.ceil(height / 512)
        num_tiles = tiles_width * tiles_height
        return 85 + num_tiles * 170
    except Exception as e:
        print(f"    警告: 处理图片时出错: {e}. 无法计算此图片的token。返回0。")
        return 0


def num_tokens_from_messages(messages, model="gpt-4"):
    """
    返回由消息列表计算出的token数。
    现在可以处理包含文本和图片的多模态内容。
    """
    try:
        encoding = tiktoken.encoding_for_model(model)
    except KeyError:
        print("Warning: Model not found. Using cl100k_base encoding.")
        encoding = tiktoken.get_encoding("cl100k_base")

    tokens_per_message = 3
    tokens_per_name = 1
    num_tokens = 0
    for message in messages:
        num_tokens += tokens_per_message
        for key, value in message.items():
            if key != 'content':
                if value is not None:
                    num_tokens += len(encoding.encode(str(value)))
                if key == "name":
                    num_tokens += tokens_per_name
            else:
                if value is None:
                    continue
                if isinstance(value, str):
                    num_tokens += len(encoding.encode(value))
                elif isinstance(value, list):
                    for item in value:
                        if isinstance(item, dict):
                            item_type = item.get('type')
                            if item_type == 'text':
                                text_content = item.get('text', '')
                                num_tokens += len(encoding.encode(text_content))
                            elif item_type == 'image_url':
                                image_url_dict = item.get('image_url', {})
                                url = image_url_dict.get('url', '')
                                if url.startswith('data:image'):
                                    num_tokens += calculate_image_tokens(url)
    num_tokens += 3
    return num_tokens


def get_tokens_from_cost_file(file_path):
    """
    安全地从指定的cost文件中读取prompt和completion token。
    如果文件或键不存在，返回 (0, 0)。
    """
    if not os.path.exists(file_path):
        return 0, 0
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
            llm_request = data.get("llm_request", {})
            prompt_tokens = llm_request.get("prompt_tokens", 0)
            completion_tokens = llm_request.get("completion_tokens", 0)
            return prompt_tokens, completion_tokens
    except (json.JSONDecodeError, Exception) as e:
        print(f"    警告: 读取或解析文件 {file_path} 时出错: {e}. 将返回 (0, 0).")
        return 0, 0


def process_folders(main_folders):
    """
    遍历指定的文件夹，处理日志并生成token文件。
    """
    TOKEN_MODEL = "gpt-4"

    for main_folder in main_folders:
        if not os.path.isdir(main_folder):
            print(f"错误：主文件夹 '{main_folder}' 不存在，已跳过。")
            continue
        print(f"--- 正在处理主文件夹: {main_folder} ---")

        for entry in os.scandir(main_folder):
            if entry.is_dir():
                subfolder_path = entry.path
                print(f"  正在处理子文件夹: {subfolder_path}")

                # 定义所有需要用到的文件路径
                request_log_path = os.path.join(subfolder_path, 'request.log')
                money_json_path = os.path.join(subfolder_path, 'gen_case_money.json')
                # 第一个输出文件
                token_output_path = os.path.join(subfolder_path, 'gen_case_token.json')
                # --- 新增：第二个输出文件路径 ---
                summary_output_path = os.path.join(subfolder_path, 'gen_case_token_sum.json')

                document_cost_path = os.path.join(subfolder_path, 'document_cost', 'document_demo_cost.json')
                video_cost_path = os.path.join(subfolder_path, 'video_cost', 'video_demo_cost.json')

                try:
                    # 使用 os.path.isfile() 进行严格检查，防止 OSError
                    if not os.path.isfile(request_log_path):
                        print(f"    警告: 核心文件 request.log 不存在或是一个文件夹，已跳过。")
                        continue
                    if not os.path.isfile(money_json_path):
                        print(f"    警告: 核心文件 gen_case_money.json 不存在或是一个文件夹，已跳过。")
                        continue

                    # 计算 n1 (step_prompt_token)
                    total_prompt_tokens = 0
                    with open(request_log_path, 'r', encoding='utf-8') as f:
                        for line in f:
                            try:
                                log_entry = json.loads(line.strip())
                                if 'prompt' in log_entry and isinstance(log_entry['prompt'], list):
                                    total_prompt_tokens += num_tokens_from_messages(log_entry['prompt'],
                                                                                    model=TOKEN_MODEL)
                            except json.JSONDecodeError:
                                print(f"    警告: request.log 中有非JSON格式的行，已忽略。")
                    n1 = total_prompt_tokens

                    # 读取 money_cost
                    money_cost = 0
                    with open(money_json_path, 'r', encoding='utf-8') as f:
                        money_data = json.load(f)
                        gen_step_data = money_data.get("gen_step", {})
                        if isinstance(gen_step_data, dict):
                            money_cost = gen_step_data.get("money_cost", 0)
                        else:
                            money_cost = gen_step_data

                    if money_cost == 0:
                        print(f"    警告: 在 {money_json_path} 中未能读取到有效的 'money_cost'。")

                    # 计算 n2 (step_completion_token)
                    # prompt_cost = n1 * 0.002
                    # total_cost_scaled = money_cost * 1000
                    # completion_tokens_float = (total_cost_scaled - prompt_cost) / 0.008
                    # n2 = max(0, int(round(completion_tokens_float)))

                    # prompt_cost = n1 * 0.01
                    # total_cost_scaled = money_cost * 1000
                    # completion_tokens_float = (total_cost_scaled - prompt_cost) / 0.04
                    # n2 = max(0, int(round(completion_tokens_float)))

                    prompt_cost = n1 * 0.00125
                    total_cost_scaled = money_cost * 1000
                    completion_tokens_float = (total_cost_scaled - prompt_cost) / 0.01
                    n2 = max(0, int(round(completion_tokens_float)))

                    # 读取可选的token信息
                    doc_p_token, doc_c_token = get_tokens_from_cost_file(document_cost_path)
                    vid_p_token, vid_c_token = get_tokens_from_cost_file(video_cost_path)

                    # --- 步骤 1: 生成 gen_case_token.json ---
                    token_data = {
                        "step_prompt_token": n1,
                        "step_completion_token": n2,
                        "document_prompt_token": doc_p_token,
                        "document_completion_token": doc_c_token,
                        "video_prompt_token": vid_p_token,
                        "video_completion_token": vid_c_token
                    }
                    with open(token_output_path, 'w', encoding='utf-8') as f:
                        json.dump(token_data, f, indent=4)
                    print(f"    成功! 已更新或创建文件: {token_output_path}")

                    # --- 步骤 2: 基于上一步的结果进行求和计算 ---
                    document_sum = doc_p_token + doc_c_token
                    video_sum = vid_p_token + vid_c_token
                    step_sum = n1 + n2
                    total_sum = document_sum + video_sum + step_sum

                    # --- 步骤 3: 生成 gen_case_step_sum.json ---
                    summary_data = {
                        "document_token_sum": document_sum,
                        "video_token_sum": video_sum,
                        "step_token_sum": step_sum,
                        "total_token_sum": total_sum
                    }
                    with open(summary_output_path, 'w', encoding='utf-8') as f:
                        json.dump(summary_data, f, indent=4)
                    print(f"    成功! 已更新或创建汇总文件: {summary_output_path}")

                except Exception as e:
                    print(f"    在处理 {subfolder_path} 时发生未知错误: {e}，已跳过。")


# --- 使用说明 ---
if __name__ == '__main__':
    # main_folders_to_process = [
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_bing_4.1_cost_complete_double',
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_m365_4.1_cost_complete_double',
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_qabench_4.1_cost_complete_double'
    # ]



    # main_folders_to_process = [
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_o3\excel_complete_double'
    # ]

    main_folders_to_process = [
            r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\bing_completion_double',
            r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\m365_completion_double',
            r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\qabench_completion_double',
    ]

    if not main_folders_to_process:
        print("请编辑脚本，在 `main_folders_to_process` 列表中填入需要处理的文件夹路径。")
    else:
        try:
            from PIL import Image
        except ImportError:
            print("\n错误: 缺少Pillow库，图片token无法计算。")
            print("请先运行: pip install Pillow\n")
            exit()
        process_folders(main_folders_to_process)