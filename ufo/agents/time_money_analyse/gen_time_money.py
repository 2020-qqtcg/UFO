import os
import json



def process_costs_in_directory(main_directory: str):
    """
    遍歷主目錄下的所有子資料夾，處理日誌文件並計算費用成本。

    Args:
        main_directory (str): 包含多個子資料夾的主目錄路徑。
    """
    if not os.path.isdir(main_directory):
        print(f"❌ 錯誤：找不到目錄 '{main_directory}'")
        return

    # 遍歷主目錄中的每個項目
    for subfolder_name in os.listdir(main_directory):
        subfolder_path = os.path.join(main_directory, subfolder_name)

        # 僅處理資料夾
        if not os.path.isdir(subfolder_path):
            continue

        print(f"📂 正在處理資料夾：{subfolder_path}")

        # 初始化結果字典
        result = {
            "gen_step": 0.0,
            "gen_video": 0.0,
            "gen_document": 0.0,
            "total_money": 0.0
        }

        # 1. 處理 response.log
        gen_step_total_cost = 0.0
        response_log_path = os.path.join(subfolder_path, 'response.log')
        try:
            with open(response_log_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if not line.strip():
                        continue
                    try:
                        data = json.loads(line)
                        # 累加每個 dict 中的 "Cost"
                        gen_step_total_cost += data.get('Cost', 0.0)
                    except json.JSONDecodeError:
                        print(f"  ⚠️ 警告：跳過 {response_log_path} 中的一個格式錯誤的 JSON 行。")
            result["gen_step"] = gen_step_total_cost
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'response.log'。")
        except Exception as e:
            print(f"  ❌ 處理 'response.log' 時發生錯誤：{e}")

        # 2. 處理 video_cost/video_demo_cost.json
        video_cost_path = os.path.join(subfolder_path, 'video_cost', 'video_demo_cost.json')
        try:
            with open(video_cost_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 從 llm_request 中獲取 cost
                video_cost = data.get('llm_request', {}).get('cost', 0.0)
                result["gen_video"] = video_cost
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'video_cost/video_demo_cost.json'。")
        except (KeyError, json.JSONDecodeError) as e:
            print(f"  ❌ 處理 'video_demo_cost.json' 時發生錯誤：{e}")

        # 3. 處理 document_cost/document_demo_cost.json
        doc_cost_path = os.path.join(subfolder_path, 'document_cost', 'document_demo_cost.json')
        try:
            with open(doc_cost_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 從 llm_request 中獲取 cost
                doc_cost = data.get('llm_request', {}).get('cost', 0.0)
                result["gen_document"] = doc_cost
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'document_cost/document_demo_cost.json'。")
        except (KeyError, json.JSONDecodeError) as e:
            print(f"  ❌ 處理 'document_demo_cost.json' 時發生錯誤：{e}")

        # 4. 計算總費用
        # 注意：根據您的第4點，這裡假設 "gen_video_time" 是 "gen_video" 的筆誤
        result["total_money"] = result["gen_step"] + result["gen_video"] + result["gen_document"]

        # 5. 將結果寫入 gen_case_money.json
        output_path = os.path.join(subfolder_path, 'gen_case_money.json')
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=4, ensure_ascii=False)
            print(f"  ✅ 結果已成功寫入：{output_path}\n")
        except IOError as e:
            print(f"  ❌ 無法寫入結果檔案 '{output_path}'。錯誤：{e}\n")


def process_subfolder_logs(main_directory: str):
    """
    遍歷主目錄下的所有子資料夾，處理日誌文件並計算時間成本。

    Args:
        main_directory (str): 包含多個子資料夾的主目錄路徑。
    """
    if not os.path.isdir(main_directory):
        print(f"❌ 錯誤：找不到目錄 '{main_directory}'")
        return

    # 遍歷主目錄中的每個項目
    for subfolder_name in os.listdir(main_directory):
        subfolder_path = os.path.join(main_directory, subfolder_name)

        # 僅處理資料夾
        if not os.path.isdir(subfolder_path):
            continue

        print(f"📂 正在處理資料夾：{subfolder_path}")

        # 初始化結果字典
        result = {
            "gen_step": 0.0,
            "gen_video": 0.0,
            "gen_document": 0.0,
            "total_time": 0.0
        }

        # 1. 處理 response.log
        gen_step_total_diff = 0.0
        response_log_path = os.path.join(subfolder_path, 'response.log')
        try:
            with open(response_log_path, 'r', encoding='utf-8') as f:
                for line in f:
                    if not line.strip():
                        continue
                    try:
                        data = json.loads(line)
                        total_time_cost = data.get('total_time_cost', 0.0)

                        # 獲取 "get_response_time_true" 或 "get_response_time"
                        response_time = data.get('get_response_time_true', data.get('get_response_time', 0.0))

                        # 獲取 "get_response" 時間
                        get_response_api_time = data.get('time_cost', {}).get('get_response', 0.0)

                        # 計算差值並累加
                        diff = total_time_cost + response_time - get_response_api_time
                        gen_step_total_diff += diff
                    except json.JSONDecodeError:
                        print(f"  ⚠️ 警告：跳過 {response_log_path} 中的一個格式錯誤的 JSON 行。")
            result["gen_step"] = gen_step_total_diff
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'response.log'。")
        except Exception as e:
            print(f"  ❌ 處理 'response.log' 時發生錯誤：{e}")

        # 2. 處理 video_cost/video_demo_cost.json
        video_cost_path = os.path.join(subfolder_path, 'video_cost', 'video_demo_cost.json')
        try:
            with open(video_cost_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                time_taken = data.get('llm_request', {}).get('time_taken_seconds', 0.0)
                gen_time = data.get('gen_document_time', 0.0)
                result["gen_video"] = time_taken + gen_time
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'video_cost/video_demo_cost.json'。")
        except (KeyError, json.JSONDecodeError) as e:
            print(f"  ❌ 處理 'video_demo_cost.json' 時發生錯誤：{e}")

        # 3. 處理 document_cost/document_demo_cost.json
        doc_cost_path = os.path.join(subfolder_path, 'document_cost', 'document_demo_cost.json')
        try:
            with open(doc_cost_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                time_taken = data.get('llm_request', {}).get('time_taken_seconds', 0.0)
                gen_time = data.get('gen_document_time', 0.0)
                result["gen_document"] = time_taken + gen_time
        except FileNotFoundError:
            print(f"  ℹ️ 提示：在 {subfolder_path} 中找不到 'document_cost/document_demo_cost.json'。")
        except (KeyError, json.JSONDecodeError) as e:
            print(f"  ❌ 處理 'document_demo_cost.json' 時發生錯誤：{e}")

        # 4. 計算總時間
        result["total_time"] = result["gen_step"] + result["gen_video"] + result["gen_document"]

        # 5. 將結果寫入 gen_case_time.json
        output_path = os.path.join(subfolder_path, 'gen_case_time.json')
        try:
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(result, f, indent=4, ensure_ascii=False)
            print(f"  ✅ 結果已成功寫入：{output_path}\n")
        except IOError as e:
            print(f"  ❌ 無法寫入結果檔案 '{output_path}'。錯誤：{e}\n")


def aggregate_folder_data(main_directories: list, output_file: str):
    """
    遍歷多個主目錄下的所有子文件夾，讀取指定的json文件，
    並將結果聚合成一個jsonl文件。

    Args:
        main_directories (list): 包含多個主目錄路徑的列表。
        output_file (str): 輸出的 jsonl 文件路徑。
    """
    print(f"🚀 開始聚合數據，結果將保存至：{output_file}")
    records_written = 0

    # 使用 'w' 模式打開輸出文件，確保每次運行都創建新文件
    with open(output_file, 'w', encoding='utf-8') as f_out:
        # 遍歷每一個主目錄
        for main_dir in main_directories:
            if not os.path.isdir(main_dir):
                print(f"⚠️  警告：找不到主目錄，已跳過：{main_dir}")
                continue

            print(f"\n📂 正在掃描目錄：{main_dir}")

            # 遍歷主目錄下的所有子文件夾
            for subfolder_name in os.listdir(main_dir):
                subfolder_path = os.path.join(main_dir, subfolder_name)

                if not os.path.isdir(subfolder_path):
                    continue

                # 定義要讀取的目標文件路徑
                money_file_path = os.path.join(subfolder_path, 'gen_case_money.json')
                time_file_path = os.path.join(subfolder_path, 'gen_case_time.json')

                money_data = None
                time_data = None

                # 讀取 gen_case_money.json 的內容
                if os.path.exists(money_file_path):
                    try:
                        with open(money_file_path, 'r', encoding='utf-8') as f_money:
                            money_data = json.load(f_money)
                    except json.JSONDecodeError:
                        print(f"  ❌ 錯誤：無法解析JSON文件：{money_file_path}")
                    except Exception as e:
                        print(f"  ❌ 讀取時發生未知錯誤 {money_file_path}: {e}")

                # 讀取 gen_case_time.json 的內容
                if os.path.exists(time_file_path):
                    try:
                        with open(time_file_path, 'r', encoding='utf-8') as f_time:
                            time_data = json.load(f_time)
                    except json.JSONDecodeError:
                        print(f"  ❌ 錯誤：無法解析JSON文件：{time_file_path}")
                    except Exception as e:
                        print(f"  ❌ 讀取時發生未知錯誤 {time_file_path}: {e}")

                # 只有當至少一個文件成功讀取時才寫入記錄
                if money_data is not None or time_data is not None:
                    # 構建輸出字典
                    output_record = {
                        "file_name": subfolder_name,
                        "money_cost": money_data,
                        "time_cost": time_data
                    }

                    # 將字典轉換為JSON字符串並寫入文件，每條記錄占一行
                    f_out.write(json.dumps(output_record, ensure_ascii=False) + '\n')
                    records_written += 1

    print(f"\n✅ 聚合完成！總共寫入了 {records_written} 條記錄到 {output_file}。")


if __name__ == '__main__':
    # --- 請在此處配置您的文件夾路徑 ---
    # 將 'path/to/...' 替換為您實際的大文件夾路徑
    # directories_to_scan = [
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_bing_4.1_cost_complete_double',
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_m365_4.1_cost_complete_double',
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_qabench_4.1_cost_complete_double',
    # ]

    # directories_to_scan = [
    #     r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_o3\excel_complete_double'
    # ]

    directories_to_scan = [
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\bing_completion_double',
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\m365_completion_double',
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_gpt5\qabench_completion_double',
    ]

    for folder in directories_to_scan:
        process_costs_in_directory(folder)
        process_subfolder_logs(folder)
    # --- 輸出文件名配置 ---
    output_jsonl_file = './time_money_result/costs_and_times.jsonl'

    # --- 執行聚合函數 ---
    aggregate_folder_data(directories_to_scan, output_jsonl_file)