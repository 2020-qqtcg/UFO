import os
import json
import shutil  # 1. 引入 shutil 模組

# --- 設定路徑 ---
# 來源路徑：包含所有待檢查子資料夾的根目錄
# base_path = r"C:\Users\v-yuhangxie\UFO_ssb_0708\logs\20250725_qabench_4.1_cost"
# base_path = r"C:\Users\v-yuhangxie\UFO_ssb_0708\logs\20250726_qabench_4.1_cost_without_api"
# base_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_o3\excel"
# base_path = r"C:\Users\v-yuhangxie\UFO_ssb_0708\logs\20250828_qabench_4.1_ufo1"
base_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_operator"
# 目標路徑：將完成的資料夾複製到這裡
# destination_base_path = r"C:\Users\v-yuhangxie\UFO_ssb_0708\logs\20250725_qabench_4.1_cost_complete"
# destination_base_path = r"C:\Users\v-yuhangxie\UFO_ssb_0708\logs\20250726_qabench_4.1_cost_without_api_complete"
# destination_base_path =r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_o3\excel_complete"
# destination_base_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_ufo1_baseilne\qabench_completion"
destination_base_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_operator_complete"

# --- 初始化 ---
completed_folders = []

# 2. 確保目標資料夾存在，如果不存在就建立
print(f"目標資料夾為: {destination_base_path}")
os.makedirs(destination_base_path, exist_ok=True)
print("-" * 30)

# --- 主邏輯：遍歷、檢查與複製 ---
for subdir in os.listdir(base_path):
    subdir_path = os.path.join(base_path, subdir)
    log_path = os.path.join(subdir_path, "evaluation.log")

    # 只處理資料夾，且目標資料夾不與來源資料夾相同
    if os.path.isdir(subdir_path) and os.path.normpath(subdir_path) != os.path.normpath(destination_base_path):

        # 檢查 evaluation.log 是否存在
        if not os.path.exists(log_path):
            continue

        try:
            with open(log_path, "r", encoding="utf-8") as f:
                for line in f:
                    try:
                        data = json.loads(line.strip())
                        # 檢查 "complete" 欄位是否為 "yes"
                        if data.get("complete") == "yes":
                            print(f"🔍 找到完成的任務: '{subdir}'")

                            # --- 3. 執行複製操作 ---
                            destination_path = os.path.join(destination_base_path, subdir)

                            try:
                                # 使用 shutil.copytree 複製整個資料夾
                                # dirs_exist_ok=True 允許目標資料夾已存在 (Python 3.8+)，方便重複執行
                                shutil.copytree(subdir_path, destination_path, dirs_exist_ok=True)
                                print(f"✅ 成功複製 '{subdir}' 至 '{destination_path}'")
                            except Exception as copy_error:
                                print(f"❌ 複製資料夾 '{subdir}' 時發生錯誤: {copy_error}")

                            completed_folders.append(subdir)
                            break  # 找到符合條件的行後，就處理下一個資料夾

                    except json.JSONDecodeError:
                        continue  # JSON 解析失敗，跳過此行
        except Exception as e:
            print(f"⚠️  跳過 '{subdir}'，因讀取或解析失敗：{e}")
        finally:
            print("-" * 30)

# --- 輸出總結 JSON 檔案 ---
if completed_folders:
    output_file = os.path.join(base_path, "completed_folders.json")
    with open(output_file, "w", encoding="utf-8") as f:
        json.dump(completed_folders, f, indent=2, ensure_ascii=False)

    print(f"🎉 處理完成！")
    print(f"共找到並複製了 {len(completed_folders)} 個已完成的資料夾。")
    print(f"詳細列表已保存至: {output_file}")
else:
    print("🤷‍♂️ 未找到任何標記為 'complete' 的資料夾。")