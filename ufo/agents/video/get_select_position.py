#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import json
import os
from pathlib import Path
from PIL import Image
import shutil


def extract_control_coordinates(response_log_path, output_path):
    """
    從 response.log 檔案中為每個步驟提取控制項座標。
    (此函數保持不變)
    """
    if not os.path.exists(response_log_path):
        raise FileNotFoundError(f"response.log file not found: {response_log_path}")

    result = {}
    try:
        with open(response_log_path, 'r', encoding='utf-8') as f:
            for line_num, line in enumerate(f, 1):
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    step = data.get('Step')
                    if step is None:
                        continue
                    control_label = data.get('ControlLabel', '')
                    if control_label and control_label.strip():
                        key = f"action_step{step}_selected_controls.png"
                    else:
                        result[f"action_step{step}_selected_controls.png"] = {}
                        key = None
                    control_log = data.get('ControlLog', [])
                    coordinates = None
                    if control_log and len(control_log) > 0:
                        first_control = control_log[0]
                        control_coordinates = first_control.get('control_coordinates', {})
                        if control_coordinates and any(control_coordinates.values()):
                            coordinates = control_coordinates
                    if key is not None:
                        if coordinates:
                            result[key] = coordinates
                        else:
                            result[key] = {}
                    else:
                        if f"action_step{step}_selected_controls.png" not in result:
                            pass
                except json.JSONDecodeError as e:
                    print(f"警告: JSON 解析失敗於行 {line_num}: {e}")
                    continue
                except Exception as e:
                    print(f"警告: 處理失敗於行 {line_num}: {e}")
                    continue
    except Exception as e:
        raise Exception(f"讀取檔案失敗: {e}")

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        # print(f"成功生成座標檔案: {output_path}") # 在批次處理中減少輸出
        return result
    except Exception as e:
        raise Exception(f"儲存檔案失敗: {e}")


def add_mouse_cursor_to_images(json_path, mouse_cursor_path, cursor_max_size=(28, 28)):
    """
    根據 JSON 檔案中的座標，將滑鼠圖示等比例縮放後添加到圖片上。
    (此函數保持不變)
    """
    json_file_path = Path(json_path)
    image_dir = json_file_path.parent

    try:
        with open(json_file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except Exception as e:
        print(f"錯誤: 無法載入 JSON 檔案 {json_file_path}: {e}")
        return

    try:
        cursor_img = Image.open(mouse_cursor_path).convert("RGBA")
        cursor_img.thumbnail(cursor_max_size, Image.Resampling.LANCZOS)
    except FileNotFoundError:
        print(f"錯誤: 找不到滑鼠圖示圖片 '{mouse_cursor_path}'。")
        return
    except Exception as e:
        print(f"錯誤: 處理滑鼠圖示時發生錯誤: {e}")
        return

    for image_filename, coords in data.items():
        if not image_filename or image_filename == "null":
            continue

        source_image_path = image_dir / image_filename
        output_filename = f"{source_image_path.stem}_mouse.png"
        output_image_path = image_dir / output_filename

        if not source_image_path.exists():
            print(f"  警告: 來源圖片不存在，跳過: {source_image_path}")
            continue

        if not isinstance(coords, dict) or not coords:
            shutil.copy(source_image_path, output_image_path)
            # print(f"  已複製 '{image_filename}' (無座標)。")
            continue

        try:
            print(coords)
            base_img = Image.open(source_image_path).convert("RGBA")
            center_x = (coords['left'] + coords['right']) // 2
            center_y = (coords['top'] + coords['bottom']) // 2

            paste_position = (center_x, center_y)
            base_img.paste(cursor_img, paste_position, cursor_img)
            base_img.save(output_image_path)
            # print(f"  已處理 '{image_filename}'。")

        except Exception as e:
            print(f"  錯誤: 處理圖片 {image_filename} 時發生錯誤: {e}")


# ★★★ 新的主流程函數 ★★★
def process_batch_directory(main_directory, mouse_cursor_path):
    """
    遞迴尋找指定目錄下所有的 response.log 檔案，並對每一個執行
    「生成 JSON」和「P上滑鼠圖示」的完整流程。

    Args:
        main_directory (str): 包含所有子資料夾的根目錄路徑。
        mouse_cursor_path (str): 滑鼠圖示圖片的路徑。
    """
    directory_path = Path(main_directory)
    if not directory_path.is_dir():
        print(f"錯誤: 找不到指定的目錄 '{main_directory}'")
        return

    cursor_path = Path(mouse_cursor_path)
    if not cursor_path.is_file():
        print(f"錯誤: 找不到滑鼠圖示檔案 '{mouse_cursor_path}'")
        return

    # 1. 使用 rglob 遞迴尋找所有 response.log 檔案
    log_files = list(directory_path.rglob('response.log'))

    if not log_files:
        print(f"在目錄 {directory_path} 中沒有找到任何 response.log 檔案。")
        return

    print(f"--- 找到 {len(log_files)} 個 response.log 檔案，開始批次處理 ---")

    # 2. 遍歷每一個找到的 log 檔案
    for i, log_file in enumerate(log_files, 1):
        print(f"\n[{i}/{len(log_files)}] 正在處理: {log_file.parent}")

        # 定義該 log 檔案對應的 JSON 輸出路徑
        json_output_path = log_file.parent / 'mouse_position.json'

        try:
            # 步驟 A: 生成 JSON 檔案
            print(f"  -> 步驟 1: 從 {log_file.name} 生成 JSON...")
            extract_control_coordinates(str(log_file), str(json_output_path))

            # 步驟 B: 根據剛生成的 JSON 檔案處理圖片
            print(f"  -> 步驟 2: 為對應圖片添加滑鼠圖示...")
            add_mouse_cursor_to_images(str(json_output_path), str(cursor_path))
            print(f"  -> 完成!")

        except Exception as e:
            print(f"!! 處理 {log_file.parent} 時發生未知錯誤，跳過此目錄。錯誤: {e}")

    print(f"\n--- 所有批次任務處理完畢 ---")


if __name__ == '__main__':
    # --- 使用者設定 ---

    # 1. 設定包含所有 log 子資料夾的大資料夾路徑
    #    程式會自動遍歷這個資料夾下的所有內容
    main_folder_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\log_result\20250716_bing_search_completed"

    main_folder_path = r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_ufo1_baseilne\qabench_completion_double_4.1_try"

    # 2. 設定您的滑鼠圖示圖片的絕對路徑
    cursor_image_path = "./data/mouse.png"

    # --- 執行批次處理 ---
    process_batch_directory(main_folder_path, cursor_image_path)