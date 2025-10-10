import json
import os
import base64  # ★★★ 新增導入
import mimetypes  # ★★★ 新增導入


def _encode_image_to_base64(image_path: str) -> str:
    """
    將圖片檔案編碼為 Base64 Data URI。

    Args:
        image_path (str): 圖片檔案的路徑。

    Returns:
        str: 格式化後的 Base64 Data URI 字串，如果失敗則返回空字串。
    """
    if not os.path.exists(image_path):
        print(f"  警告: 找不到圖片 '{image_path}'，將在HTML中忽略此圖片。")
        return ""

    # 猜測圖片的 MIME 類型 (例如 'image/png')
    mime_type, _ = mimetypes.guess_type(image_path)
    if mime_type is None:
        mime_type = "application/octet-stream"  # 如果無法確定類型，使用一個通用的二進位流類型

    try:
        # 以二進位模式讀取圖片檔案
        with open(image_path, "rb") as image_file:
            # 編碼為 Base64 並解碼為 utf-8 字串
            encoded_string = base64.b64encode(image_file.read()).decode("utf-8")

        # 返回標準的 Data URI 格式
        return f"data:{mime_type};base64,{encoded_string}"
    except Exception as e:
        print(f"  警告: 讀取或編碼圖片 '{image_path}' 時出錯: {e}")
        return ""


def create_html_document_base64(json_path: str, title: str, output_filename: str = "output.html"):
    """
    從JSON文件和標題生成一個將圖片用Base64嵌入的HTML文件。

    Args:
        json_path (str): 輸入的JSON文件路徑。
        title (str): HTML頁面的主标题。
        output_filename (str): 输出的HTML文件名。
    """
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"错误：找不到文件 '{json_path}'")
        return
    except json.JSONDecodeError:
        print(f"错误：无法解析JSON文件 '{json_path}'")
        return

    # 獲取JSON檔案所在的目錄，以便處理相對圖片路徑
    json_dir = os.path.dirname(os.path.abspath(json_path))

    # HTML 模板和 CSS 樣式 (保持不變)
    html_content = f"""
<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{title}</title>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif;
            line-height: 1.6;
            color: #333;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            background-color: #f8f9fa;
        }}
        h1 {{
            color: #2c3e50;
            text-align: center;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
        }}
        .step {{
            background-color: #ffffff;
            border-radius: 8px;
            box-shadow: 0 4px 8px rgba(0,0,0,0.1);
            margin-bottom: 25px;
            padding: 20px;
            overflow: hidden;
        }}
        .step h2 {{
            color: #3498db;
            font-size: 1.5em;
            margin-top: 0;
        }}
        .step img {{
            max-width: 100%;
            height: auto;
            border-radius: 5px;
            border: 1px solid #ddd;
            display: block;
            margin: 0 auto 15px;
        }}
        .step p {{
            font-size: 1.1em;
            color: #555;
        }}
    </style>
</head>
<body>

    <h1>{title}</h1>

"""

    # ★★★ 修改點：遍歷JSON並嵌入Base64圖片 ★★★
    for image_filename, details in data.items():
        step_title = details[0]
        step_description = details[1]

        # 組合出圖片的完整絕對路徑
        full_image_path = os.path.join(json_dir, image_filename)

        # 將圖片編碼為Base64
        base64_uri = _encode_image_to_base64(full_image_path)

        # 只有在成功生成Base64 URI時才創建<img>標籤
        img_tag = ""
        if base64_uri:
            img_tag = f'<img src="{base64_uri}" alt="{step_title}">'

        # 將生成的HTML片段添加到主內容中
        html_content += f"""
    <div class="step">
        <h2>{step_title}</h2>
        {img_tag}
        <p>{step_description}</p>
    </div>
"""

    html_content += """
</body>
</html>
"""

    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"🎉 成功生成內嵌圖片的HTML文件：'{os.path.abspath(output_filename)}'")
    except IOError:
        print(f"错误：无法写入文件 '{output_filename}'")


# --- 示例用法 ---
if __name__ == "__main__":
    json_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\document_step.json"
    title_text = "Let's walk through how to add the © symbol to your Excel spreadsheet."
    output_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\help_document_embedded.html"

    create_html_document_base64(json_file, title_text, output_file)