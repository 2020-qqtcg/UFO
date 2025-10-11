import json
import os

def create_html_document(json_path: str, title: str, output_filename: str = "output.html"):
    """
    从JSON文件和标题生成一个HTML文件。

    Args:
        json_path (str): 输入的JSON文件路径。
        title (str): HTML页面的主标题。
        output_filename (str): 输出的HTML文件名。
    """
    try:
        # 1. 读取并解析JSON文件
        with open(json_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"错误：找不到文件 '{json_path}'")
        return
    except json.JSONDecodeError:
        print(f"错误：无法解析JSON文件 '{json_path}'")
        return

    # 2. 开始构建HTML字符串，并加入CSS样式
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

    # 3. 遍历JSON数据，为每个条目生成HTML内容
    for image_path, details in data.items():
        step_title = details[0]
        step_description = details[1]

        html_content += f"""
    <div class="step">
        <h2>{step_title}</h2>
        <img src="{image_path}" alt="{step_title}">
        <p>{step_description}</p>
    </div>
"""

    # 4. 结束HTML结构
    html_content += """
</body>
</html>
"""

    # 5. 将HTML内容写入文件
    try:
        with open(output_filename, 'w', encoding='utf-8') as f:
            f.write(html_content)
        print(f"🎉 成功生成HTML文件：'{os.path.abspath(output_filename)}'")
    except IOError:
        print(f"错误：无法写入文件 '{output_filename}'")


# --- 示例用法 ---
if __name__ == "__main__":
    # 定义您的输入和输出文件路径。
    # 您应该更改这些路径以匹配您的实际文件位置。
    json_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\document_step.json"
    title_text = "Let's walk through how to add the © symbol to your Excel spreadsheet."
    # 将输出文件扩展名更改为.html
    output_file = r"C:\Users\v-yuhangxie\repos\excel_traj_v2\excel_traj_v2\ufo_execute_log\add_a_special_character_or_symbol_4f364db0-912b-46b3-8282-2d8dd49c336a\video_demo\help_document.html"

    # 调用主函数生成文档
    create_html_document(json_file, title_text, output_file)