import re
import json

def extract_request_from_md(md_path):
    with open(md_path, 'r', encoding='utf-8') as f:
        content = f.read()

    # 使用正则表达式提取 request 字段内容
    match = re.search(r'- \*\*request\*\*: (.+)', content)
    if match:
        return match.group(1).strip()
    else:
        return None


def extract_and_clean_requests(file_path='eval.log'):
    """
    读取一个包含JSON对象的日志文件，
    提取每个对象中'request'字段的值，并移除换行符。

    Args:
        file_path (str): 日志文件的路径。

    Returns:
        list: 一个包含所有清理后request字符串的列表。
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            for line in f:
                # 跳过空行
                if not line.strip():
                    continue
                try:
                    # 将行文本解析为Python字典
                    data = json.loads(line)

                    # 检查 'request' 键是否存在
                    if 'request' in data:
                        request_text = data['request']

                        # 移除换行符 '\n'
                        cleaned_text = request_text.replace('\nDo not use agents such as Copilot, ChatGPT, or any browser-based assistance.', '')
                        print(cleaned_text)
                        return(cleaned_text)

                except json.JSONDecodeError:
                    pass
                    # print(f"警告: 无法解析该行的JSON: {line.strip()}")
                except KeyError:
                    pass
                    # print(f"警告: 在该行中未找到 'request' 键: {line.strip()}")
    except FileNotFoundError:
        print(f"错误: 文件 '{file_path}' 未找到。")

    return None