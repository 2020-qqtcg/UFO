import os
import re

# 目标根目录
root_dir = r"C:\Users\v-yuhangxie\repos\UFO\logs\20250519"
# 正则表达式匹配 "select the range xxx failed"
pattern = re.compile(r"select the range [A-Z]+\d+:[A-Z]*\d* failed", re.IGNORECASE)

# 遍历子文件夹
for subdir in os.listdir(root_dir):
    sub_path = os.path.join(root_dir, subdir)
    if os.path.isdir(sub_path):
        output_md_path = os.path.join(sub_path, "output.md")
        if os.path.exists(output_md_path):
            try:
                with open(output_md_path, "r", encoding="utf-8") as file:
                    content = file.read()
                    matches = pattern.findall(content)
                    if matches:
                        print(f"子文件夹: {subdir}")
                        for match in matches:
                            print(f"  匹配内容: {match}")
            except Exception as e:
                print(f"读取文件出错: {output_md_path}, 原因: {e}")
