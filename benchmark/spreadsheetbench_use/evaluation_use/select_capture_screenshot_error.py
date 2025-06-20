import os

# 目标根目录
root_dir = r"C:\Users\v-yuhangxie\repos\UFO\logs\20250519"
# 错误关键词
error_keyword = "'capture_screenshot':"

# 遍历子文件夹
for subdir in os.listdir(root_dir):
    sub_path = os.path.join(root_dir, subdir)
    if os.path.isdir(sub_path):
        output_md_path = os.path.join(sub_path, "output.md")
        if os.path.exists(output_md_path):
            try:
                with open(output_md_path, "r", encoding="utf-8") as file:
                    # print(output_md_path)
                    content = file.read()
                    if error_keyword in content:
                        print(subdir)
            except Exception as e:
                print(f"读取文件出错: {output_md_path}, 原因: {e}")
