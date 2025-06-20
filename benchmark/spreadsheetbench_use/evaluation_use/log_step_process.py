import os
import json

# 顶层目录
root_dir = r'C:\Users\v-yuhangxie\UFO\logs\20250530_305-456'

# 遍历子文件夹
for subdir, dirs, files in os.walk(root_dir):
    if 'response.log' in files:
        input_path = os.path.join(subdir, 'response.log')
        output_path = os.path.join(subdir, 'thought_steps.txt')

        thoughts = []
        step_count = 0

        with open(input_path, 'r', encoding='utf-8') as infile:
            for line in infile:
                line = line.strip()
                if not line:
                    continue
                try:
                    data = json.loads(line)
                    if "Thought" in data:
                        step_count += 1
                        thoughts.append(f"step {step_count}: {data['Thought']}")
                except json.JSONDecodeError:
                    continue

        if thoughts:
            with open(output_path, 'w', encoding='utf-8') as outfile:
                outfile.write('\n'.join(thoughts))
            print(f"[✓] 提取 {step_count} 条 thought，保存至：{output_path}")
        else:
            print(f"[!] 未找到有效 thought：{input_path}")
