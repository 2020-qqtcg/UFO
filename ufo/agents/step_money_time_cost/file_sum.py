import os

def count_immediate_subdirs(folder_path):
    """
    统计某个文件夹下的一级子文件夹数量。
    """
    if not os.path.isdir(folder_path):
        print(f"❌ 路径无效或不是文件夹：{folder_path}")
        return 0
    return sum(
        os.path.isdir(os.path.join(folder_path, entry))
        for entry in os.listdir(folder_path)
    )

# 给定的三个文件夹路径（根据需要替换成实际路径）
folders = [
    r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_withoutapi\20250725_bing_4.1_cost_complete_without_api_complete_double",
    r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_withoutapi\20250725_m365_4.1_cost_complete_without_api_double_complete",
    r"C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_withoutapi\20250726_qabench_4.1_cost_without_api_complete_double",
]


folders = [
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_bing_4.1_cost_complete_double',
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_m365_4.1_cost_complete_double',
        r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result\20250725_qabench_4.1_cost_complete_double'
    ]

# 输出每个文件夹的一级子文件夹数量
for path in folders:
    count = count_immediate_subdirs(path)
    print(f"📁 {path} 中的一级子文件夹数量为：{count}")


# 491
# 614
