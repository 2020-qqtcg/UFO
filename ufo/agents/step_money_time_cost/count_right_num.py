import os


def count_subfolders_by_name(target_directory):
    """
    遍历指定目录，统计子文件夹名称中包含特定关键字的数量。

    Args:
        target_directory (str): 要扫描的目标文件夹路径。

    Returns:
        None: 直接打印统计结果。
    """
    # 检查路径是否存在
    if not os.path.exists(target_directory):
        print(f"错误：找不到路径 '{target_directory}'")
        return
    if not os.path.isdir(target_directory):
        print(f"错误：路径 '{target_directory}' 不是一个文件夹。")
        return

    # 初始化计数器
    n_total = 0  # 总的子文件夹数量
    n1_bing = 0  # 名字中包含 'bing' 的数量
    n2_m365 = 0  # 名字中包含 'm365' 的数量

    print(f"正在扫描文件夹: {target_directory}\n")

    # 遍历目录中的所有条目
    for item_name in os.listdir(target_directory):
        # 构建完整的路径
        full_path = os.path.join(target_directory, item_name)

        # 检查是否是文件夹
        if os.path.isdir(full_path):
            n_total += 1  # 总数加一

            # 检查文件夹名称是否包含关键字
            if 'bing' in item_name.lower():  # 使用 .lower() 进行不区分大小写的匹配
                n1_bing += 1

            if 'm365' in item_name.lower():  # 使用 .lower() 进行不区分大小写的匹配
                n2_m365 += 1

    # 计算其他类型的文件夹数量
    n_other = n_total - n1_bing - n2_m365

    # 打印最终结果
    print("--- 统计结果 ---")
    print(f"子文件夹总数量 (n): {n_total},{n_total/1559}")
    print(f"名字中包含 'bing' 的数量 (n1): {n1_bing},{n1_bing/700}")
    print(f"名字中包含 'm365' 的数量 (n2): {n2_m365},{n2_m365/721}")
    print(f"其他文件夹的数量 (n - n1 - n2): {n_other},{n_other/138}")
    print("------------------")

   #  # o5
   #  print("--- 统计结果 ---")
   #  print(f"子文件夹总数量 (n): {n_total},{443/1559}")
   #  print(f"名字中包含 'bing' 的数量 (n1): {n1_bing},{185/700}")
   #  print(f"名字中包含 'm365' 的数量 (n2): {n2_m365},{222/721}")
   #  print(f"其他文件夹的数量 (n - n1 - n2): {n_other},{36/138}")
   #  print("------------------")
   #
   # # UFO1
   #  print("--- 统计结果 ---")
   #  print(f"子文件夹总数量 (n): {n_total},{485/1559}")
   #  print(f"名字中包含 'bing' 的数量 (n1): {n1_bing},{210/700}")
   #  print(f"名字中包含 'm365' 的数量 (n2): {n2_m365},{236/721}")
   #  print(f"其他文件夹的数量 (n - n1 - n2): {n_other},{39/138}")
   #  print("------------------")
   #
   #
   #  # UFO2
   #  print("--- 统计结果 ---")
   #  print(f"子文件夹总数量 (n): {n_total},{486/1559}")
   #  print(f"名字中包含 'bing' 的数量 (n1): {n1_bing},{250/700}")
   #  print(f"名字中包含 'm365' 的数量 (n2): {n2_m365},{191/721}")
   #  print(f"其他文件夹的数量 (n - n1 - n2): {n_other},{45/138}")
   #  print("------------------")


if __name__ == '__main__':
    # --- 请在这里修改为您要扫描的文件夹路径 ---
    # 注意：在 Windows 路径中，建议使用正斜杠 '/' 或者双反斜杠 '\\' 来避免转义字符问题。
    directory_to_scan = r'C:\Users\v-yuhangxie\OneDrive - Microsoft\uiagent_result_operator_complete_double'

    # 运行统计函数
    count_subfolders_by_name(directory_to_scan)
