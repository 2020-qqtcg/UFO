import os
import shutil


def process_folders(source_dir: str, dest_dir: str):
    """
    遍历源目录中的子文件夹，如果它们包含特定的HTML和MP4文件，
    就将它们复制并重命名到目标目录下的相应子文件夹中。

    Args:
        source_dir (str): 包含多个子文件夹的源目录路径。
        dest_dir (str): 用于保存结果的目标目录路径。
    """
    # 检查源目录是否存在
    if not os.path.isdir(source_dir):
        print(f"错误：源目录 '{source_dir}' 不存在或不是一个目录。")
        return

    # 1. 定义并创建目标子文件夹
    doc_dest_folder = os.path.join(dest_dir, 'documents')
    video_dest_folder = os.path.join(dest_dir, 'videos')

    os.makedirs(doc_dest_folder, exist_ok=True)
    os.makedirs(video_dest_folder, exist_ok=True)

    print(f"目标目录 '{os.path.abspath(dest_dir)}' 已准备就绪。")
    print(f" - HTML文件将保存到: '{doc_dest_folder}'")
    print(f" - 视频文件将保存到: '{video_dest_folder}'")

    # 初始化ID计数器
    pair_id = 0

    print("\n开始扫描文件夹...")

    # 遍历源目录中的所有项目（文件和文件夹）
    for folder_name in sorted(os.listdir(source_dir)):
        subfolder_path = os.path.join(source_dir, folder_name)

        # 确保我们只处理文件夹
        if os.path.isdir(subfolder_path):
            # 定义需要检查的两个文件的原始路径
            html_path = os.path.join(subfolder_path, 'document', 'help_document.html')
            video_path = os.path.join(subfolder_path, 'video_demo', 'video_demo.mp4')

            # 检查这两个文件是否都存在
            if os.path.exists(html_path) and os.path.exists(video_path):
                print(f"  [✓] 在 '{folder_name}' 中找到有效文件对。分配 ID: {pair_id}")

                # 2. 定义新的目标文件名
                new_html_filename = f"{pair_id}_help_document.html"
                new_video_filename = f"{pair_id}_video_demo.mp4"

                # 3. 定义包含新文件名和子文件夹的完整目标路径
                dest_html_path = os.path.join(doc_dest_folder, new_html_filename)
                dest_video_path = os.path.join(video_dest_folder, new_video_filename)

                # 4. 复制文件到新的位置并重命名
                try:
                    shutil.copy2(html_path, dest_html_path)
                    shutil.copy2(video_path, dest_video_path)
                    print(f"      -> 已成功复制文件对。")

                    # 5. 递增ID，为下一个文件对做准备
                    pair_id += 1

                except Exception as e:
                    print(f"      -> 复制文件时出错: {e}")
            else:
                print(f"  [✗] 在 '{folder_name}' 中未找到完整的文件对，已跳过。")

    print(f"\n处理完成。总共找到并复制了 {pair_id} 个有效的文件对。")


# --- 如何使用 ---
if __name__ == '__main__':
    # --- 准备一个测试环境（可选） ---
    # 如果您想先用虚拟文件进行测试，可以取消下面这部分代码的注释。
    # print("--- 正在创建测试环境 ---")
    # if os.path.exists('source_folders_test'):
    #     shutil.rmtree('source_folders_test')
    # if os.path.exists('output_folders_test'):
    #     shutil.rmtree('output_folders_test')
    # os.makedirs('source_folders_test/task_A/document', exist_ok=True)
    # os.makedirs('source_folders_test/task_A/video_demo', exist_ok=True)
    # open('source_folders_test/task_A/document/help_document.html', 'w').close()
    # open('source_folders_test/task_A/video_demo/video_demo.mp4', 'w').close()
    # os.makedirs('source_folders_test/task_C/document', exist_ok=True)
    # os.makedirs('source_folders_test/task_C/video_demo', exist_ok=True)
    # open('source_folders_test/task_C/document/help_document.html', 'w').close()
    # open('source_folders_test/task_C/video_demo/video_demo.mp4', 'w').close()
    # print("--- 测试环境创建完毕 ---\n")
    # source_directory = 'source_folders_test'
    # destination_directory = 'output_folders_test'

    # --- 配置您的实际路径 ---
    # *****************************************************************
    # ** 请将这里的路径替换为您的实际工作路径 **
    source_directory = r'C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\chunk1'
    destination_directory = r'C:\Users\v-yuhangxie\OneDrive - Microsoft\qabench\qabench\logs\select'
    # *****************************************************************

    # 执行处理函数
    process_folders(source_dir=source_directory, dest_dir=destination_directory)