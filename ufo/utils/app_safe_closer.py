import os
import time

import psutil
import win32api
import win32com.client
import pythoncom
import win32con
import win32gui
import win32process

try:
    from pywinauto import Desktop

    PYWINAUTO_AVAILABLE = True
except ImportError:
    PYWINAUTO_AVAILABLE = False


# Standard dialog window class name
DIALOG_CLASS_NAME = "#32770"


def get_excel_pids():
    """Gets a list of PIDs for all running EXCEL.EXE processes"""
    excel_pids = []
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'].lower() == 'excel.exe':
                excel_pids.append(proc.info['pid'])
        except (psutil.NoSuchProcess, psutil.AccessDenied, psutil.ZombieProcess):
            # The process might have ended while we were checking it
            pass
    return excel_pids


# List of known Excel dialog window class names
# You can add more class names found through debugging as needed
KNOWN_EXCEL_DIALOG_CLASS_NAMES = [
    "#32770",  # Standard Windows dialog class name
    "bosa_sdm_XL9"  # "Format Cells" dialog class name found from your debug output
]

# List to store the handles of found Excel dialog windows
excel_dialog_hwnds = []


# (Optional) For debugging: store information of all found Excel process windows
# debug_excel_windows_info = []


def enum_windows_callback(hwnd, excel_pids_param):
    """
    Callback function for EnumWindows.
    Checks if the window is an Excel dialog and adds its handle to the global list.
    """
    global excel_dialog_hwnds  # , debug_excel_windows_info

    if not win32gui.IsWindowVisible(hwnd):
        return True

    try:
        _, found_pid = win32process.GetWindowThreadProcessId(hwnd)
    except Exception:
        return True

    if found_pid in excel_pids_param:
        try:
            class_name = win32gui.GetClassName(hwnd)
            window_text = win32gui.GetWindowText(hwnd)
            is_enabled = win32gui.IsWindowEnabled(hwnd)

            # (Optional) Record information of all windows belonging to Excel PIDs (for further debugging)
            # debug_excel_windows_info.append({
            #     "hwnd": hwnd, "title": window_text, "class": class_name,
            #     "visible": True, "enabled": is_enabled, "pid": found_pid
            # })

            # Modified dialog identification logic: check if the class name is in the known list
            if is_enabled and class_name in KNOWN_EXCEL_DIALOG_CLASS_NAMES:
                print(f"INFO: Found matching Excel dialog - Title='{window_text}', ClassName='{class_name}', HWND={hwnd}")
                if hwnd not in excel_dialog_hwnds:
                    excel_dialog_hwnds.append(hwnd)
        except Exception:
            pass  # Skip if there's an error getting window details

    return True


def close_all_excel_dialogs_fixed():
    """Finds and tries to close all Excel interactive dialogs (using updated class name list)"""
    global excel_dialog_hwnds  # , debug_excel_windows_info
    excel_dialog_hwnds = []
    # debug_excel_windows_info = []

    excel_pids = get_excel_pids()
    if not excel_pids:
        print("No running Excel processes found.")
        return

    print(f"Found Excel process PIDs: {excel_pids}")

    print("Enumerating windows...")
    win32gui.EnumWindows(enum_windows_callback, excel_pids)
    print("Window enumeration complete.")

    # (Optional) Debug information output
    # if debug_excel_windows_info:
    #     print("\n--- DEBUG: Information of all windows belonging to Excel PIDs ---")
    #     for info in debug_excel_windows_info:
    #         print(f"  HWND: {info['hwnd']}, Title: '{info['title']}', ClassName: '{info['class']}', "
    #               f"Enabled: {info['enabled']}, Visible: {info['visible']}, PID: {info['pid']}")
    #     print("--- END DEBUG ---")

    if not excel_dialog_hwnds:
        print("\nScript finished. No Excel interactive dialogs found to close based on preset conditions.")
        print(
            f"Please ensure the target dialog is open and its class name is included in the script's KNOWN_EXCEL_DIALOG_CLASS_NAMES list (current list: {KNOWN_EXCEL_DIALOG_CLASS_NAMES}).")
        return

    print(f"\nIdentified {len(excel_dialog_hwnds)} matching dialogs, attempting to close by sending ESC key...")

    closed_count = 0
    for hwnd_dialog in excel_dialog_hwnds:
        try:
            dialog_title_to_close = win32gui.GetWindowText(hwnd_dialog)
            print(f"  Attempting to close dialog: '{dialog_title_to_close}' (HWND: {hwnd_dialog})")

            try:
                win32gui.SetForegroundWindow(hwnd_dialog)
                time.sleep(0.15)
            except Exception as e_fg:
                print(f"    WARNING: Failed to set HWND {hwnd_dialog} as foreground window: {e_fg}. Still attempting to send key...")

            win32api.keybd_event(win32con.VK_ESCAPE, 0, 0, 0)
            time.sleep(0.05)
            win32api.keybd_event(win32con.VK_ESCAPE, 0, win32con.KEYEVENTF_KEYUP, 0)

            print(f"    Sent ESC key to HWND {hwnd_dialog}.")
            closed_count += 1
            time.sleep(0.25)

        except Exception as e_close:
            print(f"    Error closing dialog HWND {hwnd_dialog}: {e_close}")

    if closed_count > 0:
        print(f"\nOperation complete. Attempted to close {closed_count} dialogs.")
    else:
        print("\nNo dialogs were successfully closed.")
    print("Please check Excel to confirm if the dialogs are closed.")

def save_and_close_app(app_name: str, save_as_path: str) -> bool:
    """
    保存当前活动的Excel工作簿到指定路径，然后关闭Excel应用程序。
    如果初始连接失败，此函数将找到所有Excel窗口并发送Escape键以解除
    任何可能的繁忙状态。此版本兼容旧版的pywinauto。

    Args:
        app_name (str): 目标应用程序的进程名 (例如 "EXCEL.EXE")。
        save_as_path (str): 保存文件的完整绝对路径。目录必须预先存在。

    Returns:
        bool: 如果成功保存并关闭则返回 True，否则返回 False。
    """
    close_all_excel_dialogs_fixed()

    if app_name.upper() != "EXCEL.EXE":
        print(f"错误：此函数当前仅支持 'EXCEL.EXE'，而不是 '{app_name}'。")
        return False

    save_dir = os.path.dirname(save_as_path)
    if not os.path.isdir(save_dir):
        print(f"错误：指定的目录不存在: {save_dir}")
        return False

    excel = None
    context = None
    original_display_alerts = None

    for attempt in range(2):
        try:
            context = pythoncom.CoInitialize()
            excel = win32com.client.GetObject(Class="Excel.Application")
            _ = excel.Ready
            print("成功连接到响应的Excel实例。")
            break
        except (AttributeError, pythoncom.com_error) as e:
            if excel: excel = None
            if context: pythoncom.CoUninitialize(); context = None

            if attempt == 0:
                print("COM连接失败，Excel可能正忙。")
                print("正在尝试通过发送 'Escape' 键进行干预...")

                if not PYWINAUTO_AVAILABLE:
                    print("错误：pywinauto库未安装，无法进行主动干预。请运行 'pip install pywinauto'。")
                    return False

                try:
                    desktop = Desktop(backend="win32")
                    excel_windows = desktop.windows(class_name='XLMAIN', visible_only=True)

                    if not excel_windows:
                        print("pywinauto干预失败：未找到任何可见的Excel窗口。")
                        continue

                    print(f"找到 {len(excel_windows)} 个Excel窗口。将向所有窗口发送 'Escape' 键以确保解锁。")
                    for window in excel_windows:
                        try:
                            # --- 核心修正 ---
                            # 移除了 with_pause 参数以兼容旧版本的 pywinauto
                            window.type_keys('{ESC}')
                        except Exception as single_window_error:
                            print(f"警告：向一个窗口句柄 {window.handle} 发送ESC键时出错: {single_window_error}")

                    print("已向所有找到的Excel窗口发送 'Escape' 键。正在重试COM连接...")
                    # 这个延时仍然非常重要，确保Excel有时间处理按键消息
                    time.sleep(0.5)

                except Exception as pywinauto_error:
                    print(f"pywinauto干预失败，发生意外错误: {pywinauto_error}")
                    return False
            else:
                print("发送'Escape'键后，Excel仍然无响应。操作失败。")
                return False

    # 如果循环结束后 excel 对象仍然是 None，说明所有尝试都失败了
    if not excel:
        return False

    try:
        original_display_alerts = excel.DisplayAlerts
        excel.DisplayAlerts = False

        if excel.Workbooks.Count == 0:
            print("错误：未找到打开的Excel工作簿。")
            excel.Quit()
            return False

        workbook = excel.ActiveWorkbook
        print(f"正在将 '{workbook.Name}' 保存到 '{save_as_path}'...")
        workbook.SaveAs(save_as_path, FileFormat=51)
        print("保存成功。")

        excel.Quit()
        print("Excel应用程序已关闭。")
        return True

    except Exception as e:
        print(f"在操作Excel时发生未知错误: {e}")
        if excel:
            excel.Quit()
        return False
    finally:
        if excel and original_display_alerts is not None:
            excel.DisplayAlerts = original_display_alerts
        if context:
            pythoncom.CoUninitialize()

# --- 示例用法 ---
if __name__ == '__main__':
    # **重要提示**: 运行此示例前，请手动打开一个Excel文件。

    # 1. 定义保存路径 (请根据你的实际情况修改)
    #    确保目录 "C:\\temp\\excel_test" 存在
    save_directory = "D:\\"
    if not os.path.exists(save_directory):
        os.makedirs(save_directory)  # 如果目录不存在则创建

    file_path_to_save = os.path.join(save_directory, "my_saved_workbook.xlsx")

    # 2. 调用函数
    print("正在尝试保存并关闭活动的Excel文件...")
    success = save_and_close_app("EXCEL.EXE", file_path_to_save)

    # 3. 打印结果
    if success:
        print(f"\n操作成功完成！文件已保存到: {file_path_to_save}")
    else:
        print("\n操作失败。请检查上面的错误信息。")