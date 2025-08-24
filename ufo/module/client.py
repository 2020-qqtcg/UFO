# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.
import smtplib
import threading
from concurrent.futures import ThreadPoolExecutor
from email.header import Header
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from multiprocessing.pool import ThreadPool
from typing import List
import time

from tqdm import tqdm

from ufo.config.config import Config
from ufo.module.basic import BaseSession
from ufo.utils.azure_storage import AzureBlobStorage

_configs = Config.get_instance().config_data

import winreg


def delete_entire_options_key(version='16.0'):
    """
    删除整个 Options 键（可让 Excel 恢复默认设置）
    """
    parent_path = rf"Software\Microsoft\Office\{version}\Excel"
    parent_path_2 = fr"Software\Microsoft\Office\16.0\Common\Toolbars\Excel"
    try:
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, parent_path + r"\Options")
        print("Options key deleted.")
    except FileNotFoundError:
        print("Options key not found.")
    except PermissionError:
        print("Permission denied. Please run as administrator.")
    except OSError as e:
        print(f"Failed to delete key: {e}")

import winreg
import sys

def delete_registry_value():
    """
    從指定的登錄檔路徑中刪除 'QuickAccessToolbarStyle' 這個值。
    """
    # 確保腳本在 Windows 系統上執行
    if sys.platform != 'win32':
        print("❌ 此腳本僅能在 Windows 上執行。")
        return

    # 定義登錄檔路徑和要刪除的值的名稱
    # '16.0' 對應 Office 2016, 2019, 2021 及 Microsoft 365
    key_path = r"Software\Microsoft\Office\16.0\Common\Toolbars\Excel"
    value_name = "QuickAccessToolbarStyle"

    print(f"準備從以下路徑刪除值...")
    print(f"路徑: HKEY_CURRENT_USER\\{key_path}")
    print(f"值名稱: {value_name}\n")

    try:
        # 步驟 1: 開啟指定的登錄檔機碼，並請求寫入權限
        # winreg.HKEY_CURRENT_USER 代表這是目前登入使用者的設定
        key_handle = winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_WRITE)

        # 步驟 2: 刪除指定的值
        winreg.DeleteValue(key_handle, value_name)

        # 步驟 3: 關閉登錄檔機碼，釋放資源
        winreg.CloseKey(key_handle)

        print("✅ 操作成功！已刪除指定的登錄檔值。")
        print("此變更可能需要重新啟動 Excel 才會生效。")

    except FileNotFoundError:
        print(f"⚠️ 操作失敗：找不到路徑 '{key_path}' 或該路徑下不存在 '{value_name}' 這個值。")
    except Exception as e:
        print(f"❌ 發生未預期的錯誤: {e}")

class UFOClientManager:
    """
    The manager for the UFO clients.
    """

    def __init__(self, session_list: List[BaseSession]) -> None:
        """
        Initialize a batch UFO client.
        """

        self._session_list = session_list

    def run_all(self) -> None:
        """
        Run the batch UFO client.
        """
        blob_storage = None
        if _configs["UPLOAD"]:
            blob_storage = AzureBlobStorage()

        total = len(self.session_list)

        with ThreadPoolExecutor(max_workers=16) as executor:
            for idx, session in enumerate(tqdm(self.session_list), start=1):
                print("delete--------------------------------------------")
                # delete_entire_options_key(version="16.0")
                delete_registry_value()
                time.sleep(0.2)
                session.run()

                if _configs["MONITOR"]:
                    send_point = _configs["SEND_POINT"].split(",")
                    if str(idx) in send_point:
                        message = f"Ufo Execute Completed: {idx}/{total}"
                        send_message(message)

                if _configs["UPLOAD"]:
                    executor.submit(lambda : blob_storage.upload_folder(session.log_path, _configs["DATA_SOURCE"]))


    @property
    def session_list(self) -> List[BaseSession]:
        """
        Get the session list.
        :return: The session list.
        """
        return self._session_list

    def add_session(self, session: BaseSession) -> None:
        """
        Add a session to the session list.
        :param session: The session to add.
        """
        self._session_list.append(session)

    def next_session(self) -> BaseSession:
        """
        Get the next session.
        :return: The next session.
        """
        return self._session_list.pop(0)

def send_message(message: str) -> None:
    """
    Send a message.
    :param message: message to send.
    """
    # email info
    sender_email = _configs["FROM_EMAIL"]
    sender_password = _configs["SENDER_PASSWORD"]
    receiver_email = _configs["TO_EMAIL"]
    server_host = _configs["SMTP_SERVER"]
    machine_id = _configs["MACHINE_ID"]

    msg = MIMEMultipart()
    msg["From"] = _configs["FROM_EMAIL"]
    msg["TO"] = _configs["TO_EMAIL"]
    msg["Subject"] = Header(f"Execute Reminder: {machine_id}")
    msg.attach(MIMEText(message, "plain", 'utf-8'))

    server = None
    try:
        # Connect SMTP Server
        server = smtplib.SMTP(server_host, 587)
        server.starttls()
        server.login(sender_email, sender_password)

        # send email
        server.sendmail(sender_email, receiver_email, msg.as_string())
        print(f"Send Email to {receiver_email} sucessfully.")
    except Exception as e:
        print(f"Send Email to {receiver_email} failed: {e}.")
    finally:
        try:
            if server:
                server.quit()
        except Exception:
            return
