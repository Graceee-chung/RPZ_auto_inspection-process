
import os
import re
import time
import shutil
import logging
from datetime import datetime
import win32com.client as win32
from rpz_toxic import process_folder_toxic
from rpz_fraud import process_folder_fraud
from rpz_smoke import process_folder_smoke


today = datetime.today()
rpz_root = f"C:\\Users\\user\\Desktop\\RPZ\\RPZ_auto"
doc_folder = os.path.join(rpz_root, today.strftime('%Y%m%d'))

outlook = win32.Dispatch('Outlook.Application')

log_dir = datetime.today().strftime("%Y%m%d") + f"\\logs"
log_dir = os.path.join(rpz_root, log_dir)
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, f"{datetime.today().strftime('%Y%m%d')}.log")

logging.basicConfig(
    filename=log_file,
    filemode='a',
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

def move_to_finish_folder(folder_path, folder_name, base_folder):
    finish_dir_name = datetime.today().strftime("%Y%m%d") + "-finished"
    finish_dir = os.path.join(base_folder, finish_dir_name)
    os.makedirs(finish_dir, exist_ok=True)
    target_path = os.path.join(finish_dir, folder_name)
    shutil.move(folder_path, target_path)

def move_to_error_folder(folder_path, folder_name, base_folder):
    error_dir_name = datetime.today().strftime("%Y%m%d") + "-error"
    error_dir = os.path.join(base_folder, error_dir_name)
    os.makedirs(error_dir, exist_ok=True)
    target_path = os.path.join(error_dir, folder_name)
    shutil.move(folder_path, target_path)
    print(f"[→] 已移動錯誤資料夾至：{error_dir_name}")


for folder_name in os.listdir(doc_folder):
    folder_path = os.path.join(doc_folder, folder_name)
    if not os.path.isdir(folder_path):
        continue

    try:
        if "警署刑毒緝字" in folder_name:
            process_folder_toxic(folder_path, folder_name, doc_folder, outlook)
        elif "刑詐防字" in folder_name or "調資肆字" in folder_name:
            print(f'Enter Porocess Fraud')
            process_folder_fraud(folder_path, folder_name, doc_folder, outlook)
        elif "衛授國字" in folder_name:
            print(f'Enter Porocess Smoke')
            process_folder_smoke(folder_path, folder_name, doc_folder, outlook)
        elif not folder_name.endswith("號"):
            continue  # 忽略非"號"結尾的資料夾

        else:
            logging.warning(f"[!] 非刑毒詐文號類型，略過：{folder_name}")
            # move_to_error_folder(folder_path, folder_name, doc_folder)
        
        move_to_finish_folder(folder_path, folder_name, doc_folder)

    except Exception as e:
        logging.error(f"{folder_name} 發生錯誤：{e}")
        print(f"🚨處理 {folder_name} 時發生錯誤：{e}")
        move_to_error_folder(folder_path, folder_name, doc_folder)

    time.sleep(3) 

