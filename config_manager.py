import json
import os

CONFIG_FILE = "config.json"

def load_config(word_var, excel_var, folder_var):
    """รับตัวแปร StringVar มาโหลด Path เดิมที่เคยเซฟไว้ใส่เข้าไป"""
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                word_var.set(data.get("word_path", ""))
                excel_var.set(data.get("excel_path", ""))
                folder_var.set(data.get("folder_path", ""))
        except Exception:
            pass

def save_config(word_var, excel_var, folder_var):
    """ดึงค่าจาก StringVar มาเซฟเก็บไว้เป็นไฟล์ JSON"""
    data = {
        "word_path": word_var.get(),
        "excel_path": excel_var.get(),
        "folder_path": folder_var.get()
    }
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except Exception:
        pass