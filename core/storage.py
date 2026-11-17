import json
import os
from datetime import datetime

DATA_DIR = os.path.join(os.path.expanduser("~"), ".pcb_auto_excel")
SETTINGS_FILE = os.path.join(DATA_DIR, "settings.json")
HISTORY_FILE = os.path.join(DATA_DIR, "history.json")

DEFAULT_SETTINGS = {
    "last_posylka": "",
    "last_pp": "",
    "last_filial": "",
    "save_dir": os.path.join(os.path.expanduser("~"), "Desktop"),
    "auto_load_files": True,
    "confirm_overwrite": True,
    "backup_results": True,
    "theme": "DARK"
}

def ensure_dir():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

def load_settings():
    ensure_dir()
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                return {**DEFAULT_SETTINGS, **data}
        except Exception:
            pass
    return DEFAULT_SETTINGS.copy()

def save_settings(settings):
    ensure_dir()
    with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=4)

def load_history():
    ensure_dir()
    if os.path.exists(HISTORY_FILE):
        try:
            with open(HISTORY_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return []

def add_history_record(record):
    ensure_dir()
    history = load_history()
    record["timestamp"] = datetime.now().strftime("%d.%m.%Y %H:%M:%S")
    history.insert(0, record)
    history = history[:100]
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=4)
