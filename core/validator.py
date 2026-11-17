import os
import pandas as pd
from datetime import datetime

def get_file_info(path):
    if not path:
        return {"exists": False, "name": "Не выбран", "date": "-", "status": "Ожидание файла"}
    
    clean_path = path.strip(' "')
    if os.path.exists(clean_path) and os.path.isfile(clean_path):
        mtime = os.path.getmtime(clean_path)
        date_str = datetime.fromtimestamp(mtime).strftime("%d.%m.%Y %H:%M")
        return {
            "exists": True,
            "name": os.path.basename(clean_path),
            "path": clean_path,
            "date": date_str,
            "status": "Доступен"
        }
    return {
        "exists": False,
        "name": os.path.basename(clean_path) if clean_path else "Неизвестно",
        "path": clean_path,
        "date": "-",
        "status": "Файл не найден или недоступна сеть"
    }

def validate_excel_files(p_pos, p_pp, p_fil):
    files = [
        ("Посылка (Накладная)", p_pos, "posylka"),
        ("Сводка ПП (Голова)", p_pp, "pp"),
        ("Сводка Филиал", p_fil, "filial")
    ]
    
    for title, path, f_type in files:
        if not path or not path.strip():
            return False, f"Не указан путь для файла: {title}"
        
        clean_path = path.strip(' "')
        if not os.path.exists(clean_path):
            return False, f"Файл «{title}» не найден по пути:\n{clean_path}\nПроверьте подключение к сетевой папке."
        
        try:
            if f_type == "posylka":
                df = pd.read_excel(clean_path)
                if df.empty:
                    return False, f"Таблица в файле «{title}» абсолютно пуста."
                required_cols = ['PS', 'По счету', 'Доставка']
                missing = [c for c in required_cols if c not in df.columns]
                if missing:
                    return False, f"В файле «{title}» отсутствуют обязательные столбцы: {', '.join(missing)}"
            else:
                df = pd.read_excel(clean_path, sheet_name=0, header=None)
                if df.empty:
                    return False, f"Таблица в файле «{title}» пуста."
                min_cols = 15 if f_type == "pp" else 8
                if df.shape[1] < min_cols:
                    return False, f"Файл «{title}» имеет неверную структуру (слишком мало столбцов: {df.shape[1]}, ожидалось минимум {min_cols})."
        except PermissionError:
            return False, f"Нет доступа к файлу «{title}». Возможно, он открыт у другого пользователя в Excel."
        except Exception as e:
            return False, f"Файл «{title}» поврежден или имеет неверный формат.\nТехническая деталь: {str(e)}"
            
    return True, "Все файлы успешно прошли проверку."
