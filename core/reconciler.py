import os
import re
import pandas as pd
import numpy as np

def resolve_shortcut(path):
    clean = path.strip(' "')
    if clean.lower().endswith('.lnk') and os.path.exists(clean):
        try:
            import subprocess
            cmd = f'powershell -NoProfile -Command "(New-Object -ComObject WScript.Shell).CreateShortcut(\'{clean}\').TargetPath"'
            target = subprocess.check_output(cmd, shell=True, text=True).strip()
            if target and os.path.exists(target):
                return target
        except Exception:
            pass
    return clean

def clean_ps(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    s = str(val).strip()
    if s.lower() in ('nan', 'none', 'nat', '<na>', ''):
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    return digits if 5 <= len(digits) <= 7 else ''

def clean_schet_n(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    digits = re.sub(r'\D', '', s)
    return digits if 4 <= len(digits) <= 6 else ''

def is_board_code_or_phone(text):
    """Отсекает телефоны, децимальные номера плат (tpts..., АЮК...), даты и чисто цифровые коды."""
    s = str(text).strip()
    if not s or len(s) < 2:
        return True
    u = s.upper()
    if u in ('NAN', 'NONE', 'NAT', 'ОТК', 'ФИЛИАЛ', 'ФИРМА', 'ЗАКАЗЧИК', 'ФИРМА ЗАКАЗЧИК', 'ДОСТАВКА', 'ПРИМЕЧАНИЕ'):
        return True
    # Телефоны
    if re.match(r'^(\+7|8|7)[\s\-\(\d]{7,}', s) or re.search(r'\d{3}[\s\-]\d{2}[\s\-]\d{2}', s):
        return True
    # Даты
    if re.search(r'\b\d{2}\.\d{2}\.\d{2,4}\b', s) or 'ОТПРАВКА' in u or 'ПРИХОД' in u:
        return True
    # Коды плат и децимальные номера (латиница с цифрами, подчеркивания, коды типа АЮК6.717.489, tpts559150045)
    if re.search(r'[A-Za-z]{2,}[\-\_]?\d{3,}', s):
        return True
    if re.search(r'\d{3,}[\-\_]\d+', s) and not any(org in u for org in ['ООО', 'АО', 'ЗАО', 'ПАО', 'ИП ', 'ФГУП', 'НПЦ']):
        return True
    if re.search(r'[А-ЯA-Z]{2,}\d+\.\d+\.\d+', u) or 'ИЗМ.' in u or 'ТПТС' in u:
        return True

    digits_count = sum(c.isdigit() for c in s)
    letters_count = sum(c.isalpha() for c in s)
    if letters_count < 3:
        return True
    if digits_count > letters_count and not any(org in u for org in ['ООО', 'АО', 'ЗАО', 'ПАО', 'ИП']):
        return True
    return False

def normalize_company_key(name):
    s = str(name or "").strip().lower()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\b(дм|сдэк|сдек|cdek|самовывоз|курьер|к|ооо|ао|зао|пао|ип|фгуп|нпц|пз|мпотк)\b', '', s)
    s = re.sub(r'[^а-яa-z0-9]', '', s)
    return s

def extract_delivery_from_text(raw_text):
    if raw_text is None or (isinstance(raw_text, float) and np.isnan(raw_text)):
        return ""
    s = str(raw_text).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return ""
    u = s.upper()

    if "САМОВЫВОЗ" in u or "САМОВ" in u:
        return "Самовывоз"
    if "СДЭК" in u or "CDEK" in u or "СДЕК" in u:
        return "СДЭК"
    if re.search(r'(^|[\s,;/\(\-\.\:])ДМ($|[\s,;/\)\-\.\:])', u) or "ДЕЛОВ" in u or "ДЛ" == u:
        return "ДМ"
    if u == "К" or "КУРЬЕР" in u or "ДОСТАВКА К" in u or re.search(r'(^|[\s,;/\(\-\.\:])К($|[\s,;/\)\-\.\:])', u):
        return "К"
    return ""

def extract_firma_and_delivery(raw_text):
    if raw_text is None or (isinstance(raw_text, float) and np.isnan(raw_text)):
        return "", ""
    s = str(raw_text).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return "", ""

    deliv = extract_delivery_from_text(s)

    # Очищаем название фирмы от приписанного способа доставки в конце
    clean_name = re.sub(r'[\(\[\s,;\-\/\:]*(самовывоз|сдэк|сдек|cdek|курьер|дм|деловые\s*линии)[\)\]\s,;\-\/\.\:]*$', '', s, flags=re.IGNORECASE).strip()
    clean_name = re.sub(r'[\s,;\-\/\:]+К$', '', clean_name).strip()

    if clean_name.upper() in ("ДМ", "СДЭК", "СДЕК", "CDEK", "К", "САМОВЫВОЗ", "КУРЬЕР", ""):
        return "", deliv

    if is_board_code_or_phone(clean_name):
        return "", deliv

    return clean_name, deliv

def delivery_sort_rank(deliv_val):
    d = str(deliv_val).strip().upper()
    if d == "ДМ":
        return 1
    if d == "СДЭК":
        return 2
    if d == "К":
        return 3
    if d == "САМОВЫВОЗ":
        return 4
    return 5

def score_column_as_firma(df, col_idx, start_row):
    """Оценивает, насколько колонка похожа на колонку 'Фирма заказчик' (названия компаний/ФИО, а не коды плат)."""
    sample = df.iloc[start_row : min(start_row + 150, len(df)), col_idx].dropna().astype(str).str.strip()
    if len(sample) == 0:
        return -100
    good = 0
    bad = 0
    org_markers = 0
    for val in sample:
        if not val or val.lower() == 'nan':
            continue
        if is_board_code_or_phone(val):
            bad += 1
        else:
            good += 1
            u = val.upper()
            if any(k in u for k in ['ООО', 'АО', 'ЗАО', 'ПАО', 'ИП ', 'НПЦ', 'ФГУП', 'ЗАВОД', 'КОМПАНИЯ']):
                org_markers += 2
    return (good * 2 + org_markers) - (bad * 3)

def detect_sheet_columns(df, is_head=True):
    num_rows, num_cols = df.shape
    if num_rows == 0 or num_cols == 0:
        return None

    header_row_idx = -1
    best_score = 0

    for r_idx in range(min(15, num_rows)):
        row_vals = df.iloc[r_idx].tolist()
        row_strs = [str(x).strip().lower() for x in row_vals if pd.notna(x)]
        score = 0
        for cell in row_strs:
            if cell == 'ps' or re.search(r'\bps\b', cell):
                score += 4
            if 'фирма' in cell or 'заказчик' in cell or 'клиент' in cell:
                score += 5
            if any(k in cell for k in ['счет', 'счёт', 'кол', 'доставк', 'сч/н']):
                score += 2
        if score > best_score:
            best_score = score
            header_row_idx = r_idx

    cols = {
        "header_row": header_row_idx if best_score >= 4 else -1,
        "ps": None,
        "firma_zakazchik": None,
        "schet_pp": None,
        "schet_n": None,
        "qty": None,
        "deliv": None
    }

    start_data_row = cols["header_row"] + 1 if cols["header_row"] >= 0 else 0

    if cols["header_row"] >= 0:
        raw_headers = df.iloc[cols["header_row"]].tolist()
        headers = [str(x).strip().lower() if pd.notna(x) else "" for x in raw_headers]

        for idx, h in enumerate(headers):
            if 'фирма заказчик' in h or 'фирма-заказчик' in h:
                cols["firma_zakazchik"] = idx
                break
        if cols["firma_zakazchik"] is None:
            for idx, h in enumerate(headers):
                if any(bad in h for bad in ['тел', 'контакт', 'примеч', 'адрес', 'чертеж', 'децим', 'плата', 'изделие', 'наименование плат', 'файл']):
                    continue
                if 'заказчик' in h or h == 'фирма' or 'клиент' in h or 'покупатель' in h:
                    cols["firma_zakazchik"] = idx
                    break

        for idx, h in enumerate(headers):
            if not h:
                continue
            if cols["ps"] is None and (h == "ps" or re.search(r'\bps\b', h)):
                cols["ps"] = idx
            elif cols["schet_n"] is None and any(k in h for k in ['сч/н', 'сч\\н', 'счет н', 'счёт н']):
                cols["schet_n"] = idx
            elif cols["schet_pp"] is None and is_head and any(k in h for k in ['счет пп', 'счёт пп', 'счет', 'счёт']):
                if 'по счету' not in h and 'по счёту' not in h and 'сч/н' not in h:
                    cols["schet_pp"] = idx
            elif cols["qty"] is None and any(k in h for k in ['по счету', 'по счёту', 'кол-во', 'количество', 'шт']):
                cols["qty"] = idx
            elif cols["deliv"] is None and any(k in h for k in ['доставк', 'способ', 'отправк']):
                cols["deliv"] = idx

    # Если шапки нет или колонки не подписаны — определяем их по данным
    data_sample = df.iloc[start_data_row : min(start_data_row + 120, num_rows)]
    if not data_sample.empty:
        if cols["ps"] is None:
            for c_idx in range(num_cols):
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                if len(vals) > 0 and (vals.str.match(r'^\d{6}$').mean() > 0.3):
                    cols["ps"] = c_idx
                    break

    if is_head:
        if cols["ps"] is None:
            cols["ps"] = min(2, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(4, num_cols - 1)
        if cols["deliv"] is None and num_cols > 14:
            cols["deliv"] = 14
    else:
        if cols["ps"] is None:
            cols["ps"] = min(4, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(7, num_cols - 1)
        if cols["deliv"] is None and num_cols > 1:
            cols["deliv"] = 1

    # Проверяем/находим лучшую колонку с названиями фирм по содержимому (чтобы не взять колонку с децимальными номерами плат)
    best_f_col = cols["firma_zakazchik"]
    best_f_score = score_column_as_firma(df, best_f_col, start_data_row) if best_f_col is not None else -100

    for c_idx in range(num_cols):
        if c_idx in (cols["ps"], cols["qty"], cols["schet_pp"], cols["schet_n"]):
            continue
        sc = score_column_as_firma(df, c_idx, start_data_row)
        if sc > best_f_score and sc > 2:
            best_f_score = sc
            best_f_col = c_idx

    cols["firma_zakazchik"] = best_f_col
    return cols

def build_workbook_lookup(excel_path, is_head=True):
    """
    Сканирует все листы книги Excel.
    Извлекает Фирму заказчика, Счет ПП, Количество и Доставку из целевых и соседних ячеек строки.
    """
    xl = pd.ExcelFile(excel_path)
    ps_lookup = {}
    schet_n_lookup = {}
    schet_deliv_lookup = {}
    firma_deliv_history = {}
    seen_ps = set()
    duplicates = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        cols = detect_sheet_columns(df, is_head=is_head)
        if cols is None:
            continue

        start_r = cols["header_row"] + 1 if cols["header_row"] >= 0 else 0
        sub = df.iloc[start_r:].copy().reset_index(drop=True)

        ps_col = cols["ps"]
        fz_col = cols["firma_zakazchik"]
        schet_pp_col = cols["schet_pp"]
        schet_n_col = cols["schet_n"]
        qty_col = cols["qty"]
        deliv_col = cols["deliv"]

        last_schet_n = ""
        last_fz_firma = ""
        last_fz_deliv = ""
        last_schet_pp = ""

        for i in range(len(sub)):
            row_vals = sub.iloc[i].tolist()

            if all(pd.isna(x) or str(x).strip() == '' for x in row_vals):
                last_schet_n = ""
                last_fz_firma = ""
                last_fz_deliv = ""
                last_schet_pp = ""
                continue

            # 1. Ищем Сч/Н
            curr_sn = clean_schet_n(row_vals[schet_n_col]) if (schet_n_col is not None and schet_n_col < len(row_vals)) else ""
            if not curr_sn:
                for c_i in range(min(10, len(row_vals))):
                    if c_i in (ps_col, qty_col, fz_col, schet_pp_col):
                        continue
                    cand_sn = clean_schet_n(row_vals[c_i])
                    if len(cand_sn) == 5:
                        curr_sn = cand_sn
                        break

            if curr_sn and curr_sn != last_schet_n:
                last_schet_n = curr_sn
                last_fz_firma = ""
                last_fz_deliv = ""
                last_schet_pp = ""

            # 2. Ищем Фирму и Доставку в колонке "Фирма заказчик", а если там пусто — во всех ячейках строки
            fz_firma, fz_deliv = "", ""
            if fz_col is not None and fz_col < len(row_vals):
                fz_firma, fz_deliv = extract_firma_and_delivery(row_vals[fz_col])

            if not fz_firma:
                for c_i in range(len(row_vals)):
                    if c_i in (ps_col, qty_col, schet_pp_col, schet_n_col, deliv_col):
                        continue
                    cand_f, cand_d = extract_firma_and_delivery(row_vals[c_i])
                    if cand_f and len(cand_f) >= 3:
                        fz_firma = cand_f
                        if cand_d and not fz_deliv:
                            fz_deliv = cand_d
                        break

            # Проверяем доставку во всех ячейках строки
            row_deliv = fz_deliv
            if not row_deliv and deliv_col is not None and deliv_col < len(row_vals):
                row_deliv = extract_delivery_from_text(row_vals[deliv_col])
            if not row_deliv:
                for c_i in range(len(row_vals)):
                    if c_i in (ps_col, qty_col, schet_pp_col, schet_n_col):
                        continue
                    cand_d = extract_delivery_from_text(row_vals[c_i])
                    if cand_d:
                        row_deliv = cand_d
                        break

            # Наследование внутри объединённых ячеек одного счёта
            if fz_firma:
                last_fz_firma = fz_firma
            elif last_fz_firma:
                fz_firma = last_fz_firma

            if row_deliv:
                last_fz_deliv = row_deliv
            elif last_fz_deliv and (curr_sn == last_schet_n or fz_firma == last_fz_firma):
                row_deliv = last_fz_deliv

            # 3. Ищем Счет ПП (4 цифры, например 0611, 0682)
            curr_schet_pp = ""
            if is_head:
                raw_schet = ""
                if schet_pp_col is not None and schet_pp_col < len(row_vals) and pd.notna(row_vals[schet_pp_col]):
                    raw_schet = str(row_vals[schet_pp_col]).strip()
                if not raw_schet:
                    for c_i in range(len(row_vals)):
                        if c_i in (ps_col, qty_col, deliv_col, fz_col, schet_n_col):
                            continue
                        v = str(row_vals[c_i]).strip() if pd.notna(row_vals[c_i]) else ""
                        if v.endswith('.0'):
                            v = v[:-2]
                        if re.match(r'^0\d{3}$', v):
                            raw_schet = v
                            break
                if raw_schet.endswith('.0'):
                    raw_schet = raw_schet[:-2]
                digits = re.sub(r'\D', '', raw_schet)
                if digits and len(digits) <= 4:
                    curr_schet_pp = digits.zfill(4)
                    last_schet_pp = curr_schet_pp
                elif last_schet_pp and (curr_sn == last_schet_n or fz_firma == last_fz_firma):
                    curr_schet_pp = last_schet_pp

            # Запоминаем историю по фирме и по номеру счёта
            if fz_firma and row_deliv:
                f_key = normalize_company_key(fz_firma)
                if f_key:
                    firma_deliv_history[f_key] = row_deliv

            if curr_sn:
                sn_entry = schet_n_lookup.get(curr_sn, {"firma": "", "schet_pp": "", "deliv": ""})
                if fz_firma:
                    sn_entry["firma"] = fz_firma
                if curr_schet_pp:
                    sn_entry["schet_pp"] = curr_schet_pp
                if row_deliv:
                    sn_entry["deliv"] = row_deliv
                    schet_deliv_lookup[curr_sn] = row_deliv
                schet_n_lookup[curr_sn] = sn_entry

            if curr_schet_pp and row_deliv:
                schet_deliv_lookup[curr_schet_pp] = row_deliv

            if ps_col is None or ps_col >= len(row_vals):
                continue

            ps = clean_ps(row_vals[ps_col])
            if not ps:
                continue

            if ps in seen_ps and ps not in duplicates:
                duplicates.append(ps)
            seen_ps.add(ps)

            entry = ps_lookup.get(ps, {"firma": "", "schet_pp": "", "schet_n": curr_sn, "qty": np.nan, "deliv": ""})

            if fz_firma:
                entry["firma"] = fz_firma
            if curr_sn:
                entry["schet_n"] = curr_sn
            if curr_schet_pp:
                entry["schet_pp"] = curr_schet_pp

            if qty_col is not None and qty_col < len(row_vals) and pd.notna(row_vals[qty_col]):
                q_val = row_vals[qty_col]
                if str(q_val).strip() != '':
                    entry["qty"] = q_val

            if row_deliv:
                entry["deliv"] = row_deliv

            ps_lookup[ps] = entry

    # Второй проход: заполняем пропуски в PS из справочника Сч/Н и истории фирм
    for ps, entry in ps_lookup.items():
        sn = entry.get("schet_n", "")
        if sn and sn in schet_n_lookup:
            if not entry["firma"] and schet_n_lookup[sn]["firma"]:
                entry["firma"] = schet_n_lookup[sn]["firma"]
            if not entry["schet_pp"] and schet_n_lookup[sn]["schet_pp"]:
                entry["schet_pp"] = schet_n_lookup[sn]["schet_pp"]
            if not entry["deliv"] and schet_n_lookup[sn]["deliv"]:
                entry["deliv"] = schet_n_lookup[sn]["deliv"]

        if not entry["deliv"]:
            if entry.get("schet_pp") and entry["schet_pp"] in schet_deliv_lookup:
                entry["deliv"] = schet_deliv_lookup[entry["schet_pp"]]
            elif sn and sn in schet_deliv_lookup:
                entry["deliv"] = schet_deliv_lookup[sn]
            elif entry.get("firma"):
                f_key = normalize_company_key(entry["firma"])
                if f_key in firma_deliv_history:
                    entry["deliv"] = firma_deliv_history[f_key]

    return ps_lookup, schet_n_lookup, schet_deliv_lookup, firma_deliv_history, duplicates

def find_last_block_start(df_raw):
    last_sep = 0
    for idx in range(len(df_raw) - 1, -1, -1):
        row_vals = [str(x).strip().upper() for x in df_raw.iloc[idx].tolist() if pd.notna(x)]
        row_str = " ".join(row_vals)
        if "ОТПРАВКА" in row_str or "ПРИХОД" in row_str:
            last_sep = idx
            break
        if "ФИРМА" in row_vals and "PS" in row_vals:
            last_sep = idx
            break
    return last_sep

def build_posylka_full_history(df_raw, end_idx, idx_firma, idx_ps, idx_schet_pp, idx_schet_n, idx_deliv):
    """
    Читает все предыдущие (верхние) отправки в самом файле Посылка!
    Запоминает по каждой Фирме, по каждому Сч/Н и по каждому PS:
    - Название фирмы
    - Счет ПП (например, 0611 для Сч/Н 20052 ВНИИА)
    - Доставку (ДМ, СДЭК, К, Самовывоз)
    """
    firma_to_deliv = {}
    firma_to_schet_pp = {}
    sn_to_info = {}
    ps_to_info = {}

    for idx in range(1, end_idx):
        row_vals = df_raw.iloc[idx].tolist()
        f_raw = str(row_vals[idx_firma]).strip() if (idx_firma < len(row_vals) and pd.notna(row_vals[idx_firma])) else ""
        f_clean, f_deliv = extract_firma_and_delivery(f_raw)
        d_col = extract_delivery_from_text(row_vals[idx_deliv]) if idx_deliv < len(row_vals) else ""
        d_final = d_col or f_deliv

        ps_val = clean_ps(row_vals[idx_ps]) if idx_ps < len(row_vals) else ""
        sn_val = clean_schet_n(row_vals[idx_schet_n]) if idx_schet_n < len(row_vals) else ""

        spp_raw = str(row_vals[idx_schet_pp]).strip() if (idx_schet_pp < len(row_vals) and pd.notna(row_vals[idx_schet_pp])) else ""
        if spp_raw.endswith('.0'):
            spp_raw = spp_raw[:-2]
        spp_digits = re.sub(r'\D', '', spp_raw)
        spp_val = spp_digits.zfill(4) if (spp_digits and len(spp_digits) <= 4) else ""

        if f_clean:
            f_key = normalize_company_key(f_clean)
            if f_key:
                if d_final:
                    firma_to_deliv[f_key] = d_final
                if spp_val:
                    firma_to_schet_pp[f_key] = spp_val

        if sn_val:
            info = sn_to_info.get(sn_val, {"firma": "", "schet_pp": "", "deliv": ""})
            if f_clean:
                info["firma"] = f_clean
            if spp_val:
                info["schet_pp"] = spp_val
            if d_final:
                info["deliv"] = d_final
            sn_to_info[sn_val] = info

        if ps_val:
            info = ps_to_info.get(ps_val, {"firma": "", "schet_pp": "", "deliv": ""})
            if f_clean:
                info["firma"] = f_clean
            if spp_val:
                info["schet_pp"] = spp_val
            if d_final:
                info["deliv"] = d_final
            ps_to_info[ps_val] = info

    return firma_to_deliv, firma_to_schet_pp, sn_to_info, ps_to_info

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение файлов и поиск незаполненной посылки внизу...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Анализ Сводки ПП 2026 и Сводки Филиал 2025...")

    pp_dict, pp_sn_lookup, pp_schet_deliv, pp_firma_hist, pp_duplicates = build_workbook_lookup(real_pp, is_head=True)
    filial_dict, filial_sn_lookup, filial_schet_deliv, filial_firma_hist, filial_duplicates = build_workbook_lookup(real_fil, is_head=False)

    header_row = df_pos_raw.iloc[0].tolist()
    col_map = {}
    for c_idx, val in enumerate(header_row):
        name = str(val).strip()
        col_map[name] = c_idx

    idx_firma = col_map.get('Фирма', 0)
    idx_ps = col_map.get('PS', 1)
    idx_schet_pp = col_map.get('Счет пп', 2)
    idx_schet_n = col_map.get('Сч/Н', 3)
    idx_qty = col_map.get('По счету', 4)
    idx_otk = col_map.get('Этикетка', 7)
    idx_deliv = col_map.get('Доставка', 10)

    start_idx = find_last_block_start(df_pos_raw)
    pos_f_deliv, pos_f_spp, pos_sn_info, pos_ps_info = build_posylka_full_history(
        df_pos_raw, start_idx, idx_firma, idx_ps, idx_schet_pp, idx_schet_n, idx_deliv
    )

    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True)
    df_block = df_block.astype(object)
    num_block_cols = df_block.shape[1]

    if progress_callback:
        progress_callback(0.6, "Заполнение Печатных плат и Филиала...")

    stats = {
        "total_rows": len(df_block),
        "empty_ps_rows": 0,
        "already_filled": 0,
        "matched_head": 0,
        "matched_filial": 0,
        "not_found": 0,
        "duplicates_in_summaries": len(pp_duplicates) + len(filial_duplicates),
        "conflicts_both_sources": 0,
        "warnings": [],
        "row_statuses": {},
        "header_row": header_row
    }

    if pp_duplicates:
        stats["warnings"].append(f"В Сводке ПП есть дубликаты PS ({len(pp_duplicates)} шт.): {', '.join(pp_duplicates[:5])}")
    if filial_duplicates:
        stats["warnings"].append(f"В Сводке Филиала есть дубликаты PS ({len(filial_duplicates)} шт.): {', '.join(filial_duplicates[:5])}")

    current_zone = "head"
    row_meta = []

    for idx in range(len(df_block)):
        row_vals = df_block.iloc[idx].tolist()
        first_cell = str(row_vals[0]).strip().upper() if len(row_vals) > 0 and pd.notna(row_vals[0]) else ""
        row_joined = " ".join([str(x).upper() for x in row_vals if pd.notna(x)])

        if "ОТПРАВКА" in row_joined or "ПРИХОД" in row_joined:
            row_meta.append({"idx": idx, "type": "banner_date", "zone": current_zone, "status": "banner_date", "deliv": ""})
            continue
        if first_cell == "ФИЛИАЛ":
            current_zone = "filial"
            row_meta.append({"idx": idx, "type": "banner_filial", "zone": current_zone, "status": "banner_filial", "deliv": ""})
            continue
        if first_cell == "ФИРМА":
            row_meta.append({"idx": idx, "type": "header", "zone": current_zone, "status": "header", "deliv": ""})
            continue

        ps = clean_ps(row_vals[idx_ps]) if idx_ps < len(row_vals) else ""
        schet_n_val = clean_schet_n(row_vals[idx_schet_n]) if idx_schet_n < len(row_vals) else ""

        if not ps:
            stats["empty_ps_rows"] += 1
            row_meta.append({"idx": idx, "type": "empty", "zone": current_zone, "status": "empty", "deliv": ""})
            continue

        in_head = ps in pp_dict
        in_filial = ps in filial_dict

        if in_head and in_filial:
            stats["conflicts_both_sources"] += 1

        # Если в самой ячейке Фирма в Посылке была дописана доставка, извлекаем её
        if idx_firma < num_block_cols and pd.notna(row_vals[idx_firma]):
            f_clean_existing, f_deliv_existing = extract_firma_and_delivery(row_vals[idx_firma])
            if f_deliv_existing and idx_deliv < num_block_cols:
                df_block.iat[idx, idx_deliv] = f_deliv_existing

        if idx_deliv < num_block_cols:
            existing_deliv = extract_delivery_from_text(df_block.iat[idx, idx_deliv])
            if existing_deliv:
                df_block.iat[idx, idx_deliv] = existing_deliv

        def is_cell_empty(col_i):
            if col_i >= num_block_cols:
                return False
            v = df_block.iat[idx, col_i]
            return pd.isna(v) or str(v).strip() == ''

        if idx_firma < num_block_cols and not is_cell_empty(idx_firma):
            if is_board_code_or_phone(df_block.iat[idx, idx_firma]):
                df_block.iat[idx, idx_firma] = np.nan

        # Блок 1: Печатные платы
        if current_zone == "head":
            info = pp_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            sn_fallback = pp_sn_lookup.get(schet_n_val, {"firma": "", "schet_pp": "", "deliv": ""})
            pos_sn = pos_sn_info.get(schet_n_val, {"firma": "", "schet_pp": "", "deliv": ""})
            pos_ps = pos_ps_info.get(ps, {"firma": "", "schet_pp": "", "deliv": ""})

            firma_best = info["firma"] or sn_fallback["firma"] or pos_ps["firma"] or pos_sn["firma"]
            if is_cell_empty(idx_firma) and firma_best:
                df_block.iat[idx, idx_firma] = firma_best
            elif firma_best and is_board_code_or_phone(df_block.iat[idx, idx_firma]):
                df_block.iat[idx, idx_firma] = firma_best

            firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
            f_key = normalize_company_key(firma_now)

            schet_pp_best = (
                info["schet_pp"]
                or sn_fallback["schet_pp"]
                or pos_ps["schet_pp"]
                or pos_sn["schet_pp"]
                or pos_f_spp.get(f_key, "")
            )
            if is_cell_empty(idx_schet_pp) and schet_pp_best:
                df_block.iat[idx, idx_schet_pp] = schet_pp_best

            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]

            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"

            if is_cell_empty(idx_deliv):
                d_val = (
                    info["deliv"]
                    or sn_fallback["deliv"]
                    or pp_schet_deliv.get(schet_pp_best, "")
                    or pp_schet_deliv.get(schet_n_val, "")
                    or pp_firma_hist.get(f_key, "")
                    or pos_ps["deliv"]
                    or pos_sn["deliv"]
                    or pos_f_deliv.get(f_key, "")
                )
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if in_head or firma_best:
                stats["matched_head"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            if final_d.upper() == "САМОВЫВОЗ":
                st = "samovyvoz_head"
            elif not in_head and is_cell_empty(idx_qty):
                st = "not_found"
            else:
                st = "head"
            row_meta.append({"idx": idx, "type": "data", "zone": "head", "status": st, "deliv": final_d})

        # Блок 2: Филиал
        else:
            info = filial_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            sn_fallback = filial_sn_lookup.get(schet_n_val, {"firma": "", "schet_pp": "", "deliv": ""})
            pos_sn = pos_sn_info.get(schet_n_val, {"firma": "", "schet_pp": "", "deliv": ""})
            pos_ps = pos_ps_info.get(ps, {"firma": "", "schet_pp": "", "deliv": ""})

            if is_cell_empty(idx_firma):
                f_best = info["firma"] or sn_fallback["firma"] or pos_ps["firma"] or pos_sn["firma"]
                if f_best:
                    df_block.iat[idx, idx_firma] = f_best

            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]

            if is_cell_empty(idx_deliv):
                firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
                f_key = normalize_company_key(firma_now)
                d_val = (
                    info["deliv"]
                    or sn_fallback["deliv"]
                    or filial_schet_deliv.get(schet_n_val, "")
                    or filial_firma_hist.get(f_key, "")
                    or pos_ps["deliv"]
                    or pos_sn["deliv"]
                    or pos_f_deliv.get(f_key, "")
                    or pp_firma_hist.get(f_key, "")
                )
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if in_filial:
                stats["matched_filial"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            if final_d.upper() == "САМОВЫВОЗ":
                st = "samovyvoz_filial"
            elif not in_filial and is_cell_empty(idx_qty):
                st = "not_found"
            else:
                st = "filial"
            row_meta.append({"idx": idx, "type": "data", "zone": "filial", "status": st, "deliv": final_d})

    if progress_callback:
        progress_callback(0.85, "Сортировка строк: ДМ → СДЭК → К → Самовывоз...")

    sorted_rows = []
    sorted_statuses = {}
    new_idx = 0

    head_data = []
    filial_data = []
    seen_filial_banner = False

    for meta in row_meta:
        orig_idx = meta["idx"]
        if meta["type"] in ("banner_date", "header"):
            sorted_rows.append(df_block.iloc[orig_idx].tolist())
            sorted_statuses[new_idx] = meta["status"]
            new_idx += 1
        elif meta["type"] == "banner_filial":
            head_data.sort(key=lambda item: delivery_sort_rank(item[1]["deliv"]))
            for r_list, m in head_data:
                sorted_rows.append(r_list)
                sorted_statuses[new_idx] = m["status"]
                new_idx += 1
            head_data = []

            sorted_rows.append(df_block.iloc[orig_idx].tolist())
            sorted_statuses[new_idx] = "banner_filial"
            new_idx += 1
            seen_filial_banner = True
        elif meta["type"] == "data":
            if not seen_filial_banner:
                head_data.append((df_block.iloc[orig_idx].tolist(), meta))
            else:
                filial_data.append((df_block.iloc[orig_idx].tolist(), meta))

    if head_data:
        head_data.sort(key=lambda item: delivery_sort_rank(item[1]["deliv"]))
        for r_list, m in head_data:
            sorted_rows.append(r_list)
            sorted_statuses[new_idx] = m["status"]
            new_idx += 1

    if filial_data:
        filial_data.sort(key=lambda item: delivery_sort_rank(item[1]["deliv"]))
        for r_list, m in filial_data:
            sorted_rows.append(r_list)
            sorted_statuses[new_idx] = m["status"]
            new_idx += 1

    df_result = pd.DataFrame(sorted_rows)
    clean_cols = [str(c) if pd.notna(c) else f"Колонка_{i+1}" for i, c in enumerate(header_row)]
    df_result.columns = clean_cols[:df_result.shape[1]]
    stats["row_statuses"] = sorted_statuses

    if progress_callback:
        progress_callback(1.0, "Последняя посылка успешно сведена!")

    return df_result, stats
