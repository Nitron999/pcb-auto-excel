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

def normalize_company_key(name):
    """Нормализует название компании для поиска совпадений по предыдущим заказам."""
    s = str(name or "").strip().lower()
    # Убираем пометки доставки и скобки для точного сопоставления имени фирмы
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\b(дм|сдэк|cdek|самовывоз|курьер|к|ооо|ао|зао|пао|ип|фгуп|нпц)\b', '', s)
    s = re.sub(r'[^а-яa-z0-9]', '', s)
    return s

def extract_firma_and_delivery(raw_text):
    """
    Разбирает ячейку из колонки 'Фирма заказчик':
    извлекает чистое название фирмы и способ доставки (ДМ, СДЭК, К, Самовывоз).
    """
    if raw_text is None or (isinstance(raw_text, float) and np.isnan(raw_text)):
        return "", ""
    s = str(raw_text).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return "", ""

    u = s.upper()
    deliv = ""

    if "САМОВЫВОЗ" in u or "САМОВ" in u:
        deliv = "Самовывоз"
    elif "СДЭК" in u or "CDEK" in u or "СДЕК" in u:
        deliv = "СДЭК"
    elif re.search(r'(^|[\s,;/\(\-\.])ДМ($|[\s,;/\)\-\.])', u) or "ДЕЛОВ" in u:
        deliv = "ДМ"
    elif u == "К" or "КУРЬЕР" in u or re.search(r'(^|[\s,;/\(\-\.])К($|[\s,;/\)\-\.])', u):
        deliv = "К"

    # Очищаем название фирмы от приписанного в конце/скобках способа доставки
    clean_name = re.sub(r'[\(\[\s,;\-\/]*(самовывоз|сдэк|cdek|сдек|курьер|дм)[\)\]\s,;\-\/\.]*$', '', s, flags=re.IGNORECASE).strip()
    clean_name = re.sub(r'\s+К$', '', clean_name).strip()

    # Если вся ячейка состояла только из слова "ДМ" / "СДЭК" / "Самовывоз" / "К", то это не название фирмы
    if clean_name.upper() in ("ДМ", "СДЭК", "CDEK", "К", "САМОВЫВОЗ", "КУРЬЕР", ""):
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

def detect_sheet_columns(df, is_head=True):
    num_rows, num_cols = df.shape
    if num_rows == 0 or num_cols == 0:
        return None

    header_row_idx = 0
    best_score = -1

    for r_idx in range(min(15, num_rows)):
        row_vals = df.iloc[r_idx].tolist()
        row_strs = [str(x).strip().lower() for x in row_vals if pd.notna(x)]
        score = 0
        for cell in row_strs:
            if cell == 'ps' or re.search(r'\bps\b', cell):
                score += 4
            if 'фирма заказчик' in cell or 'заказчик' in cell or 'фирма' in cell:
                score += 6
            if any(k in cell for k in ['счет', 'счёт', 'кол', 'доставк', 'сч/н']):
                score += 2
        if score > best_score:
            best_score = score
            header_row_idx = r_idx

    raw_headers = df.iloc[header_row_idx].tolist()
    headers = [str(x).strip().lower() if pd.notna(x) else "" for x in raw_headers]
    cols = {
        "header_row": header_row_idx,
        "ps": None,
        "firma_zakazchik": None,
        "schet_pp": None,
        "schet_n": None,
        "qty": None,
        "deliv": None
    }

    # 1. Приоритет: точный поиск колонки "Фирма заказчик"
    for idx, h in enumerate(headers):
        if 'фирма заказчик' in h or 'фирма-заказчик' in h:
            cols["firma_zakazchik"] = idx
            break
    if cols["firma_zakazchik"] is None:
        for idx, h in enumerate(headers):
            if any(bad in h for bad in ['тел', 'контакт', 'примеч', 'адрес', 'чертеж', 'децим', 'плата', 'изделие']):
                continue
            if 'заказчик' in h or h == 'фирма' or 'клиент' in h or 'покупатель' in h:
                cols["firma_zakazchik"] = idx
                break

    # 2. Остальные колонки
    for idx, h in enumerate(headers):
        if not h:
            continue
        if cols["ps"] is None and (h == "ps" or re.search(r'\bps\b', h)):
            cols["ps"] = idx
        elif cols["schet_n"] is None and any(k in h for k in ['сч/н', 'сч\\н', 'счет н', 'счёт н']):
            cols["schet_n"] = idx
        elif cols["schet_pp"] is None and is_head and any(k in h for k in ['счет пп', 'счёт пп', '№ счета', 'номер счета', 'счет', 'счёт']):
            if 'по счету' not in h and 'по счёту' not in h:
                cols["schet_pp"] = idx
        elif cols["qty"] is None and any(k in h for k in ['по счету', 'по счёту', 'кол-во', 'количество', 'шт']):
            cols["qty"] = idx
        elif cols["deliv"] is None and any(k in h for k in ['доставк', 'способ', 'отправк']):
            cols["deliv"] = idx

    data_sample = df.iloc[header_row_idx + 1 : min(header_row_idx + 100, num_rows)]
    if not data_sample.empty and cols["ps"] is None:
        for c_idx in range(num_cols):
            vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
            if len(vals) > 0 and (vals.str.match(r'^\d{6}$').mean() > 0.35):
                cols["ps"] = c_idx
                break

    if is_head:
        if cols["ps"] is None:
            cols["ps"] = min(2, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(4, num_cols - 1)
    else:
        if cols["ps"] is None:
            cols["ps"] = min(4, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(7, num_cols - 1)

    return cols

def build_workbook_lookup(excel_path, is_head=True):
    """
    Читает все листы сводки.
    Берёт название фирмы и доставку из колонки 'Фирма заказчик',
    а также запоминает историю доставок по номеру счёта и по названию фирмы.
    """
    xl = pd.ExcelFile(excel_path)
    ps_lookup = {}
    schet_deliv_lookup = {}
    firma_deliv_history = {}
    seen_ps = set()
    duplicates = []

    for sheet_name in xl.sheet_names:
        df = xl.parse(sheet_name, header=None)
        cols = detect_sheet_columns(df, is_head=is_head)
        if cols is None:
            continue

        start_r = cols["header_row"] + 1
        sub = df.iloc[start_r:].copy().reset_index(drop=True)

        ps_col = cols["ps"]
        fz_col = cols["firma_zakazchik"]
        schet_pp_col = cols["schet_pp"]
        schet_n_col = cols["schet_n"]
        qty_col = cols["qty"]
        deliv_col = cols["deliv"]

        for i in range(len(sub)):
            row_vals = sub.iloc[i].tolist()

            # Извлекаем данные из колонки "Фирма заказчик" даже если в этой строке нет PS (для накопления истории по фирме/счёту)
            fz_firma, fz_deliv = "", ""
            if fz_col is not None and fz_col < len(row_vals):
                fz_firma, fz_deliv = extract_firma_and_delivery(row_vals[fz_col])

            # Проверяем также отдельную колонку доставки, если она есть
            col_deliv_val = ""
            if deliv_col is not None and deliv_col < len(row_vals):
                _, col_deliv_val = extract_firma_and_delivery(row_vals[deliv_col])

            final_row_deliv = fz_deliv or col_deliv_val

            # Запоминаем доставку по названию фирмы (чтобы подставлять, если у нового заказа доставка не подписана)
            if fz_firma and final_row_deliv:
                f_key = normalize_company_key(fz_firma)
                if f_key:
                    firma_deliv_history[f_key] = final_row_deliv

            # Запоминаем доставку по номеру счёта
            sn_val = clean_schet_n(row_vals[schet_n_col]) if (schet_n_col is not None and schet_n_col < len(row_vals)) else ""
            if sn_val and final_row_deliv:
                schet_deliv_lookup[sn_val] = final_row_deliv

            if ps_col is None or ps_col >= len(row_vals):
                continue

            ps = clean_ps(row_vals[ps_col])
            if not ps:
                continue

            if ps in seen_ps and ps not in duplicates:
                duplicates.append(ps)
            seen_ps.add(ps)

            entry = ps_lookup.get(ps, {"firma": "", "schet_pp": "", "schet_n": sn_val, "qty": np.nan, "deliv": ""})

            if fz_firma:
                entry["firma"] = fz_firma
            if sn_val:
                entry["schet_n"] = sn_val

            # Счет ПП (4 цифры с ведущим нулём, например 0682)
            if is_head:
                raw_schet = ""
                if schet_pp_col is not None and schet_pp_col < len(row_vals) and pd.notna(row_vals[schet_pp_col]):
                    raw_schet = str(row_vals[schet_pp_col]).strip()
                if not raw_schet:
                    for c_i in range(min(12, len(row_vals))):
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
                    entry["schet_pp"] = digits.zfill(4)
                    if final_row_deliv:
                        schet_deliv_lookup[entry["schet_pp"]] = final_row_deliv

            # Количество по счету
            if qty_col is not None and qty_col < len(row_vals) and pd.notna(row_vals[qty_col]):
                q_val = row_vals[qty_col]
                if str(q_val).strip() != '':
                    entry["qty"] = q_val

            if final_row_deliv:
                entry["deliv"] = final_row_deliv

            ps_lookup[ps] = entry

    # Второй проход по собранному справочнику: если у PS не была указана доставка,
    # добираем её по номеру счёта или по истории этой же фирмы в этой сводке
    for ps, entry in ps_lookup.items():
        if not entry["deliv"]:
            if entry.get("schet_pp") and entry["schet_pp"] in schet_deliv_lookup:
                entry["deliv"] = schet_deliv_lookup[entry["schet_pp"]]
            elif entry.get("schet_n") and entry["schet_n"] in schet_deliv_lookup:
                entry["deliv"] = schet_deliv_lookup[entry["schet_n"]]
            elif entry.get("firma"):
                f_key = normalize_company_key(entry["firma"])
                if f_key in firma_deliv_history:
                    entry["deliv"] = firma_deliv_history[f_key]

    return ps_lookup, schet_deliv_lookup, firma_deliv_history, duplicates

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

def build_posylka_history(df_raw, end_idx, idx_firma, idx_deliv):
    """Собирает историю доставок по фирмам из предыдущих (верхних) блоков файла Посылка."""
    hist = {}
    for idx in range(1, end_idx):
        row_vals = df_raw.iloc[idx].tolist()
        if idx_firma >= len(row_vals) or idx_deliv >= len(row_vals):
            continue
        f_raw = str(row_vals[idx_firma]).strip() if pd.notna(row_vals[idx_firma]) else ""
        _, d_norm = extract_firma_and_delivery(row_vals[idx_deliv])
        if f_raw and d_norm:
            f_key = normalize_company_key(f_raw)
            if f_key:
                hist[f_key] = d_norm
    return hist

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение файлов и поиск незаполненной посылки внизу...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Чтение колонки «Фирма заказчик» из Сводки ПП 2026 и Сводки Филиал 2025...")

    pp_dict, pp_schet_deliv, pp_firma_hist, pp_duplicates = build_workbook_lookup(real_pp, is_head=True)
    filial_dict, filial_schet_deliv, filial_firma_hist, filial_duplicates = build_workbook_lookup(real_fil, is_head=False)

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
    posylka_firma_hist = build_posylka_history(df_pos_raw, start_idx, idx_firma, idx_deliv)

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

        if idx_deliv < num_block_cols:
            _, existing_deliv = extract_firma_and_delivery(row_vals[idx_deliv])
            if existing_deliv:
                df_block.iat[idx, idx_deliv] = existing_deliv

        def is_cell_empty(col_i):
            if col_i >= num_block_cols:
                return False
            v = df_block.iat[idx, col_i]
            return pd.isna(v) or str(v).strip() == ''

        # Блок 1: Печатные платы (ориентируемся на Сводку ООО Печатные платы 2026)
        if current_zone == "head":
            info = pp_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            # Всегда ставим правильное название фирмы из колонки "Фирма заказчик", если оно найдено в сводке ПП
            if info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_schet_pp) and info["schet_pp"]:
                df_block.iat[idx, idx_schet_pp] = info["schet_pp"]
            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"

            if is_cell_empty(idx_deliv):
                firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
                f_key = normalize_company_key(firma_now)
                d_val = (
                    info["deliv"]
                    or pp_schet_deliv.get(info.get("schet_pp", ""), "")
                    or pp_schet_deliv.get(schet_n_val, "")
                    or pp_firma_hist.get(f_key, "")
                    or posylka_firma_hist.get(f_key, "")
                )
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if in_head:
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

        # Блок 2: Филиал (ориентируемся на Сводку филиал 2025 -> колонка "Фирма заказчик")
        else:
            info = filial_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]

            if is_cell_empty(idx_deliv):
                firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
                f_key = normalize_company_key(firma_now)
                d_val = (
                    info["deliv"]
                    or filial_schet_deliv.get(schet_n_val, "")
                    or filial_firma_hist.get(f_key, "")
                    or posylka_firma_hist.get(f_key, "")
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
