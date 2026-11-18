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

def is_phone_or_number(text):
    s = str(text).strip()
    if not s:
        return True
    # Если строка состоит в основном из цифр, пробелов, скобок, плюсов и дефисов — это телефон или код, а не фирма
    digits_count = sum(c.isdigit() for c in s)
    letters_count = sum(c.isalpha() for c in s)
    if letters_count < 2 and digits_count >= 5:
        return True
    if re.match(r'^(\+7|8|7)[\s\-\(]?\d{3}', s):
        return True
    return False

def normalize_delivery(raw_val):
    if raw_val is None or (isinstance(raw_val, float) and np.isnan(raw_val)):
        return ""
    s = str(raw_val).strip().upper()
    if not s or s in ('NAN', 'NONE', 'NAT'):
        return ""
    if "САМОВЫВОЗ" in s or "САМОВ" in s:
        return "Самовывоз"
    if "СДЭК" in s or "CDEK" in s or "СДЕК" in s:
        return "СДЭК"
    if "ДМ" in s or "ДЕЛОВ" in s:
        return "ДМ"
    if s == "К" or "КУРЬЕР" in s or re.search(r'(^|[\s,;/\(])К($|[\s,;/\)\.])', s):
        return "К"
    return ""

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
            if any(k in cell for k in ['заказчик', 'фирма', 'клиент', 'контрагент', 'покупатель']):
                score += 3
            if any(k in cell for k in ['счет', 'счёт', 'кол', 'доставк', 'отправ', 'сч/н']):
                score += 2
        if score > best_score:
            best_score = score
            header_row_idx = r_idx

    raw_headers = df.iloc[header_row_idx].tolist()
    headers = [str(x).strip().lower() if pd.notna(x) else "" for x in raw_headers]
    cols = {"header_row": header_row_idx, "ps": None, "firma": None, "schet_pp": None, "schet_n": None, "qty": None, "deliv": None}

    for idx, h in enumerate(headers):
        if not h:
            continue
        # Исключаем телефонные и контактные колонки из поиска фирмы
        if any(bad in h for bad in ['тел', 'phone', 'контакт', 'примеч', 'коммент', 'адрес']):
            continue

        if cols["ps"] is None and (h == "ps" or re.search(r'\bps\b', h)):
            cols["ps"] = idx
        elif cols["firma"] is None and any(k in h for k in ['заказчик', 'фирма', 'клиент', 'организац', 'покупатель', 'наименование']):
            cols["firma"] = idx
        elif cols["schet_n"] is None and any(k in h for k in ['сч/н', 'сч\\н', 'счет н', 'счёт н']):
            cols["schet_n"] = idx
        elif cols["schet_pp"] is None and is_head and any(k in h for k in ['счет пп', 'счёт пп', '№ счета', 'номер счета', 'счет', 'счёт']):
            if 'по счету' not in h and 'по счёту' not in h:
                cols["schet_pp"] = idx
        elif cols["qty"] is None and any(k in h for k in ['по счету', 'по счёту', 'кол-во', 'количество', 'шт']):
            cols["qty"] = idx
        elif cols["deliv"] is None and any(k in h for k in ['доставк', 'способ', 'отправк', 'транспорт', 'перевоз']):
            cols["deliv"] = idx

    data_sample = df.iloc[header_row_idx + 1 : min(header_row_idx + 100, num_rows)]
    if not data_sample.empty:
        # Автопоиск колонки PS по 6-значным числам
        if cols["ps"] is None:
            for c_idx in range(num_cols):
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                if len(vals) > 0 and (vals.str.match(r'^\d{6}$').mean() > 0.35):
                    cols["ps"] = c_idx
                    break

        # Автопоиск колонки Фирма (текстовая колонка с названиями компаний, не телефон)
        if cols["firma"] is None:
            for c_idx in range(min(8, num_cols)):
                if c_idx in (cols["ps"], cols["qty"], cols["deliv"], cols["schet_pp"]):
                    continue
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.strip()
                if len(vals) > 0:
                    text_ratio = vals.apply(lambda v: not is_phone_or_number(v) and len(v) >= 3).mean()
                    if text_ratio > 0.5:
                        cols["firma"] = c_idx
                        break

        # Автопоиск колонки Доставка
        if cols["deliv"] is None:
            for c_idx in range(num_cols):
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.upper()
                if len(vals) > 0:
                    hits = vals.apply(lambda v: normalize_delivery(v) != "").mean()
                    if hits > 0.12:
                        cols["deliv"] = c_idx
                        break

    if is_head:
        if cols["ps"] is None:
            cols["ps"] = min(2, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(4, num_cols - 1)
        if cols["deliv"] is None:
            cols["deliv"] = min(14, num_cols - 1)
    else:
        if cols["ps"] is None:
            cols["ps"] = min(4, num_cols - 1)
        if cols["qty"] is None:
            cols["qty"] = min(7, num_cols - 1)
        if cols["deliv"] is None:
            cols["deliv"] = min(1, num_cols - 1)

    return cols

def build_workbook_lookup(excel_path, is_head=True):
    """Читает все листы книги Excel и собирает справочники по PS и по номеру счета (Сч/Н)."""
    xl = pd.ExcelFile(excel_path)
    ps_lookup = {}
    schet_deliv_lookup = {}
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
        firma_col = cols["firma"]
        schet_pp_col = cols["schet_pp"]
        schet_n_col = cols["schet_n"]
        qty_col = cols["qty"]
        deliv_col = cols["deliv"]

        for i in range(len(sub)):
            row_vals = sub.iloc[i].tolist()
            if ps_col is None or ps_col >= len(row_vals):
                continue

            ps = clean_ps(row_vals[ps_col])
            if not ps:
                continue

            if ps in seen_ps and ps not in duplicates:
                duplicates.append(ps)
            seen_ps.add(ps)

            entry = ps_lookup.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})

            # 1. Фирма (строго не телефон!)
            if firma_col is not None and firma_col < len(row_vals) and pd.notna(row_vals[firma_col]):
                f_val = str(row_vals[firma_col]).strip()
                if f_val and f_val.lower() != 'nan' and not is_phone_or_number(f_val):
                    entry["firma"] = f_val
            if not entry["firma"]:
                # Запасной поиск названия фирмы в соседних текстовых ячейках строки
                for c_i in range(min(10, len(row_vals))):
                    if c_i in (ps_col, qty_col, deliv_col, schet_pp_col, schet_n_col):
                        continue
                    cand_f = str(row_vals[c_i]).strip() if pd.notna(row_vals[c_i]) else ""
                    if cand_f and cand_f.lower() != 'nan' and not is_phone_or_number(cand_f) and normalize_delivery(cand_f) == "" and len(cand_f) >= 3:
                        entry["firma"] = cand_f
                        break

            # 2. Счет ПП (4 цифры с ведущим нулем, например 0682)
            if is_head:
                raw_schet = ""
                if schet_pp_col is not None and schet_pp_col < len(row_vals) and pd.notna(row_vals[schet_pp_col]):
                    raw_schet = str(row_vals[schet_pp_col]).strip()
                if not raw_schet:
                    for c_i in range(min(12, len(row_vals))):
                        if c_i in (ps_col, qty_col, deliv_col, firma_col, schet_n_col):
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

            # 3. Количество по счету
            if qty_col is not None and qty_col < len(row_vals) and pd.notna(row_vals[qty_col]):
                q_val = row_vals[qty_col]
                if str(q_val).strip() != '':
                    entry["qty"] = q_val

            # 4. Доставка (ищем в колонке доставки или в любой ячейке строки)
            deliv_found = ""
            if deliv_col is not None and deliv_col < len(row_vals) and pd.notna(row_vals[deliv_col]):
                deliv_found = normalize_delivery(row_vals[deliv_col])
            if not deliv_found:
                for c_i in range(len(row_vals)):
                    if c_i in (ps_col, firma_col):
                        continue
                    cand = normalize_delivery(row_vals[c_i])
                    if cand:
                        deliv_found = cand
                        break
            if deliv_found:
                entry["deliv"] = deliv_found
                # Запоминаем доставку также по номеру счета Сч/Н на случай объединенных строк в сводке
                if schet_n_col is not None and schet_n_col < len(row_vals):
                    sn = clean_schet_n(row_vals[schet_n_col])
                    if sn:
                        schet_deliv_lookup[sn] = deliv_found

            ps_lookup[ps] = entry

    return ps_lookup, schet_deliv_lookup, duplicates

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

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение файлов и поиск незаполненной посылки внизу...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Сканирование всех листов Сводки ПП и Сводки Филиала...")

    pp_dict, pp_schet_deliv, pp_duplicates = build_workbook_lookup(real_pp, is_head=True)
    filial_dict, filial_schet_deliv, filial_duplicates = build_workbook_lookup(real_fil, is_head=False)

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
    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True)
    df_block = df_block.astype(object)
    num_block_cols = df_block.shape[1]

    if progress_callback:
        progress_callback(0.6, "Заполнение данных Печатных плат и Филиала...")

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
            existing_deliv = normalize_delivery(row_vals[idx_deliv])
            if existing_deliv:
                df_block.iat[idx, idx_deliv] = existing_deliv

        def is_cell_empty(col_i):
            if col_i >= num_block_cols:
                return False
            v = df_block.iat[idx, col_i]
            return pd.isna(v) or str(v).strip() == ''

        # Если в поле Фирма случайно стоял телефон, очищаем его для замены на реальное название компании
        if idx_firma < num_block_cols and not is_cell_empty(idx_firma):
            if is_phone_or_number(df_block.iat[idx, idx_firma]):
                df_block.iat[idx, idx_firma] = np.nan

        if current_zone == "head" and (in_head or not in_filial):
            info = pp_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_schet_pp) and info["schet_pp"]:
                df_block.iat[idx, idx_schet_pp] = info["schet_pp"]
            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"
            if is_cell_empty(idx_deliv):
                d_val = info["deliv"] or pp_schet_deliv.get(schet_n_val, "")
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

        elif in_filial or current_zone == "filial":
            info = filial_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_deliv):
                d_val = info["deliv"] or filial_schet_deliv.get(schet_n_val, "")
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
