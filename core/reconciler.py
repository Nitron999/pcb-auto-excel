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
    return digits if len(digits) >= 5 else s

def normalize_delivery(raw_val):
    if raw_val is None or (isinstance(raw_val, float) and np.isnan(raw_val)):
        return ""
    s = str(raw_val).strip().upper()
    if not s or s in ('NAN', 'NONE', 'NAT'):
        return ""
    if "САМОВЫВОЗ" in s:
        return "Самовывоз"
    if "СДЭК" in s or "CDEK" in s:
        return "СДЭК"
    if "ДМ" in s or "ДЕЛОВ" in s:
        return "ДМ"
    if s == "К" or "КУРЬЕР" in s or re.search(r'\bК\b', s):
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

def detect_summary_columns(df, is_head=True):
    num_rows, num_cols = df.shape
    header_row_idx = 0
    best_score = -1

    for r_idx in range(min(15, num_rows)):
        row_vals = df.iloc[r_idx].tolist()
        row_strs = [str(x).strip().lower() for x in row_vals if pd.notna(x)]
        score = 0
        for cell in row_strs:
            if cell == 'ps' or 'ps' in cell:
                score += 3
            if any(k in cell for k in ['заказчик', 'фирма', 'клиент', 'контрагент']):
                score += 2
            if any(k in cell for k in ['счет', 'счёт', 'кол', 'доставк', 'отправ']):
                score += 2
        if score > best_score:
            best_score = score
            header_row_idx = r_idx

    raw_headers = df.iloc[header_row_idx].tolist()
    headers = [str(x).strip().lower() if pd.notna(x) else "" for x in raw_headers]
    cols = {"header_row": header_row_idx, "ps": None, "firma": None, "schet_pp": None, "qty": None, "deliv": None}

    for idx, h in enumerate(headers):
        if not h:
            continue
        if cols["ps"] is None and (h == "ps" or re.search(r'\bps\b', h)):
            cols["ps"] = idx
        elif cols["firma"] is None and any(k in h for k in ['заказчик', 'фирма', 'клиент', 'организац', 'покупатель']):
            cols["firma"] = idx
        elif cols["schet_pp"] is None and is_head and any(k in h for k in ['счет пп', 'счёт пп', '№ счета', 'номер счета', 'счет']):
            if 'по счету' not in h and 'по счёту' not in h:
                cols["schet_pp"] = idx
        elif cols["qty"] is None and any(k in h for k in ['по счету', 'по счёту', 'кол-во', 'количество', 'шт']):
            cols["qty"] = idx
        elif cols["deliv"] is None and any(k in h for k in ['доставк', 'способ', 'отправк', 'транспорт', 'перевоз']):
            cols["deliv"] = idx

    # Безопасный поиск по содержимому через iloc (работает при любом количестве колонок)
    data_sample = df.iloc[header_row_idx + 1 : min(header_row_idx + 80, num_rows)]
    if not data_sample.empty:
        if cols["ps"] is None:
            for c_idx in range(num_cols):
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.strip().str.replace(r'\.0$', '', regex=True)
                if len(vals) > 0 and (vals.str.match(r'^\d{5,7}$').mean() > 0.4):
                    cols["ps"] = c_idx
                    break

        if cols["deliv"] is None:
            for c_idx in range(num_cols):
                vals = data_sample.iloc[:, c_idx].dropna().astype(str).str.upper()
                if len(vals) > 0:
                    hits = vals.apply(lambda v: any(k in v for k in ['ДМ', 'СДЭК', 'САМОВЫВОЗ', 'КУРЬЕР']) or v.strip() == 'К').mean()
                    if hits > 0.15:
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

def build_summary_lookup(df, is_head=True):
    cols = detect_summary_columns(df, is_head=is_head)
    start_r = cols["header_row"] + 1
    sub = df.iloc[start_r:].copy().reset_index(drop=True)
    num_cols = sub.shape[1]

    ps_col = cols["ps"]
    firma_col = cols["firma"]
    schet_col = cols["schet_pp"]
    qty_col = cols["qty"]
    deliv_col = cols["deliv"]

    lookup = {}
    seen_ps = set()
    duplicates = []

    for i in range(len(sub)):
        row_vals = sub.iloc[i].tolist()
        if ps_col is None or ps_col >= len(row_vals):
            continue

        ps = clean_ps(row_vals[ps_col])
        if not re.match(r'^\d{5,7}$', ps):
            continue

        if ps in seen_ps and ps not in duplicates:
            duplicates.append(ps)
        seen_ps.add(ps)

        entry = lookup.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})

        # Фирма (Заказчик)
        if firma_col is not None and firma_col < len(row_vals) and pd.notna(row_vals[firma_col]):
            f_val = str(row_vals[firma_col]).strip()
            if f_val and f_val.lower() != 'nan':
                entry["firma"] = f_val

        # Счет ПП (4 цифры с ведущим нулем)
        if is_head:
            raw_schet = ""
            if schet_col is not None and schet_col < len(row_vals) and pd.notna(row_vals[schet_col]):
                raw_schet = str(row_vals[schet_col]).strip()
            else:
                for c_i in range(min(12, len(row_vals))):
                    if c_i in (ps_col, qty_col, deliv_col, firma_col):
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

        # Количество по счету
        if qty_col is not None and qty_col < len(row_vals) and pd.notna(row_vals[qty_col]):
            q_val = row_vals[qty_col]
            if str(q_val).strip() != '':
                entry["qty"] = q_val

        # Способ доставки (строго ДМ, СДЭК, К, Самовывоз)
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

        lookup[ps] = entry

    return lookup, duplicates

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
    df_pp_raw = pd.read_excel(real_pp, sheet_name=0, header=None)
    df_fil_raw = pd.read_excel(real_fil, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Анализ колонок в Сводке ПП и Сводке Филиала...")

    pp_dict, pp_duplicates = build_summary_lookup(df_pp_raw, is_head=True)
    filial_dict, filial_duplicates = build_summary_lookup(df_fil_raw, is_head=False)

    header_row = df_pos_raw.iloc[0].tolist()
    col_map = {}
    for c_idx, val in enumerate(header_row):
        name = str(val).strip()
        col_map[name] = c_idx

    idx_firma = col_map.get('Фирма', 0)
    idx_ps = col_map.get('PS', 1)
    idx_schet_pp = col_map.get('Счет пп', 2)
    idx_qty = col_map.get('По счету', 4)
    idx_otk = col_map.get('Этикетка', 7)
    idx_deliv = col_map.get('Доставка', 10)

    start_idx = find_last_block_start(df_pos_raw)
    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True)

    # Переводим все столбцы в object для безопасной записи текста и чисел
    df_block = df_block.astype(object)
    num_block_cols = df_block.shape[1]

    if progress_callback:
        progress_callback(0.55, "Заполнение данных Головы (ПП) и Филиала...")

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
        if not ps or not ps.isdigit():
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

        if current_zone == "head" and in_head:
            info = pp_dict[ps]
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_schet_pp) and info["schet_pp"]:
                df_block.iat[idx, idx_schet_pp] = info["schet_pp"]
            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"
            if is_cell_empty(idx_deliv) and info["deliv"]:
                df_block.iat[idx, idx_deliv] = info["deliv"]

            stats["matched_head"] += 1
            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_head" if final_d.upper() == "САМОВЫВОЗ" else "head"
            row_meta.append({"idx": idx, "type": "data", "zone": "head", "status": st, "deliv": final_d})

        elif in_filial:
            info = filial_dict[ps]
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_deliv) and info["deliv"]:
                df_block.iat[idx, idx_deliv] = info["deliv"]

            stats["matched_filial"] += 1
            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_filial" if final_d.upper() == "САМОВЫВОЗ" else "filial"
            row_meta.append({"idx": idx, "type": "data", "zone": current_zone, "status": st, "deliv": final_d})

        elif in_head:
            info = pp_dict[ps]
            if is_cell_empty(idx_firma) and info["firma"]:
                df_block.iat[idx, idx_firma] = info["firma"]
            if is_cell_empty(idx_qty) and pd.notna(info["qty"]):
                df_block.iat[idx, idx_qty] = info["qty"]
            if is_cell_empty(idx_schet_pp) and info["schet_pp"]:
                df_block.iat[idx, idx_schet_pp] = info["schet_pp"]
            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"
            if is_cell_empty(idx_deliv) and info["deliv"]:
                df_block.iat[idx, idx_deliv] = info["deliv"]

            stats["matched_head"] += 1
            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_head" if final_d.upper() == "САМОВЫВОЗ" else "head"
            row_meta.append({"idx": idx, "type": "data", "zone": current_zone, "status": st, "deliv": final_d})

        else:
            has_qty = not is_cell_empty(idx_qty)
            has_deliv = not is_cell_empty(idx_deliv)
            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            if has_qty and has_deliv:
                stats["already_filled"] += 1
                if final_d.upper() == "САМОВЫВОЗ":
                    st = "samovyvoz_head" if current_zone == "head" else "samovyvoz_filial"
                else:
                    st = "already_filled"
            else:
                stats["not_found"] += 1
                st = "not_found"
            row_meta.append({"idx": idx, "type": "data", "zone": current_zone, "status": st, "deliv": final_d})

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
