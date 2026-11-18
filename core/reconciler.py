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

def clean_qty(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return np.nan
    if isinstance(val, (int, float, np.number)) and not pd.isna(val):
        return int(val) if float(val).is_integer() else val
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return np.nan
    if s.endswith('.0'):
        s = s[:-2]
    if s.isdigit():
        return int(s)
    m = re.match(r'^(\d+)', s)
    if m:
        return int(m.group(1))
    return np.nan

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

def parse_firma_zakazchik_cell(raw_text):
    if raw_text is None or (isinstance(raw_text, float) and np.isnan(raw_text)):
        return "", ""
    s = str(raw_text).strip()
    if not s or s.lower() in ('nan', 'none', 'nat'):
        return "", ""

    deliv = extract_delivery_from_text(s)

    clean_name = re.sub(
        r'[\(\[\s,;\-\/\:]*(самовывоз|самов\.?|сдэк|сдек|cdek|курьер|дм|деловые\s*линии)[\)\]\s,;\-\/\.\:]*$',
        '', s, flags=re.IGNORECASE
    ).strip()
    clean_name = re.sub(r'[\s,;\-\/\:]+К$', '', clean_name).strip()

    if clean_name.upper() in ("ДМ", "СДЭК", "СДЕК", "CDEK", "К", "САМОВЫВОЗ", "КУРЬЕР", ""):
        return "", deliv
    
    if len(clean_name) < 2:
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

def normalize_company_key(name):
    s = str(name or "").strip().lower()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\b(дм|сдэк|сдек|cdek|самовывоз|курьер|к|ооо|ао|зао|пао|ип|фгуп|нпц|пз|мпотк|сертификат|сертификаты)\b', '', s)
    s = re.sub(r'[^а-яa-z0-9]', '', s)
    return s

def detect_sheet_columns(df, is_head=True):
    num_rows, num_cols = df.shape
    if num_rows < 2 or num_cols < 4:
        return None

    header_row_idx = -1
    best_score = 0

    for r_idx in range(min(20, num_rows)):
        row_vals = df.iloc[r_idx].tolist()
        row_strs = [str(x).strip().lower() for x in row_vals if pd.notna(x)]
        score = 0
        for cell in row_strs:
            if cell == 'ps' or re.search(r'\bps\b', cell):
                score += 5
            if 'фирма' in cell or 'заказчик' in cell:
                score += 6
            if any(k in cell for k in ['счет', 'счёт', 'кол', 'доставк', 'сч/н', 'шт']):
                score += 2
        if score > best_score:
            best_score = score
            header_row_idx = r_idx

    cols = {
        "header_row": header_row_idx if best_score >= 4 else 0,
        "ps": None,
        "firma_zakazchik": None,
        "schet_pp": None,
        "schet_n": None,
        "qty": None,
        "deliv": None
    }

    raw_headers = df.iloc[cols["header_row"]].tolist()
    headers = [str(x).strip().lower() if pd.notna(x) else "" for x in raw_headers]

    for idx, h in enumerate(headers):
        if 'фирма заказчик' in h or 'фирма-заказчик' in h:
            cols["firma_zakazchik"] = idx
            break
    if cols["firma_zakazchik"] is None:
        for idx, h in enumerate(headers):
            if any(bad in h for bad in ['плата', 'чертеж', 'децим', 'файл', 'наименование плат', 'тел', 'контакт']):
                continue
            if 'заказчик' in h or 'фирма' in h or 'клиент' in h or 'покупатель' in h:
                cols["firma_zakazchik"] = idx
                break

    for idx, h in enumerate(headers):
        if not h:
            continue
        if cols["ps"] is None and (h == "ps" or re.search(r'\bps\b', h)):
            cols["ps"] = idx
        elif cols["schet_n"] is None and any(k in h for k in ['сч/н', 'сч\\н', 'счет н', 'счёт н']):
            cols["schet_n"] = idx
        elif cols["schet_pp"] is None and is_head and any(k in h for k in ['счет пп', 'счёт пп', '№ счета', 'счет', 'счёт']):
            if 'по счету' not in h and 'по счёту' not in h and 'сч/н' not in h:
                cols["schet_pp"] = idx
        elif cols["qty"] is None and any(k in h for k in ['по счету', 'по счёту', 'кол-во', 'количество', 'кол.', 'шт']):
            cols["qty"] = idx
        elif cols["deliv"] is None and any(k in h for k in ['доставк', 'способ', 'отправк', 'транспорт']):
            cols["deliv"] = idx

    if is_head:
        if cols["ps"] is None and num_cols > 2: cols["ps"] = 2
        if cols["qty"] is None and num_cols > 4: cols["qty"] = 4
        if cols["deliv"] is None and num_cols > 14: cols["deliv"] = 14
        if cols["firma_zakazchik"] is None and num_cols > 1: cols["firma_zakazchik"] = 1
        if cols["schet_pp"] is None and num_cols > 0: cols["schet_pp"] = 0
    else:
        if cols["ps"] is None and num_cols > 4: cols["ps"] = 4
        if cols["qty"] is None and num_cols > 7: cols["qty"] = 7
        if cols["firma_zakazchik"] is None and num_cols > 1: cols["firma_zakazchik"] = 1
        if cols["deliv"] is None and num_cols > 1: cols["deliv"] = 1

    return cols

def build_workbook_lookup(excel_path, is_head=True):
    xl = pd.ExcelFile(excel_path)
    ps_lookup = {}
    sn_lookup = {}
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

        # Прямое чтение строки, без наследования сверху (без forward fill)
        for i in range(len(sub)):
            row_vals = sub.iloc[i].tolist()

            ps = clean_ps(row_vals[ps_col]) if (ps_col is not None and ps_col < len(row_vals)) else ""
            curr_sn = clean_schet_n(row_vals[schet_n_col]) if (schet_n_col is not None and schet_n_col < len(row_vals)) else ""

            if not ps and not curr_sn:
                continue

            # Фирма и Доставка
            fz_firma, fz_deliv = "", ""
            if fz_col is not None and fz_col < len(row_vals):
                fz_firma, fz_deliv = parse_firma_zakazchik_cell(row_vals[fz_col])

            row_deliv = fz_deliv
            if not row_deliv and deliv_col is not None and deliv_col < len(row_vals):
                row_deliv = extract_delivery_from_text(row_vals[deliv_col])

            if fz_firma and row_deliv:
                f_key = normalize_company_key(fz_firma)
                if f_key:
                    firma_deliv_history[f_key] = row_deliv

            # Счет ПП (4 цифры)
            curr_schet_pp = ""
            if is_head and schet_pp_col is not None and schet_pp_col < len(row_vals):
                raw_s = str(row_vals[schet_pp_col]).strip()
                if raw_s.endswith('.0'):
                    raw_s = raw_s[:-2]
                digits = re.sub(r'\D', '', raw_s)
                if digits and len(digits) <= 4:
                    curr_schet_pp = digits.zfill(4)

            # Количество
            curr_qty = np.nan
            if qty_col is not None and qty_col < len(row_vals):
                curr_qty = clean_qty(row_vals[qty_col])

            if curr_sn:
                sn_entry = sn_lookup.get(curr_sn, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
                if fz_firma: sn_entry["firma"] = fz_firma
                if curr_schet_pp: sn_entry["schet_pp"] = curr_schet_pp
                if pd.notna(curr_qty): sn_entry["qty"] = curr_qty
                if row_deliv: sn_entry["deliv"] = row_deliv
                sn_lookup[curr_sn] = sn_entry

            if ps:
                if ps in seen_ps and ps not in duplicates:
                    duplicates.append(ps)
                seen_ps.add(ps)
                
                entry = ps_lookup.get(ps, {"firma": "", "schet_pp": "", "schet_n": curr_sn, "qty": np.nan, "deliv": ""})
                if fz_firma: entry["firma"] = fz_firma
                if curr_schet_pp: entry["schet_pp"] = curr_schet_pp
                if curr_sn: entry["schet_n"] = curr_sn
                if pd.notna(curr_qty): entry["qty"] = curr_qty
                if row_deliv: entry["deliv"] = row_deliv
                ps_lookup[ps] = entry

    # Восстановление пропусков внутри одной сводки
    for ps, entry in ps_lookup.items():
        sn = entry.get("schet_n", "")
        if sn and sn in sn_lookup:
            if not entry["firma"] and sn_lookup[sn]["firma"]:
                entry["firma"] = sn_lookup[sn]["firma"]
            if not entry["schet_pp"] and sn_lookup[sn]["schet_pp"]:
                entry["schet_pp"] = sn_lookup[sn]["schet_pp"]
            if not entry["deliv"] and sn_lookup[sn]["deliv"]:
                entry["deliv"] = sn_lookup[sn]["deliv"]

        if not entry["deliv"] and entry.get("firma"):
            f_key = normalize_company_key(entry["firma"])
            if f_key in firma_deliv_history:
                entry["deliv"] = firma_deliv_history[f_key]

    return ps_lookup, sn_lookup, firma_deliv_history, duplicates

def find_last_block_start(df_raw, idx_ps=1):
    last_data_row = -1
    for idx in range(len(df_raw) - 1, -1, -1):
        row_vals = df_raw.iloc[idx].tolist()
        ps_val = clean_ps(row_vals[idx_ps]) if idx_ps < len(row_vals) else ""
        if ps_val:
            last_data_row = idx
            break
    if last_data_row == -1:
        return 0
    for idx in range(last_data_row, -1, -1):
        row_vals = [str(x).strip().upper() for x in df_raw.iloc[idx].tolist() if pd.notna(x)]
        row_str = " ".join(row_vals)
        if "ОТПРАВКА" in row_str or "ПРИХОД" in row_str:
            return idx
    return 0

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение файлов и поиск нижней незаполненной посылки...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Анализ Сводки ООО Печатные платы 2026 и Сводки Филиал 2025...")

    pp_dict, pp_sn_lookup, pp_firma_hist, pp_duplicates = build_workbook_lookup(real_pp, is_head=True)
    filial_dict, filial_sn_lookup, filial_firma_hist, filial_duplicates = build_workbook_lookup(real_fil, is_head=False)

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

    start_idx = find_last_block_start(df_pos_raw, idx_ps=idx_ps)
    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True)
    df_block = df_block.astype(object)
    num_block_cols = df_block.shape[1]

    last_valid_idx = len(df_block) - 1
    while last_valid_idx > 0:
        r_vals = df_block.iloc[last_valid_idx].tolist()
        if any(pd.notna(x) and str(x).strip() != '' for x in r_vals):
            break
        last_valid_idx -= 1
    df_block = df_block.iloc[:last_valid_idx + 1].copy().reset_index(drop=True)

    if progress_callback:
        progress_callback(0.6, "Заполнение первого блока (Печатные платы) и второго блока (Филиал)...")

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

        def is_cell_empty(col_i):
            if col_i >= num_block_cols:
                return False
            v = df_block.iat[idx, col_i]
            return pd.isna(v) or str(v).strip() == ''

        # БЛОК 1: ПЕЧАТНЫЕ ПЛАТЫ
        if current_zone == "head":
            info = pp_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            sn_fallback = pp_sn_lookup.get(schet_n_val, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})

            firma_best = info["firma"] or sn_fallback["firma"]
            if is_cell_empty(idx_firma) and firma_best:
                df_block.iat[idx, idx_firma] = firma_best

            schet_pp_best = info["schet_pp"] or sn_fallback["schet_pp"]
            if is_cell_empty(idx_schet_pp) and schet_pp_best:
                df_block.iat[idx, idx_schet_pp] = schet_pp_best

            qty_best = info["qty"] if pd.notna(info["qty"]) else sn_fallback["qty"]
            if is_cell_empty(idx_qty) and pd.notna(qty_best):
                df_block.iat[idx, idx_qty] = qty_best

            if is_cell_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"

            if is_cell_empty(idx_deliv):
                firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
                f_key = normalize_company_key(firma_now)
                d_val = info["deliv"] or sn_fallback["deliv"] or pp_firma_hist.get(f_key, "")
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if ps in pp_dict or schet_n_val in pp_sn_lookup:
                stats["matched_head"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_head" if final_d.upper() == "САМОВЫВОЗ" else "head"
            row_meta.append({"idx": idx, "type": "data", "zone": "head", "status": st, "deliv": final_d})

        # БЛОК 2: ФИЛИАЛ
        else:
            info = filial_dict.get(ps, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})
            sn_fallback = filial_sn_lookup.get(schet_n_val, {"firma": "", "schet_pp": "", "qty": np.nan, "deliv": ""})

            f_best = info["firma"] or sn_fallback["firma"]
            if is_cell_empty(idx_firma) and f_best:
                df_block.iat[idx, idx_firma] = f_best

            qty_best = info["qty"] if pd.notna(info["qty"]) else sn_fallback["qty"]
            if is_cell_empty(idx_qty) and pd.notna(qty_best):
                df_block.iat[idx, idx_qty] = qty_best

            if is_cell_empty(idx_deliv):
                firma_now = str(df_block.iat[idx, idx_firma] or "").strip()
                f_key = normalize_company_key(firma_now)
                d_val = info["deliv"] or sn_fallback["deliv"] or filial_firma_hist.get(f_key, "")
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if ps in filial_dict or schet_n_val in filial_sn_lookup:
                stats["matched_filial"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_filial" if final_d.upper() == "САМОВЫВОЗ" else "filial"
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
