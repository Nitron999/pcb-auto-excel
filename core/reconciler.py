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
    return s if re.match(r'^\d{5,7}$', s) else ''

def clean_schet_n(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    return s if re.match(r'^\d{4,6}$', s) else ''

def clean_schet_pp(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    s = str(val).strip()
    if s.endswith('.0'):
        s = s[:-2]
    # Счет пп — это 3-4 цифры с ведущим нулем (например 0328, 0682)
    if re.match(r'^0\d{2,3}$', s):
        return s.zfill(4)
    return ''

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
    return np.nan

def normalize_delivery(raw_val):
    if raw_val is None or (isinstance(raw_val, float) and np.isnan(raw_val)):
        return ""
    s = str(raw_val).strip().upper()
    if not s or s in ('NAN', 'NONE', 'NAT'):
        return ""
    # Важно: "курьер СДЭК" относится к СДЭК, поэтому СДЭК проверяем раньше Курьера!
    if "САМОВЫВОЗ" in s or "САМОВ" in s:
        return "Самовывоз"
    if "СДЭК" in s or "CDEK" in s or "СДЕК" in s:
        return "СДЭК"
    if "ДОСТАВКА МОСКВА" in s or "ДЕЛОВ" in s or re.search(r'(^|[\s,;/\(\-\.\:])ДМ($|[\s,;/\)\-\.\:])', s):
        return "ДМ"
    if s == "К" or "КУРЬЕР" in s or re.search(r'(^|[\s,;/\(\-\.\:])К($|[\s,;/\)\-\.\:])', s):
        return "К"
    return ""

def is_non_firma_text(text):
    """Отсекает служебные подписи в колонке 'Фирма, заказчик', оставляя только названия компаний/ФИО."""
    s = str(text or "").strip()
    if not s or len(s) < 2:
        return True
    u = s.upper()
    if normalize_delivery(s) != "":
        return True
    bad_prefixes = (
        'ЗАЯВКА', '(ЗАЯВКА', 'ИГК', 'НУМЕРАЦИЯ', 'ОДИН СЧЕТ', 'СПЕЦСЧЕТ',
        'ДОСТАВКА', 'СЕРТИФИКАТ', 'В РАБОТЕ', 'ПЕРЕДЕЛКА', 'ОПЛАТА',
        'Г.', 'Г ', 'МОСКВА', 'УЛ.', 'УЛ ', 'ПР-Т', 'ПРОСП', 'ДОМ ', 'Д.',
        '+7', '8-', '8(', '8 ', '79', '89', 'ТЕЛ', 'ОТК', 'ФИРМА', 'ЗАКАЗЧИК'
    )
    if any(u.startswith(p) for p in bad_prefixes):
        return True
    if re.match(r'^[\d\s\-\(\)\+\.,/]+$', s):
        return True
    if re.search(r'\d{3}[\s\-]?\d{2}[\s\-]?\d{2}', s):
        return True
    return False

def clean_firma_name(raw_text):
    if raw_text is None or (isinstance(raw_text, float) and np.isnan(raw_text)):
        return ""
    s = str(raw_text).strip()
    if is_non_firma_text(s):
        return ""
    # Убираем приписки в скобках вида "(заявка 790)" или "(3604)"
    s = re.sub(r'\s*\((?:заявка\s*)?\d+\)\s*$', '', s, flags=re.IGNORECASE).strip()
    return s if not is_non_firma_text(s) else ""

def normalize_company_key(name):
    s = str(name or "").strip().lower()
    s = re.sub(r'\(.*?\)', '', s)
    s = re.sub(r'\b(дм|сдэк|сдек|cdek|самовывоз|курьер|к|ооо|ао|зао|пао|ип|фгуп|нпц|пз|мпотк|по|фгбоу|во)\b', '', s)
    s = re.sub(r'[^а-яa-z0-9]', '', s)
    return s

def delivery_sort_rank(deliv_val):
    d = str(deliv_val).strip().upper()
    if d == "ДМ": return 1
    if d == "СДЭК": return 2
    if d == "К": return 3
    if d == "САМОВЫВОЗ": return 4
    return 5

def parse_pp_workbook(excel_path):
    """
    Разбирает 'сводка ООО Печатные платы 2026.xlsx' (лист СВОДКА):
    - col 2 (C): PS
    - col 4 (E): количество (По счету)
    - col 12 (M): Счет пп (4 цифры, например 0328, 0467)
    - col 13 (N): Сч/Н (5 цифр, например 10418, 10181)
    - col 14 (O): Фирма, заказчик + строкой ниже Доставка ('Доставка Москва' -> ДМ, 'КУРЬЕР' -> К, 'Самовывоз' -> Самовывоз)
    """
    xl = pd.ExcelFile(excel_path)
    sheet_name = 'СВОДКА' if 'СВОДКА' in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet_name, header=None)

    ps_qty = {}          # PS -> последнее актуальное количество
    sn_info = {}         # Сч/Н -> {"firma": ..., "schet_pp": ..., "deliv": ...}
    ps_info = {}         # PS -> {"firma": ..., "schet_pp": ..., "schet_n": ..., "deliv": ...}
    firma_deliv = {}     # norm_firma -> deliv
    seen_ps = set()
    dups = []

    curr_sn = ""
    curr_firma = ""
    curr_spp = ""
    curr_deliv = ""

    for i in range(3, len(df)):
        row = df.iloc[i].tolist()
        if len(row) < 15:
            continue

        ps = clean_ps(row[2])
        qty = clean_qty(row[4])
        spp = clean_schet_pp(row[12])
        sn = clean_schet_n(row[13])
        col14 = str(row[14]).strip() if pd.notna(row[14]) else ""

        # Если встретился новый номер счета Сч/Н в колонке 13 — обновляем контекст заказа
        if sn:
            curr_sn = sn
            curr_spp = spp if spp else ""
            f_cand = clean_firma_name(col14)
            curr_firma = f_cand if f_cand else ""
            curr_deliv = normalize_delivery(col14)
        else:
            if spp and not curr_spp:
                curr_spp = spp
            if not curr_firma:
                f_cand = clean_firma_name(col14)
                if f_cand:
                    curr_firma = f_cand
            d_cand = normalize_delivery(col14)
            if d_cand:
                curr_deliv = d_cand

        if curr_firma and curr_deliv:
            fk = normalize_company_key(curr_firma)
            if fk:
                firma_deliv[fk] = curr_deliv

        if curr_sn:
            entry_sn = sn_info.get(curr_sn, {"firma": "", "schet_pp": "", "deliv": ""})
            if curr_firma: entry_sn["firma"] = curr_firma
            if curr_spp: entry_sn["schet_pp"] = curr_spp
            if curr_deliv: entry_sn["deliv"] = curr_deliv
            sn_info[curr_sn] = entry_sn

        if ps:
            if ps in seen_ps and ps not in dups:
                dups.append(ps)
            seen_ps.add(ps)

            if pd.notna(qty):
                ps_qty[ps] = qty

            entry_ps = ps_info.get(ps, {"firma": "", "schet_pp": "", "schet_n": "", "deliv": ""})
            if curr_firma: entry_ps["firma"] = curr_firma
            if curr_spp: entry_ps["schet_pp"] = curr_spp
            if curr_sn: entry_ps["schet_n"] = curr_sn
            if curr_deliv: entry_ps["deliv"] = curr_deliv
            ps_info[ps] = entry_ps

    return ps_qty, ps_info, sn_info, firma_deliv, dups

def parse_filial_workbook(excel_path):
    """
    Разбирает 'сводка филиал 2025.xlsx' (лист ФИЛИАЛ):
    - col 0 (A): № счета (Сч/Н)
    - col 1 (B): Фирма, заказчик + строкой ниже Доставка ('Доставка Москва' -> ДМ, 'курьер СДЭК' -> СДЭК, 'Курьер' -> К, 'Самовывоз' -> Самовывоз)
    - col 4 (E): PS
    - col 7 (H): кол-во (По счету)
    """
    xl = pd.ExcelFile(excel_path)
    sheet_name = 'ФИЛИАЛ' if 'ФИЛИАЛ' in xl.sheet_names else xl.sheet_names[0]
    df = xl.parse(sheet_name, header=None)

    ps_qty = {}
    sn_info = {}
    ps_info = {}
    firma_deliv = {}
    seen_ps = set()
    dups = []

    curr_sn = ""
    curr_firma = ""
    curr_deliv = ""

    for i in range(12, len(df)):
        row = df.iloc[i].tolist()
        if len(row) < 8:
            continue

        sn = clean_schet_n(row[0])
        col1 = str(row[1]).strip() if pd.notna(row[1]) else ""
        ps = clean_ps(row[4])
        qty = clean_qty(row[7])

        # Новый счет в колонке 0 открывает новый блок заказа
        if sn:
            curr_sn = sn
            f_cand = clean_firma_name(col1)
            curr_firma = f_cand if f_cand else ""
            curr_deliv = normalize_delivery(col1)
        else:
            d_cand = normalize_delivery(col1)
            if d_cand:
                curr_deliv = d_cand

        if curr_firma and curr_deliv:
            fk = normalize_company_key(curr_firma)
            if fk:
                firma_deliv[fk] = curr_deliv

        if curr_sn:
            entry_sn = sn_info.get(curr_sn, {"firma": "", "deliv": ""})
            if curr_firma: entry_sn["firma"] = curr_firma
            if curr_deliv: entry_sn["deliv"] = curr_deliv
            sn_info[curr_sn] = entry_sn

        if ps:
            if ps in seen_ps and ps not in dups:
                dups.append(ps)
            seen_ps.add(ps)

            if pd.notna(qty):
                ps_qty[ps] = qty

            entry_ps = ps_info.get(ps, {"firma": "", "schet_n": "", "deliv": ""})
            if curr_firma: entry_ps["firma"] = curr_firma
            if curr_sn: entry_ps["schet_n"] = curr_sn
            if curr_deliv: entry_ps["deliv"] = curr_deliv
            ps_info[ps] = entry_ps

    return ps_qty, ps_info, sn_info, firma_deliv, dups

def build_posylka_history(df_raw, end_idx):
    """Читает предыдущие заполненные блоки Посылки для добора доставки и счета ПП по фирме/счету."""
    head_firma_deliv = {}
    head_sn_info = {}
    fil_firma_deliv = {}
    in_head = False

    for idx in range(1, end_idx):
        row = df_raw.iloc[idx].tolist()
        s0 = str(row[0]).strip() if pd.notna(row[0]) else ""
        s3 = str(row[3]).strip() if len(row) > 3 and pd.notna(row[3]) else ""

        if "ОТПРАВКА" in s3.upper():
            in_head = True
            continue
        if s0.upper() == "ФИЛИАЛ":
            in_head = False
            continue
        if not s0 or s0.upper() == "ФИРМА":
            continue

        d = normalize_delivery(row[10]) if len(row) > 10 else ""
        spp = clean_schet_pp(row[2]) if len(row) > 2 else ""
        sn = clean_schet_n(s3)
        fk = normalize_company_key(s0)

        if in_head:
            if fk and d:
                head_firma_deliv[fk] = d
            if sn:
                item = head_sn_info.get(sn, {"firma": "", "schet_pp": "", "deliv": ""})
                item["firma"] = s0
                if spp: item["schet_pp"] = spp
                if d: item["deliv"] = d
                head_sn_info[sn] = item
        else:
            if fk and d:
                fil_firma_deliv[fk] = d

    return head_firma_deliv, head_sn_info, fil_firma_deliv

def find_last_block_start(df_raw, idx_ps=1):
    last_data_row = -1
    for idx in range(len(df_raw) - 1, -1, -1):
        row_vals = df_raw.iloc[idx].tolist()
        if idx_ps < len(row_vals) and clean_ps(row_vals[idx_ps]):
            last_data_row = idx
            break
    if last_data_row == -1:
        return 0
    for idx in range(last_data_row, -1, -1):
        row_vals = [str(x).strip().upper() for x in df_raw.iloc[idx].tolist() if pd.notna(x)]
        if "ОТПРАВКА" in " ".join(row_vals):
            return idx
    return 0

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение Посылки и сводок...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Анализ Сводки ПП 2026 и Сводки Филиала 2025...")

    pp_qty, pp_ps_info, pp_sn_info, pp_firma_deliv, pp_dups = parse_pp_workbook(real_pp)
    fil_qty, fil_ps_info, fil_sn_info, fil_firma_deliv, fil_dups = parse_filial_workbook(real_fil)

    header_row = df_pos_raw.iloc[0].tolist()
    col_map = {str(val).strip(): c_idx for c_idx, val in enumerate(header_row)}

    idx_firma = col_map.get('Фирма', 0)
    idx_ps = col_map.get('PS', 1)
    idx_schet_pp = col_map.get('Счет пп', 2)
    idx_schet_n = col_map.get('Сч/Н', 3)
    idx_qty = col_map.get('По счету', 4)
    idx_otk = col_map.get('Этикетка', 7)
    idx_deliv = col_map.get('Доставка', 10)

    start_idx = find_last_block_start(df_pos_raw, idx_ps=idx_ps)
    pos_h_deliv, pos_h_sn, pos_f_deliv = build_posylka_history(df_pos_raw, start_idx)

    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True).astype(object)
    num_block_cols = df_block.shape[1]

    last_valid_idx = len(df_block) - 1
    while last_valid_idx > 0:
        if any(pd.notna(x) and str(x).strip() != '' for x in df_block.iloc[last_valid_idx].tolist()):
            break
        last_valid_idx -= 1
    df_block = df_block.iloc[:last_valid_idx + 1].copy().reset_index(drop=True)

    if progress_callback:
        progress_callback(0.6, "Заполнение Печатных плат и Филиала...")

    stats = {
        "total_rows": len(df_block),
        "empty_ps_rows": 0,
        "already_filled": 0,
        "matched_head": 0,
        "matched_filial": 0,
        "not_found": 0,
        "duplicates_in_summaries": len(pp_dups) + len(fil_dups),
        "conflicts_both_sources": 0,
        "warnings": [],
        "row_statuses": {},
        "header_row": header_row
    }

    if pp_dups:
        stats["warnings"].append(f"Повторяющиеся PS в Сводке ПП ({len(pp_dups)} шт.): {', '.join(pp_dups[:5])}")
    if fil_dups:
        stats["warnings"].append(f"Повторяющиеся PS в Сводке Филиала ({len(fil_dups)} шт.): {', '.join(fil_dups[:5])}")

    current_zone = "head"
    row_meta = []

    for idx in range(len(df_block)):
        row_vals = df_block.iloc[idx].tolist()
        first_cell = str(row_vals[0]).strip().upper() if len(row_vals) > 0 and pd.notna(row_vals[0]) else ""
        row_joined = " ".join([str(x).upper() for x in row_vals if pd.notna(x)])

        if "ОТПРАВКА" in row_joined or "ПРИХОД" in row_joined:
            row_meta.append({"idx": idx, "type": "banner_date", "status": "banner_date", "deliv": ""})
            continue
        if first_cell == "ФИЛИАЛ":
            current_zone = "filial"
            row_meta.append({"idx": idx, "type": "banner_filial", "status": "banner_filial", "deliv": ""})
            continue
        if first_cell == "ФИРМА":
            row_meta.append({"idx": idx, "type": "header", "status": "header", "deliv": ""})
            continue

        ps = clean_ps(row_vals[idx_ps]) if idx_ps < len(row_vals) else ""
        sn = clean_schet_n(row_vals[idx_schet_n]) if idx_schet_n < len(row_vals) else ""

        if not ps:
            stats["empty_ps_rows"] += 1
            row_meta.append({"idx": idx, "type": "empty", "status": "empty", "deliv": ""})
            continue

        def is_empty(c_i):
            if c_i >= num_block_cols:
                return False
            v = df_block.iat[idx, c_i]
            return pd.isna(v) or str(v).strip() == ''

        if not is_empty(idx_deliv):
            d_norm = normalize_delivery(df_block.iat[idx, idx_deliv])
            df_block.iat[idx, idx_deliv] = d_norm if d_norm else np.nan

        # =====================================================================
        # БЛОК 1: ПЕЧАТНЫЕ ПЛАТЫ (до красной строки "филиал")
        # =====================================================================
        if current_zone == "head":
            p_inf = pp_ps_info.get(ps, {"firma": "", "schet_pp": "", "schet_n": sn, "deliv": ""})
            s_inf = pp_sn_info.get(sn, {"firma": "", "schet_pp": "", "deliv": ""})
            h_sn = pos_h_sn.get(sn, {"firma": "", "schet_pp": "", "deliv": ""})

            # 1. Фирма
            f_best = p_inf["firma"] or s_inf["firma"] or h_sn["firma"]
            if is_empty(idx_firma) and f_best:
                df_block.iat[idx, idx_firma] = f_best

            # 2. Счет пп (4 цифры с ведущим нулем)
            spp_best = p_inf["schet_pp"] or s_inf["schet_pp"] or h_sn["schet_pp"]
            if is_empty(idx_schet_pp) and spp_best:
                df_block.iat[idx, idx_schet_pp] = spp_best

            # 3. По счету (количество из колонки E Сводки ПП)
            if is_empty(idx_qty) and ps in pp_qty:
                df_block.iat[idx, idx_qty] = pp_qty[ps]

            # 4. Этикетка -> ОТК
            if is_empty(idx_otk):
                df_block.iat[idx, idx_otk] = "ОТК"

            # 5. Доставка (ДМ, СДЭК, К, Самовывоз)
            if is_empty(idx_deliv):
                f_now = str(df_block.iat[idx, idx_firma] or "").strip()
                fk = normalize_company_key(f_now)
                d_val = (
                    p_inf["deliv"]
                    or s_inf["deliv"]
                    or pp_firma_deliv.get(fk, "")
                    or h_sn["deliv"]
                    or pos_h_deliv.get(fk, "")
                )
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if ps in pp_ps_info or sn in pp_sn_info:
                stats["matched_head"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_head" if final_d.upper() == "САМОВЫВОЗ" else "head"
            row_meta.append({"idx": idx, "type": "data", "status": st, "deliv": final_d})

        # =====================================================================
        # БЛОК 2: ФИЛИАЛ (после красной строки "филиал")
        # =====================================================================
        else:
            f_inf = fil_ps_info.get(ps, {"firma": "", "schet_n": sn, "deliv": ""})
            fs_inf = fil_sn_info.get(sn, {"firma": "", "deliv": ""})

            # 1. Фирма (если пусто)
            if is_empty(idx_firma):
                f_best = f_inf["firma"] or fs_inf["firma"]
                if f_best:
                    df_block.iat[idx, idx_firma] = f_best

            # 2. По счету (количество из колонки H Сводки Филиала)
            if is_empty(idx_qty) and ps in fil_qty:
                df_block.iat[idx, idx_qty] = fil_qty[ps]

            # 3. Доставка (ДМ, СДЭК, К, Самовывоз)
            if is_empty(idx_deliv):
                f_now = str(df_block.iat[idx, idx_firma] or "").strip()
                fk = normalize_company_key(f_now)
                d_val = (
                    f_inf["deliv"]
                    or fs_inf["deliv"]
                    or fil_firma_deliv.get(fk, "")
                    or pos_f_deliv.get(fk, "")
                )
                if d_val:
                    df_block.iat[idx, idx_deliv] = d_val

            if ps in fil_ps_info or sn in fil_sn_info:
                stats["matched_filial"] += 1
            else:
                stats["not_found"] += 1

            final_d = str(df_block.iat[idx, idx_deliv] or "").strip() if idx_deliv < num_block_cols else ""
            st = "samovyvoz_filial" if final_d.upper() == "САМОВЫВОЗ" else "filial"
            row_meta.append({"idx": idx, "type": "data", "status": st, "deliv": final_d})

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
        progress_callback(1.0, "Сведение завершено!")

    return df_result, stats
