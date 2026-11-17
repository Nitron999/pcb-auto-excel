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
    s = str(val).strip()
    if s.lower() in ('nan', 'none', 'nat', '<na>', ''):
        return ''
    if s.endswith('.0'):
        s = s[:-2]
    return s

def extract_schet_pp(row_dict):
    # Ищем 4-значный номер счета ПП (например, 0682, 0654, 0611) в приоритетных колонках Сводки ПП
    for col_idx in [3, 0, 1, 5, 6]:
        if col_idx in row_dict and pd.notna(row_dict[col_idx]):
            raw = str(row_dict[col_idx]).strip()
            if raw.endswith('.0'):
                raw = raw[:-2]
            if raw.isdigit() and len(raw) <= 4:
                return raw.zfill(4)
            match = re.search(r'\b(0\d{3}|\d{4})\b', raw)
            if match:
                return match.group(1).zfill(4)
    return ""

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
        progress_callback(0.1, "Чтение файлов и поиск последнего блока Посылки...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    df_pos_raw = pd.read_excel(real_pos, sheet_name=0, header=None)
    df_pp = pd.read_excel(real_pp, sheet_name=0, header=None)
    df_filial = pd.read_excel(real_fil, sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Подготовка сводок Головы (ПП) и Филиала...")

    # Сводка ПП (Голова): 2 - PS, 4 - По счету, 14 - Доставка
    df_pp['_PS_clean'] = df_pp[2].apply(clean_ps)
    df_pp_valid = df_pp[df_pp['_PS_clean'] != ''].copy()
    pp_duplicates = [str(x) for x in df_pp_valid[df_pp_valid.duplicated(subset=['_PS_clean'], keep=False)]['_PS_clean'].unique().tolist()]
    df_pp_unique = df_pp_valid.drop_duplicates(subset=['_PS_clean'], keep='last')
    pp_dict = df_pp_unique.set_index('_PS_clean').to_dict('index')

    # Сводка Филиала: 4 - PS, 7 - По счету, 1 - Доставка
    df_filial['_PS_clean'] = df_filial[4].apply(clean_ps)
    df_filial_valid = df_filial[df_filial['_PS_clean'] != ''].copy()
    filial_duplicates = [str(x) for x in df_filial_valid[df_filial_valid.duplicated(subset=['_PS_clean'], keep=False)]['_PS_clean'].unique().tolist()]
    df_filial_unique = df_filial_valid.drop_duplicates(subset=['_PS_clean'], keep='last')
    filial_dict = df_filial_unique.set_index('_PS_clean').to_dict('index')

    header_row = df_pos_raw.iloc[0].tolist()
    col_map = {}
    for c_idx, val in enumerate(header_row):
        name = str(val).strip()
        col_map[name] = c_idx

    idx_ps = col_map.get('PS', 1)
    idx_schet_pp = col_map.get('Счет пп', 2)
    idx_qty = col_map.get('По счету', 4)
    idx_otk = col_map.get('Этикетка', 7)
    idx_deliv = col_map.get('Доставка', 10)

    start_idx = find_last_block_start(df_pos_raw)
    df_block = df_pos_raw.iloc[start_idx:].copy().reset_index(drop=True)

    # Приводим текстовые колонки к object, чтобы без проблем записывать "0682" и "ОТК"
    for col_i in [idx_schet_pp, idx_otk, idx_deliv]:
        if col_i < df_block.shape[1]:
            df_block[col_i] = df_block[col_i].astype(object)

    if progress_callback:
        progress_callback(0.5, f"Обработка последнего блока ({len(df_block)} строк)...")

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
        "row_zones": {},  # 'head' (верхний блок ПП) или 'filial' (нижний блок после плашки филиал)
        "header_row": header_row
    }

    if pp_duplicates:
        stats["warnings"].append(f"В Сводке ПП найдены дубликаты PS ({len(pp_duplicates)} шт.): {', '.join(pp_duplicates[:5])}")
    if filial_duplicates:
        stats["warnings"].append(f"В Сводке Филиала найдены дубликаты PS ({len(filial_duplicates)} шт.): {', '.join(filial_duplicates[:5])}")

    current_zone = "head"  # До красной строки "филиал" идёт первый блок (Печатные платы)

    for idx in range(len(df_block)):
        row = df_block.iloc[idx]
        first_cell = str(row[0]).strip().upper() if pd.notna(row[0]) else ""
        row_joined = " ".join([str(x).upper() for x in row.tolist() if pd.notna(x)])

        if "ОТПРАВКА" in row_joined or "ПРИХОД" in row_joined:
            stats["row_statuses"][idx] = "banner_date"
            stats["row_zones"][idx] = current_zone
            continue
        if first_cell == "ФИЛИАЛ":
            current_zone = "filial"  # Переключаемся на второй блок (Филиал Электроконнект)
            stats["row_statuses"][idx] = "banner_filial"
            stats["row_zones"][idx] = current_zone
            continue
        if first_cell == "ФИРМА":
            stats["row_statuses"][idx] = "header"
            stats["row_zones"][idx] = current_zone
            continue

        stats["row_zones"][idx] = current_zone

        ps = clean_ps(row[idx_ps])
        if not ps or not ps.isdigit():
            stats["empty_ps_rows"] += 1
            continue

        val_qty_curr = row[idx_qty]
        val_deliv_curr = row[idx_deliv]
        val_schet_curr = row[idx_schet_pp] if idx_schet_pp < len(row) else np.nan
        val_otk_curr = row[idx_otk] if idx_otk < len(row) else np.nan

        needs_qty = pd.isna(val_qty_curr) or str(val_qty_curr).strip() == ''
        needs_deliv = pd.isna(val_deliv_curr) or str(val_deliv_curr).strip() == ''
        needs_schet = pd.isna(val_schet_curr) or str(val_schet_curr).strip() == ''
        needs_otk = pd.isna(val_otk_curr) or str(val_otk_curr).strip() == ''

        in_head = ps in pp_dict
        in_filial = ps in filial_dict

        if in_head and in_filial:
            stats["conflicts_both_sources"] += 1

        # Если строка в первом блоке (Голова / ПП)
        if current_zone == "head" and in_head:
            if needs_qty:
                df_block.iat[idx, idx_qty] = pp_dict[ps].get(4, np.nan)
            if needs_deliv:
                df_block.iat[idx, idx_deliv] = pp_dict[ps].get(14, np.nan)
            if needs_schet and idx_schet_pp < df_block.shape[1]:
                schet_val = extract_schet_pp(pp_dict[ps])
                if schet_val:
                    df_block.iat[idx, idx_schet_pp] = schet_val
            if needs_otk and idx_otk < df_block.shape[1]:
                df_block.iat[idx, idx_otk] = "ОТК"

            stats["matched_head"] += 1
            final_deliv = str(df_block.iat[idx, idx_deliv]).upper()
            stats["row_statuses"][idx] = "samovyvoz_head" if "САМОВЫВОЗ" in final_deliv else "head"

        # Если строка во втором блоке (Филиал) или найдена в Филиале
        elif in_filial:
            if needs_qty:
                df_block.iat[idx, idx_qty] = filial_dict[ps].get(7, np.nan)
            if needs_deliv:
                df_block.iat[idx, idx_deliv] = filial_dict[ps].get(1, np.nan)
            stats["matched_filial"] += 1
            final_deliv = str(df_block.iat[idx, idx_deliv]).upper()
            stats["row_statuses"][idx] = "samovyvoz_filial" if "САМОВЫВОЗ" in final_deliv else "filial"

        # Запасной вариант, если заказ из Головы попал ниже плашки филиала
        elif in_head:
            if needs_qty:
                df_block.iat[idx, idx_qty] = pp_dict[ps].get(4, np.nan)
            if needs_deliv:
                df_block.iat[idx, idx_deliv] = pp_dict[ps].get(14, np.nan)
            if needs_schet and idx_schet_pp < df_block.shape[1]:
                schet_val = extract_schet_pp(pp_dict[ps])
                if schet_val:
                    df_block.iat[idx, idx_schet_pp] = schet_val
            if needs_otk and idx_otk < df_block.shape[1]:
                df_block.iat[idx, idx_otk] = "ОТК"
            stats["matched_head"] += 1
            final_deliv = str(df_block.iat[idx, idx_deliv]).upper()
            stats["row_statuses"][idx] = "samovyvoz_head" if "САМОВЫВОЗ" in final_deliv else "head"

        else:
            if not needs_qty and not needs_deliv:
                stats["already_filled"] += 1
                final_deliv = str(val_deliv_curr).upper()
                if "САМОВЫВОЗ" in final_deliv:
                    stats["row_statuses"][idx] = "samovyvoz_head" if current_zone == "head" else "samovyvoz_filial"
            else:
                stats["not_found"] += 1
                stats["row_statuses"][idx] = "not_found"

    df_result = df_block.copy()
    clean_cols = [str(c) if pd.notna(c) else f"Колонка_{i+1}" for i, c in enumerate(header_row)]
    df_result.columns = clean_cols[:len(df_result.columns)]

    if progress_callback:
        progress_callback(1.0, "Последний блок успешно сведён!")

    return df_result, stats
