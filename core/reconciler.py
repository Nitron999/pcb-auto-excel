import pandas as pd
import numpy as np

def clean_ps(series):
    return (
        series.fillna('')
        .astype(str)
        .str.strip()
        .str.replace(r'\.0$', '', regex=True)
        .replace({'nan': '', 'None': '', 'NaN': '', '<NA>': ''})
    )

def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение исходных файлов Excel...")

    df_posylka = pd.read_excel(p_pos.strip(' "'))
    df_pp = pd.read_excel(p_pp.strip(' "'), sheet_name=0, header=None)
    df_filial = pd.read_excel(p_fil.strip(' "'), sheet_name=0, header=None)

    if progress_callback:
        progress_callback(0.3, "Анализ структуры и поиск дубликатов в сводках...")

    # Подготовка ключа PS в Посылке
    df_posylka['_PS_clean'] = clean_ps(df_posylka['PS'])

    # Подготовка Сводки ПП (Голова): 2 - PS, 4 - По счету, 14 - Доставка
    df_pp['_PS_clean'] = clean_ps(df_pp[2])
    df_pp_valid = df_pp[df_pp['_PS_clean'] != ''].copy()
    pp_duplicates = [str(x) for x in df_pp_valid[df_pp_valid.duplicated(subset=['_PS_clean'], keep=False)]['_PS_clean'].unique().tolist()]
    df_pp_unique = df_pp_valid.drop_duplicates(subset=['_PS_clean'], keep='last')
    pp_dict = df_pp_unique.set_index('_PS_clean').to_dict('index')

    # Подготовка Сводки Филиала: 4 - PS, 7 - По счету, 1 - Доставка
    df_filial['_PS_clean'] = clean_ps(df_filial[4])
    df_filial_valid = df_filial[df_filial['_PS_clean'] != ''].copy()
    filial_duplicates = [str(x) for x in df_filial_valid[df_filial_valid.duplicated(subset=['_PS_clean'], keep=False)]['_PS_clean'].unique().tolist()]
    df_filial_unique = df_filial_valid.drop_duplicates(subset=['_PS_clean'], keep='last')
    filial_dict = df_filial_unique.set_index('_PS_clean').to_dict('index')

    if progress_callback:
        progress_callback(0.5, "Сопоставление заказов и заполнение пустых ячеек...")

    stats = {
        "total_rows": len(df_posylka),
        "empty_ps_rows": 0,
        "already_filled": 0,
        "matched_head": 0,
        "matched_filial": 0,
        "not_found": 0,
        "duplicates_in_summaries": len(pp_duplicates) + len(filial_duplicates),
        "conflicts_both_sources": 0,
        "warnings": [],
        "row_statuses": {}
    }

    if pp_duplicates:
        stats["warnings"].append(f"В Сводке ПП найдены дубликаты PS ({len(pp_duplicates)} шт.): {', '.join(pp_duplicates[:5])}")
    if filial_duplicates:
        stats["warnings"].append(f"В Сводке Филиала найдены дубликаты PS ({len(filial_duplicates)} шт.): {', '.join(filial_duplicates[:5])}")

    pos_dups = df_posylka[df_posylka['_PS_clean'] != '']
    pos_dup_list = [str(x) for x in pos_dups[pos_dups.duplicated(subset=['_PS_clean'], keep=False)]['_PS_clean'].unique().tolist()]
    if pos_dup_list:
        stats["warnings"].append(f"В самой Посылке повторяются номера PS ({len(pos_dup_list)} шт.): {', '.join(pos_dup_list[:5])}")

    total = len(df_posylka)
    for idx, row in df_posylka.iterrows():
        if progress_callback and isinstance(idx, int) and idx % 50 == 0:
            progress_callback(0.5 + 0.4 * (idx / max(total, 1)), f"Обработано строк: {idx} из {total}")

        ps = str(row['_PS_clean']).strip()
        if not ps:
            stats["empty_ps_rows"] += 1
            continue

        needs_qty = pd.isna(row['По счету']) or str(row['По счету']).strip() == ''
        needs_deliv = pd.isna(row['Доставка']) or str(row['Доставка']).strip() == ''

        if not needs_qty and not needs_deliv:
            stats["already_filled"] += 1
            continue

        in_head = ps in pp_dict
        in_filial = ps in filial_dict

        if in_head and in_filial:
            stats["conflicts_both_sources"] += 1
            stats["warnings"].append(f"Расхождение: заказ PS {ps} найден одновременно и в Голове, и в Филиале (приоритет отдан Голове).")

        if in_head:
            if needs_qty:
                val_qty = pp_dict[ps].get(4, np.nan)
                df_posylka.at[idx, 'По счету'] = val_qty
            if needs_deliv:
                val_deliv = pp_dict[ps].get(14, np.nan)
                df_posylka.at[idx, 'Доставка'] = val_deliv
            stats["matched_head"] += 1
            stats["row_statuses"][idx] = "conflict" if in_filial else "head"

        elif in_filial:
            if needs_qty:
                val_qty = filial_dict[ps].get(7, np.nan)
                df_posylka.at[idx, 'По счету'] = val_qty
            if needs_deliv:
                val_deliv = filial_dict[ps].get(1, np.nan)
                df_posylka.at[idx, 'Доставка'] = val_deliv
            stats["matched_filial"] += 1
            stats["row_statuses"][idx] = "filial"

        else:
            stats["not_found"] += 1
            stats["row_statuses"][idx] = "not_found"

    df_posylka.drop(columns=['_PS_clean'], inplace=True)

    if progress_callback:
        progress_callback(1.0, "Сведение завершено!")

    return df_posylka, stats
