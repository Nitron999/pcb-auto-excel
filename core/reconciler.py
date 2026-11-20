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


def get_delivery_rank(val):
    """Строгая иерархия сортировки: ДМ (1) -> СДЭК (2) -> К (3) -> Самовывоз (4)"""
    if not val:
        return 5
    s = str(val).strip().lower()
    if s in ['дм', 'доставка москва'] or 'москва' in s:
        return 1
    elif 'сдэк' in s or 'cdek' in s:
        return 2
    elif s in ['к', 'k'] or 'курьер' in s:
        return 3
    elif 'самовывоз' in s:
        return 4
    return 5


def clean_firm_name(val):
    """Нормализация имени фирмы для плотной группировки одинаковых заказчиков"""
    if not val or pd.isna(val):
        return ""
    s = str(val).strip().lower()
    s = re.sub(r'[\"«»\(\)]', '', s)
    for prefix in ['ооо', 'зао', 'пао', 'ао', 'пп', 'ип']:
        s = re.sub(rf'\b{prefix}\b', '', s)
    return ' '.join(s.split())


def clean_int_str(val):
    if val is None or (isinstance(val, float) and np.isnan(val)):
        return ''
    s = str(val).replace('.0', '').strip()
    if s.lower() in ('nan', 'none', 'nat', ''):
        return ''
    return s


def lookup_plate_info(df_plat, ps, schn):
    """Поиск по сводке ООО 'Печатные платы': Фирма, Счет пп (4 знака), По счету, Доставка"""
    ps_str = clean_int_str(ps)
    schn_str = clean_int_str(schn)

    col_ps = df_plat['№ заказа'] if '№ заказа' in df_plat.columns else df_plat.iloc[:, 2]
    col_qty = df_plat['количество'] if 'количество' in df_plat.columns else df_plat.iloc[:, 4]

    matches = df_plat[col_ps.astype(str).str.replace('.0', '', regex=False).str.strip() == ps_str]

    firm, sch_pp, po_schetu, delivery = None, None, None, 'К'

    for idx in matches.index:
        # Поднимаемся вверх по блоку заказа, чтобы забрать номер счёта, счёт пп и фирму
        for r in range(idx, max(-1, idx - 15), -1):
            row_schn = str(df_plat.iloc[r, 13]).replace('.0', '').strip()
            if row_schn == schn_str:
                firm = str(df_plat.iloc[r, 14]).strip() if pd.notna(df_plat.iloc[r, 14]) else firm
                sch_pp = str(df_plat.iloc[r, 12]).replace('.0', '').strip() if pd.notna(df_plat.iloc[r, 12]) else sch_pp
                break
        if firm:
            raw_q = col_qty.loc[idx]
            if pd.notna(raw_q):
                try:
                    po_schetu = int(float(raw_q))
                except ValueError:
                    po_schetu = None

            # Проверяем примечания/фирму в соседних строках на "Самовывоз" или "Доставка Москва"
            block_text = " ".join(
                df_plat.iloc[max(0, idx - 2):min(len(df_plat), idx + 3), 14].dropna().astype(str)
            ).lower()
            if 'самовывоз' in block_text or (firm and 'внииа' in firm.lower()):
                delivery = 'Самовывоз'
            elif 'москва' in block_text or (firm and 'авэкс' in firm.lower()):
                delivery = 'ДМ'
            break

    # Запасной поиск по номеру счёта Сч/Н, если по PS строка не зацепилась
    if not firm and schn_str:
        sn_matches = df_plat[df_plat.iloc[:, 13].astype(str).str.replace('.0', '', regex=False).str.strip() == schn_str]
        if not sn_matches.empty:
            r = sn_matches.index[-1]
            firm = str(df_plat.iloc[r, 14]).strip() if pd.notna(df_plat.iloc[r, 14]) else None
            sch_pp = str(df_plat.iloc[r, 12]).replace('.0', '', regex=False).strip() if pd.notna(df_plat.iloc[r, 12]) else None
            block_text = " ".join(
                df_plat.iloc[max(0, r - 2):min(len(df_plat), r + 3), 14].dropna().astype(str)
            ).lower()
            if 'самовывоз' in block_text or (firm and 'внииа' in firm.lower()):
                delivery = 'Самовывоз'
            elif 'москва' in block_text or (firm and 'авэкс' in firm.lower()):
                delivery = 'ДМ'

    if sch_pp:
        digits = re.sub(r'\D', '', str(sch_pp))
        if digits and len(digits) <= 4:
            sch_pp = digits.zfill(4)

    return firm, sch_pp, po_schetu, delivery


def lookup_filial_info(df_filial, ps, schn):
    """
    Поиск по сводке ФИЛИАЛ строго по номеру счёта (Col A).
    Доставка берётся ИСКЛЮЧИТЕЛЬНО из 2-й колонки (Col B) внутри блока этого счёта!
    """
    schn_str = clean_int_str(schn)
    ps_str = clean_int_str(ps)

    # Ищем вхождение счёта в 1-й колонке (индекс 0), берём самое свежее (снизу)
    m = df_filial[df_filial.iloc[:, 0].astype(str).str.replace('.0', '', regex=False).str.strip() == schn_str]
    if m.empty:
        return None, 'К'

    start_idx = m.index[-1]
    end_idx = start_idx + 1
    # Граница блока счёта — пока в 1-й колонке пусто (NaN)
    while end_idx < len(df_filial) and pd.isna(df_filial.iloc[end_idx, 0]):
        end_idx += 1

    po_schetu = None
    delivery = None

    for r in range(start_idx, end_idx):
        # 1. Ищем количество "По счету" из 8-й графы (индекс 7) по точному совпадению PS (индекс 4)
        row_ps = str(df_filial.iloc[r, 4]).replace('.0', '').strip()
        if row_ps == ps_str and pd.notna(df_filial.iloc[r, 7]):
            try:
                po_schetu = int(float(df_filial.iloc[r, 7]))
            except ValueError:
                pass

        # 2. Читаем способ доставки строго из 2-й колонки (индекс 1)
        col_b_text = str(df_filial.iloc[r, 1]).strip().lower() if pd.notna(df_filial.iloc[r, 1]) else ''
        if 'самовывоз' in col_b_text:
            delivery = 'Самовывоз'
        elif 'доставка москва' in col_b_text:
            delivery = 'ДМ'
        elif 'сдэк' in col_b_text:
            delivery = 'СДЭК'
        elif 'курьер' in col_b_text:
            delivery = 'К'

    return po_schetu, (delivery or 'К')


def run_reconciliation(p_pos, p_pp, p_fil, progress_callback=None):
    if progress_callback:
        progress_callback(0.1, "Чтение Посылки и сводных таблиц...")

    real_pos = resolve_shortcut(p_pos)
    real_pp = resolve_shortcut(p_pp)
    real_fil = resolve_shortcut(p_fil)

    xl_pos = pd.ExcelFile(real_pos)
    pos_sheet = 'Посылка' if 'Посылка' in xl_pos.sheet_names else xl_pos.sheet_names[0]
    df_pos = xl_pos.parse(pos_sheet, header=None)

    xl_plat = pd.ExcelFile(real_pp)
    plat_sheet = 'СВОДКА' if 'СВОДКА' in xl_plat.sheet_names else xl_plat.sheet_names[0]
    df_plat = xl_plat.parse(plat_sheet)

    xl_fil = pd.ExcelFile(real_fil)
    fil_sheet = 'ФИЛИАЛ' if 'ФИЛИАЛ' in xl_fil.sheet_names else xl_fil.sheet_names[0]
    df_filial = xl_fil.parse(fil_sheet)

    if progress_callback:
        progress_callback(0.35, "Поиск последней отгрузки и сбор строк...")

    # Находим заголовок ПОСЛЕДНЕЙ отгрузки в файле (не трогая историю)
    last_hdr_row = None
    for r in range(len(df_pos) - 1, -1, -1):
        c3 = str(df_pos.iloc[r, 3] if df_pos.shape[1] > 3 and pd.notna(df_pos.iloc[r, 3]) else '')
        c4 = str(df_pos.iloc[r, 4] if df_pos.shape[1] > 4 and pd.notna(df_pos.iloc[r, 4]) else '')
        if 'ОТПРАВКА' in c3.upper() or 'ОТПРАВКА' in c4.upper():
            last_hdr_row = r
            break

    if last_hdr_row is None:
        raise ValueError("Не найден заголовок отгрузки (строка с 'ОТПРАВКА')!")

    header_banner_row = df_pos.iloc[last_hdr_row].tolist()

    # Собираем сырые строки плат и филиала из хвоста
    raw_plates, raw_filial = [], []
    mode = 'plates'
    for r in range(last_hdr_row + 1, len(df_pos)):
        row = df_pos.iloc[r].tolist()
        c1 = row[0] if len(row) > 0 and pd.notna(row[0]) else None
        c2 = row[1] if len(row) > 1 and pd.notna(row[1]) else None

        if c1 is not None and 'филиал' in str(c1).lower():
            mode = 'filial'
            continue
        if c2 is None or not clean_int_str(c2).isdigit():
            continue

        ps_val = int(float(c2))
        schn_raw = row[3] if len(row) > 3 and pd.notna(row[3]) else 0
        schn_val = int(float(schn_raw)) if clean_int_str(schn_raw).isdigit() else clean_int_str(schn_raw)

        pol_raw = row[5] if len(row) > 5 and pd.notna(row[5]) else 0
        otg_raw = row[6] if len(row) > 6 and pd.notna(row[6]) else 0

        try:
            pol_val = int(float(str(pol_raw).replace('.', '').strip())) if str(pol_raw).replace('.', '').strip().isdigit() else pol_raw
        except Exception:
            pol_val = pol_raw

        try:
            otg_val = int(float(str(otg_raw).replace('.', '').strip())) if str(otg_raw).replace('.', '').strip().isdigit() else otg_raw
        except Exception:
            otg_val = otg_raw

        item = {
            'orig_idx': r,
            'Фирма': str(c1).strip() if c1 is not None else '',
            'PS': ps_val,
            'Сч/Н': schn_val,
            'Получено': pol_val,
            'Отгрузить!': otg_val,
            'на складе!': row[8] if len(row) > 8 and pd.notna(row[8]) else None,
            'долг': row[9] if len(row) > 9 and pd.notna(row[9]) else None,
            'Расходная накладная': row[11] if len(row) > 11 and pd.notna(row[11]) else None,
        }
        if mode == 'plates':
            raw_plates.append(item)
        else:
            raw_filial.append(item)

    if progress_callback:
        progress_callback(0.65, "Сопоставление заказов со сводками ПП и Филиала...")

    stats = {
        "total_rows": len(raw_plates) + len(raw_filial),
        "empty_ps_rows": 0,
        "already_filled": 0,
        "matched_head": 0,
        "matched_filial": 0,
        "not_found": 0,
        "duplicates_in_summaries": 0,
        "conflicts_both_sources": 0,
        "warnings": [],
        "row_statuses": {},
        "source_posylka_path": real_pos,
        "last_hdr_row": last_hdr_row,
        "header_banner_row": header_banner_row,
        "sorted_plates": [],
        "sorted_filial": []
    }

    # Обогащаем платы из сводки ПП
    processed_plates = []
    for it in raw_plates:
        firm, sch_pp, po_schetu, deliv = lookup_plate_info(df_plat, it['PS'], it['Сч/Н'])
        if firm or po_schetu is not None:
            stats["matched_head"] += 1
        else:
            stats["not_found"] += 1
            stats["warnings"].append(f"Печатные платы: не найден заказ PS {it['PS']} (Сч/Н {it['Сч/Н']})")

        processed_plates.append({
            **it,
            'Фирма': firm or it['Фирма'],
            'Счет пп   ': sch_pp,
            'По счету': po_schetu if po_schetu is not None else it['Отгрузить!'],
            'Этикетка': 'ОТК',
            'Доставка': deliv
        })

    # Обогащаем филиал из сводки ФИЛИАЛ
    processed_filial = []
    for it in raw_filial:
        po_schetu, deliv = lookup_filial_info(df_filial, it['PS'], it['Сч/Н'])
        if po_schetu is not None:
            stats["matched_filial"] += 1
        else:
            stats["not_found"] += 1
            stats["warnings"].append(f"Филиал: не найдено количество для PS {it['PS']} (Сч/Н {it['Сч/Н']})")

        processed_filial.append({
            **it,
            'Счет пп   ': None,
            'По счету': po_schetu if po_schetu is not None else it['Отгрузить!'],
            'Этикетка': None,
            'Доставка': deliv
        })

    if progress_callback:
        progress_callback(0.85, "Сортировка: ДМ → СДЭК → К → Самовывоз и группировка фирм...")

    # Сортируем оба блока: ДМ (1) -> СДЭК (2) -> К (3) -> Самовывоз (4), затем фирма к фирме
    sorted_plates = sorted(
        processed_plates,
        key=lambda x: (get_delivery_rank(x['Доставка']), clean_firm_name(x['Фирма']), x['PS'])
    )
    sorted_filial = sorted(
        processed_filial,
        key=lambda x: (get_delivery_rank(x['Доставка']), clean_firm_name(x['Фирма']), x['PS'])
    )

    stats["sorted_plates"] = sorted_plates
    stats["sorted_filial"] = sorted_filial

    # Формируем DataFrame для предпросмотра в интерфейсе
    preview_cols = [
        'Фирма', 'PS', 'Счет пп   ', 'Сч/Н', 'По счету', 'Получено', 'Отгрузить!',
        'Этикетка', 'на складе!', 'долг', 'Доставка', 'Расходная накладная'
    ]
    preview_rows = []
    row_statuses = {}
    r_i = 0

    banner_dict = {c: None for c in preview_cols}
    banner_dict['Сч/Н'] = header_banner_row[3] if len(header_banner_row) > 3 else None
    banner_dict['По счету'] = header_banner_row[4] if len(header_banner_row) > 4 else None
    preview_rows.append(banner_dict)
    row_statuses[r_i] = "banner_date"
    r_i += 1

    for p in sorted_plates:
        preview_rows.append({c: p.get(c) for c in preview_cols})
        row_statuses[r_i] = "samovyvoz_head" if p.get('Доставка') == 'Самовывоз' else "head"
        r_i += 1

    if sorted_filial:
        fil_banner = {c: None for c in preview_cols}
        fil_banner['Фирма'] = 'филиал'
        preview_rows.append(fil_banner)
        row_statuses[r_i] = "banner_filial"
        r_i += 1

        for f in sorted_filial:
            preview_rows.append({c: f.get(c) for c in preview_cols})
            row_statuses[r_i] = "samovyvoz_filial" if f.get('Доставка') == 'Самовывоз' else "filial"
            r_i += 1

    stats["row_statuses"] = row_statuses
    df_preview = pd.DataFrame(preview_rows, columns=preview_cols)

    if progress_callback:
        progress_callback(1.0, "Последняя посылка успешно сведена!")

    return df_preview, stats
