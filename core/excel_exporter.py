import os
import shutil
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ==============================================================================
# НАСТРОЙКИ СТИЛЕЙ И ОФОРМЛЕНИЯ (СТАНДАРТ ТАБЛИЦЫ ПОСЫЛКИ)
# ==============================================================================
FONT_FIRM = Font(name='Tahoma', size=18, bold=True)
FONT_DATA_BLACK = Font(name='Tahoma', size=16, bold=True, color='000000')
FONT_DATA_RED = Font(name='Tahoma', size=16, bold=True, color='FF0000')
FONT_DATA_BLUE = Font(name='Tahoma', size=16, bold=True, color='0000FF')
FONT_FILIAL_HDR = Font(name='Tahoma', size=18, bold=True, color='FFFFFF')

FILL_SAMOVYVOZ = PatternFill(start_color='FFFFC000', end_color='FFFFC000', fill_type='solid')  # Оранжевый #FFC000
FILL_OTK = PatternFill(start_color='FFFF99CC', end_color='FFFF99CC', fill_type='solid')        # Розовый #FF99CC
FILL_FILIAL = PatternFill(start_color='FFFF0000', end_color='FFFF0000', fill_type='solid')     # Красный #FF0000

THICK_SIDE = Side(style='thick')
THICK_BORDER = Border(left=THICK_SIDE, right=THICK_SIDE, top=THICK_SIDE, bottom=THICK_SIDE)

ALIGN_CENTER = Alignment(horizontal='center', vertical='center')
ALIGN_LEFT = Alignment(horizontal='left', vertical='center')

COLS_ORDER = [
    'Фирма', 'PS', 'Счет пп   ', 'Сч/Н', 'По счету', 'Получено', 'Отгрузить!',
    'Этикетка', 'на складе!', 'долг', 'Доставка', 'Расходная накладная'
]


def write_styled_row(ws, row_idx, row_dict):
    is_samovyvoz = (row_dict.get('Доставка') == 'Самовывоз')
    ws.row_dimensions[row_idx].height = 28

    for c_idx, col_name in enumerate(COLS_ORDER, start=1):
        cell = ws.cell(row=row_idx, column=c_idx)
        val = row_dict.get(col_name)
        cell.value = val
        cell.border = THICK_BORDER

        if is_samovyvoz:
            cell.fill = FILL_SAMOVYVOZ

        if c_idx == 1:  # Фирма
            cell.font = FONT_FIRM
            cell.alignment = ALIGN_LEFT
        else:
            cell.alignment = ALIGN_CENTER
            if c_idx == 3:  # Счет пп (4 знака с ведущим нулём)
                cell.font = FONT_DATA_RED
                cell.number_format = '@'
                if val is not None and str(val).strip() != '':
                    cell.value = str(val).zfill(4)
            elif c_idx == 7:  # Отгрузить!
                cell.font = FONT_DATA_RED
            elif c_idx == 8:  # Этикетка
                if val and 'ОТК' in str(val):
                    cell.font = FONT_DATA_BLUE
                    cell.fill = FILL_OTK
                else:
                    cell.font = FONT_DATA_BLACK
            elif c_idx == 11:  # Доставка
                d_val = str(val).strip() if val else ''
                cell.font = FONT_DATA_RED if d_val in ['ДМ', 'К', 'Самовывоз'] else FONT_DATA_BLACK
            else:
                cell.font = FONT_DATA_BLACK


def export_to_excel(df, output_path, row_statuses, make_backup=True, backup_dir=None, stats=None):
    """
    Сохраняет итоговый Excel-файл с полной сохранностью шапки, стилей и структуры листа 'Посылка'.
    """
    if make_backup and os.path.exists(output_path):
        b_dir = backup_dir if backup_dir else os.path.join(os.path.dirname(output_path), "Backups")
        if not os.path.exists(b_dir):
            os.makedirs(b_dir)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.basename(output_path)
        shutil.copy2(output_path, os.path.join(b_dir, f"{ts}_{base_name}"))

    # Если передана полная статистика из run_reconciliation, редактируем копию исходной книги с сохранением всех стилей
    if stats and stats.get("source_posylka_path") and os.path.exists(stats["source_posylka_path"]):
        wb = openpyxl.load_workbook(stats["source_posylka_path"])
        ws = wb['Посылка'] if 'Посылка' in wb.sheetnames else wb.active

        last_hdr_row = stats["last_hdr_row"] + 1  # перевод из 0-based pandas в 1-based openpyxl
        sorted_plates = stats.get("sorted_plates", [])
        sorted_filial = stats.get("sorted_filial", [])

        # Сохраняем всю строку заголовка целиком (включая жёлтую "Субботняя отгрузка" #FFFF00 и др.)
        header_snapshot = []
        for c in range(1, 13):
            cell = ws.cell(row=last_hdr_row, column=c)
            header_snapshot.append((
                c,
                cell.value,
                cell.font.copy(),
                cell.fill.copy(),
                cell.border.copy(),
                cell.alignment.copy()
            ))

        # Очищаем хвост начиная с last_hdr_row и записываем идеально выстроенный блок
        if ws.max_row >= last_hdr_row:
            ws.delete_rows(last_hdr_row, ws.max_row - last_hdr_row + 1)

        # 1. Восстанавливаем неприкосновенную шапку даты (сохраняя "Субботняя отгрузка" и цвета)
        for c, val, fnt, fl, brd, aln in header_snapshot:
            cell = ws.cell(row=last_hdr_row, column=c)
            cell.value = val
            cell.font = fnt
            cell.fill = fl
            cell.border = brd
            cell.alignment = aln

        curr_r = last_hdr_row + 2  # 1 пустая строка после заголовка

        # 2. Записываем платы
        for p_row in sorted_plates:
            write_styled_row(ws, curr_r, p_row)
            curr_r += 1

        if sorted_filial:
            curr_r += 1  # 1 пустая строка перед плашкой филиал

            # 3. Записываем красную плашку "филиал"
            ws.row_dimensions[curr_r].height = 28
            c_fil = ws.cell(row=curr_r, column=1)
            c_fil.value = 'филиал'
            c_fil.font = FONT_FILIAL_HDR
            c_fil.fill = FILL_FILIAL
            c_fil.border = THICK_BORDER
            c_fil.alignment = ALIGN_CENTER

            curr_r += 2  # 1 пустая строка после плашки филиал

            # 4. Записываем филиал
            for f_row in sorted_filial:
                write_styled_row(ws, curr_r, f_row)
                curr_r += 1

        wb.save(output_path)
        return output_path

    # Резервный экспорт, если вызван без stats
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Посылка"
    ws.append(COLS_ORDER)
    for idx, row_series in df.iterrows():
        write_styled_row(ws, idx + 2, row_series.to_dict())
    wb.save(output_path)
    return output_path
