import os
import shutil
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def export_to_excel(df, output_path, row_statuses, make_backup=True, backup_dir=None):
    if make_backup and os.path.exists(output_path):
        b_dir = backup_dir if backup_dir else os.path.join(os.path.dirname(output_path), "Backups")
        if not os.path.exists(b_dir):
            os.makedirs(b_dir)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base_name = os.path.basename(output_path)
        shutil.copy2(output_path, os.path.join(b_dir, f"{ts}_{base_name}"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Сведенная посылка"

    # Заливки точно как на образцах
    fill_header_teal = PatternFill(start_color="008080", end_color="008080", fill_type="solid")
    fill_header_yellow = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
    fill_filial_banner = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    fill_samovyvoz_head = PatternFill(start_color="FF99CC", end_color="FF99CC", fill_type="solid")   # Розовый (Первый блок ПП)
    fill_samovyvoz_filial = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid") # Оранжевый (Второй блок Филиал)
    fill_not_found = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    # Шрифты (Calibri 11 Bold)
    font_header_black = Font(name="Calibri", size=11, bold=True, color="000000")
    font_header_red = Font(name="Calibri", size=11, bold=True, color="FF0000")
    font_white_bold = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_black_bold = Font(name="Calibri", size=11, bold=True, color="000000")
    font_red_bold = Font(name="Calibri", size=11, bold=True, color="FF0000")
    font_blue_bold = Font(name="Calibri", size=11, bold=True, color="0000FF")

    # Жирная чёрная граница для всех ячеек
    thick_black_border = Border(
        left=Side(style='medium', color='000000'),
        right=Side(style='medium', color='000000'),
        top=Side(style='medium', color='000000'),
        bottom=Side(style='medium', color='000000')
    )

    headers = list(df.columns)
    ws.append(headers)
    ws.row_dimensions[1].height = 24

    for c_idx, h_name in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.border = thick_black_border
        cell.alignment = Alignment(horizontal="center", vertical="center")
        h_clean = str(h_name).strip()
        if h_clean in ("Отгрузить!", "на складе!"):
            cell.fill = fill_header_yellow
            cell.font = font_header_red if h_clean == "Отгрузить!" else font_header_black
        else:
            cell.fill = fill_header_teal
            cell.font = font_header_black

    # Индексы колонок
    col_idx_firma = 1
    col_idx_schet_pp = next((i + 1 for i, c in enumerate(headers) if "Счет пп" in str(c)), 3)
    col_idx_otgruzit = next((i + 1 for i, c in enumerate(headers) if "Отгрузить" in str(c)), 7)
    col_idx_otk = next((i + 1 for i, c in enumerate(headers) if "Этикетка" in str(c)), 8)
    col_idx_dostavka = next((i + 1 for i, c in enumerate(headers) if "Доставка" in str(c)), 11)

    for df_idx, row_series in df.iterrows():
        r_idx = df_idx + 2
        row_vals = [None if str(v) in ('nan', 'None', 'NaT', '') else v for v in row_series.tolist()]
        ws.append(row_vals)
        ws.row_dimensions[r_idx].height = 20
        status = row_statuses.get(df_idx, "")

        for c_idx in range(1, len(row_vals) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thick_black_border

            # Форматируем "Счет пп" как текст с сохранением ведущего нуля (0682)
            if c_idx == col_idx_schet_pp and cell.value is not None:
                val_s = str(cell.value).strip()
                if val_s.endswith('.0'):
                    val_s = val_s[:-2]
                if val_s.isdigit() and len(val_s) < 4:
                    val_s = val_s.zfill(4)
                cell.value = val_s
                cell.number_format = '@'

            if status == "banner_date":
                cell.font = font_red_bold
                cell.alignment = Alignment(horizontal="center", vertical="center")
            elif status == "banner_filial":
                if c_idx == 1:
                    cell.fill = fill_filial_banner
                    cell.font = font_white_bold
                    cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Фоновая заливка строки
                if status == "samovyvoz_head":
                    cell.fill = fill_samovyvoz_head
                elif status == "samovyvoz_filial":
                    cell.fill = fill_samovyvoz_filial
                elif status == "not_found":
                    cell.fill = fill_not_found

                # Шрифт и выравнивание по колонкам
                if c_idx == col_idx_firma:
                    cell.font = font_black_bold
                    cell.alignment = Alignment(horizontal="left", vertical="center")
                elif c_idx in (col_idx_schet_pp, col_idx_otgruzit, col_idx_dostavka):
                    cell.font = font_red_bold
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                elif c_idx == col_idx_otk:
                    cell.font = font_blue_bold
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    cell.font = font_black_bold
                    cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.freeze_panes = "A2"

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = min(max(max_len + 4, 11), 42)

    wb.save(output_path)
    return output_path
