import os
import shutil
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
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
    ws.title = "Сведенная накладная"

    header_fill = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    fill_head = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    fill_filial = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid")
    fill_conflict = PatternFill(start_color="FEF08A", end_color="FEF08A", fill_type="solid")
    fill_not_found = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
        ws.append(row)
        df_idx = r_idx - 2

        for c_idx in range(1, len(row) + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            cell.border = thin_border

            if r_idx == 1:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            else:
                cell.font = Font(name="Calibri", size=10)
                status = row_statuses.get(df_idx)
                if status == "head":
                    cell.fill = fill_head
                elif status == "filial":
                    cell.fill = fill_filial
                elif status == "conflict":
                    cell.fill = fill_conflict
                elif status == "not_found":
                    cell.fill = fill_not_found

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    for col in ws.columns:
        max_length = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            try:
                val_len = len(str(cell.value or ""))
                if val_len > max_length:
                    max_length = val_len
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max(max_length + 4, 12), 45)

    wb.save(output_path)
    return output_path
