import flet as ft
import pandas as pd
import numpy as np

def main(page: ft.Page):
    # Настройки окна программы
    page.title = "Автоматическое сведение накладных - pcb-auto-excel"
    page.theme_mode = ft.ThemeMode.DARK
    page.window_width = 1000
    page.window_height = 700
    page.padding = 30

    processed_df = None

    # Обработчики выбора файлов
    def pick_posylka_result(e: ft.FilePickerResultEvent):
        if e.files:
            posylka_input.value = e.files[0].path
            page.update()

    def pick_pp_result(e: ft.FilePickerResultEvent):
        if e.files:
            pp_input.value = e.files[0].path
            page.update()

    def pick_filial_result(e: ft.FilePickerResultEvent):
        if e.files:
            filial_input.value = e.files[0].path
            page.update()

    def save_file_result(e: ft.FilePickerResultEvent):
        nonlocal processed_df
        if e.path and processed_df is not None:
            save_path = e.path if e.path.endswith('.xlsx') else e.path + '.xlsx'
            try:
                processed_df.to_excel(save_path, index=False)
                page.overlay.append(ft.SnackBar(ft.Text(f"Успешно сохранено:\n{save_path}"), bgcolor=ft.colors.GREEN_700, open=True))
            except Exception as ex:
                page.overlay.append(ft.SnackBar(ft.Text(f"Ошибка сохранения: {ex}"), bgcolor=ft.colors.RED_700, open=True))
            page.update()

    # Невидимые компоненты для диалоговых окон Windows
    posylka_picker = ft.FilePicker(on_result=pick_posylka_result)
    pp_picker = ft.FilePicker(on_result=pick_pp_result)
    filial_picker = ft.FilePicker(on_result=pick_filial_result)
    save_picker = ft.FilePicker(on_result=save_file_result)
    page.overlay.extend([posylka_picker, pp_picker, filial_picker, save_picker])

    # Поля ввода путей
    posylka_input = ft.TextField(label="1. Путь к файлу: Посылка", expand=True, border_color=ft.colors.BLUE_GREY_700)
    pp_input = ft.TextField(label="2. Путь к файлу: Сводка ПП (Голова)", expand=True, border_color=ft.colors.BLUE_GREY_700)
    filial_input = ft.TextField(label="3. Путь к файлу: Сводка Филиал", expand=True, border_color=ft.colors.BLUE_GREY_700)

    # Текстовые поля для результатов
    stats_text = ft.Text("Голова: 0  |  Филиал: 0", weight=ft.FontWeight.BOLD, size=16, color=ft.colors.GREEN_400)
    preview_box = ft.TextField(multiline=True, read_only=True, expand=True, text_size=12, text_style=ft.TextStyle(font_family="Consolas"))

    # Логика обработки Excel
    def process_data(e):
        nonlocal processed_df
        p_pos = posylka_input.value
        p_pp = pp_input.value
        p_fil = filial_input.value

        if not p_pos or not p_pp or not p_fil:
            page.overlay.append(ft.SnackBar(ft.Text("Пожалуйста, укажите все 3 файла!"), bgcolor=ft.colors.RED_700, open=True))
            page.update()
            return

        btn_process.disabled = True
        btn_process.text = "ОБРАБОТКА..."
        page.update()

        try:
            # Чтение файлов
            df_posylka = pd.read_excel(p_pos.strip(' "'))
            df_pp = pd.read_excel(p_pp.strip(' "'), sheet_name=0, header=None)
            df_filial = pd.read_excel(p_fil.strip(' "'), sheet_name=0, header=None)

            # Очистка колонки PS
            df_posylka['PS_str'] = df_posylka['PS'].astype(str).str.strip().str.replace('.0', '', regex=False)
            
            # Словари из сводок
            df_pp['PS_str'] = df_pp[2].astype(str).str.strip().str.replace('.0', '', regex=False)
            pp_dict = df_pp.dropna(subset=['PS_str']).set_index('PS_str').to_dict('index')
            
            df_filial['PS_str'] = df_filial[4].astype(str).str.strip().str.replace('.0', '', regex=False)
            filial_dict = df_filial.dropna(subset=['PS_str']).set_index('PS_str').to_dict('index')

            count_head = 0
            count_filial = 0

            # Сведение
            for idx, row in df_posylka.iterrows():
                ps = str(row['PS_str'])
                if ps == 'nan' or not ps:
                    continue
                
                needs_update = pd.isna(row['По счету']) or pd.isna(row['Доставка'])
                
                if needs_update:
                    if ps in pp_dict:
                        if pd.isna(row['По счету']):
                            df_posylka.at[idx, 'По счету'] = pp_dict[ps].get(4, np.nan)
                        if pd.isna(row['Доставка']):
                            df_posylka.at[idx, 'Доставка'] = pp_dict[ps].get(14, np.nan)
                        count_head += 1
                            
                    elif ps in filial_dict:
                        if pd.isna(row['По счету']):
                            df_posylka.at[idx, 'По счету'] = filial_dict[ps].get(7, np.nan)
                        if pd.isna(row['Доставка']):
                            df_posylka.at[idx, 'Доставка'] = filial_dict[ps].get(1, np.nan)
                        count_filial += 1

            df_posylka.drop(columns=['PS_str'], inplace=True)
            processed_df = df_posylka
            
            # Формирование превью
            preview_csv = df_posylka.to_csv(sep=';', index=False, na_rep='')
            stats_text.value = f"Голова: {count_head}  |  Филиал: {count_filial}"
            preview_box.value = preview_csv
            btn_download.disabled = False
            
            page.overlay.append(ft.SnackBar(ft.Text("Данные успешно сведены!"), bgcolor=ft.colors.GREEN_700, open=True))

        except Exception as ex:
            page.overlay.append(ft.SnackBar(ft.Text(f"Ошибка: {str(ex)}"), bgcolor=ft.colors.RED_700, open=True))

        btn_process.disabled = False
        btn_process.text = "▶ СВЕСТИ НАКЛАДНУЮ"
        page.update()

    # Кнопки
    btn_process = ft.ElevatedButton(
        "▶ СВЕСТИ НАКЛАДНУЮ", 
        on_click=process_data,
        style=ft.ButtonStyle(color=ft.colors.WHITE, bgcolor=ft.colors.BLUE_700, shape=ft.RoundedRectangleBorder(radius=8)),
        height=50, expand=True
    )
    
    btn_download = ft.ElevatedButton(
        "📥 СКАЧАТЬ EXCEL", 
        disabled=True,
        on_click=lambda _: save_picker.save_file(allowed_extensions=["xlsx"], file_name="Посылка_СВЕДЕНО.xlsx"),
        style=ft.ButtonStyle(color=ft.colors.WHITE, bgcolor=ft.colors.GREEN_600, shape=ft.RoundedRectangleBorder(radius=8)),
        height=50, expand=True
    )

    # Построение интерфейса
    page.add(
        ft.Text("Автоматическое сведение накладных", size=24, weight=ft.FontWeight.BOLD),
        ft.Text("Выбери файлы через проводник или вставь прямой путь", color=ft.colors.WHITE54),
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        
        ft.Row([posylka_input, ft.IconButton(ft.icons.FOLDER_OPEN, icon_size=30, on_click=lambda _: posylka_picker.pick_files())]),
        ft.Row([pp_input, ft.IconButton(ft.icons.FOLDER_OPEN, icon_size=30, on_click=lambda _: pp_picker.pick_files())]),
        ft.Row([filial_input, ft.IconButton(ft.icons.FOLDER_OPEN, icon_size=30, on_click=lambda _: filial_picker.pick_files())]),
        
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        ft.Row([btn_process, ft.Container(width=20), btn_download]),
        ft.Divider(height=20, color=ft.colors.TRANSPARENT),
        
        ft.Row([ft.Text("Результат сведения:", weight=ft.FontWeight.BOLD, color=ft.colors.GREEN_400), stats_text], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
        preview_box
    )

ft.app(target=main)
