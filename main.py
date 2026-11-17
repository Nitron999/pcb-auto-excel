import os
import subprocess
import threading
from datetime import datetime
import flet as ft
import pandas as pd

from core.storage import load_settings, save_settings, load_history, add_history_record
from core.validator import get_file_info, validate_excel_files
from core.reconciler import run_reconciliation
from core.excel_exporter import export_to_excel

def main(page: ft.Page):
    settings = load_settings()

    page.title = "PCB Auto Excel — Система сведения накладных"
    page.theme_mode = ft.ThemeMode.DARK if settings.get("theme", "DARK") == "DARK" else ft.ThemeMode.LIGHT
    page.window.width = 1280
    page.window.height = 820
    page.window.min_width = 1000
    page.window.min_height = 680
    page.padding = 0
    page.bgcolor = "#0F172A" if page.theme_mode == ft.ThemeMode.DARK else "#F8FAFC"

    state = {
        "posylka_path": settings.get("last_posylka", "") if settings.get("auto_load_files", True) else "",
        "pp_path": settings.get("last_pp", "") if settings.get("auto_load_files", True) else "",
        "filial_path": settings.get("last_filial", "") if settings.get("auto_load_files", True) else "",
        "result_df": None,
        "last_stats": None,
        "is_processing": False
    }

    def notify(message: str, kind: str = "info"):
        colors = {
            "success": ft.Colors.GREEN_700,
            "warning": ft.Colors.AMBER_800,
            "error": ft.Colors.RED_700,
            "info": ft.Colors.BLUE_700
        }
        snack = ft.SnackBar(
            content=ft.Text(message, color=ft.Colors.WHITE, weight=ft.FontWeight.W_500),
            bgcolor=colors.get(kind, ft.Colors.BLUE_700),
            duration=5000
        )
        page.overlay.append(snack)
        snack.open = True
        page.update()

    def on_pick_posylka(e: ft.FilePickerResultEvent):
        if e.files:
            update_file_path("posylka_path", e.files[0].path)

    def on_pick_pp(e: ft.FilePickerResultEvent):
        if e.files:
            update_file_path("pp_path", e.files[0].path)

    def on_pick_filial(e: ft.FilePickerResultEvent):
        if e.files:
            update_file_path("filial_path", e.files[0].path)

    def on_pick_save_dir(e: ft.FilePickerResultEvent):
        if e.path:
            settings["save_dir"] = e.path
            save_settings(settings)
            txt_save_dir.value = e.path
            notify(f"Папка сохранения изменена на: {e.path}", "info")
            page.update()

    def on_save_excel_dialog(e: ft.FilePickerResultEvent):
        if e.path and state["result_df"] is not None:
            out_path = e.path if e.path.endswith(".xlsx") else e.path + ".xlsx"
            perform_save(out_path)

    picker_posylka = ft.FilePicker(on_result=on_pick_posylka)
    picker_pp = ft.FilePicker(on_result=on_pick_pp)
    picker_filial = ft.FilePicker(on_result=on_pick_filial)
    picker_save_dir = ft.FilePicker(on_result=on_pick_save_dir)
    picker_save_file = ft.FilePicker(on_result=on_save_excel_dialog)
    page.services.extend([picker_posylka, picker_pp, picker_filial, picker_save_dir, picker_save_file])

    def update_file_path(key: str, new_path: str):
        clean = new_path.strip(' "') if new_path else ""
        state[key] = clean
        if key == "posylka_path":
            settings["last_posylka"] = clean
        elif key == "pp_path":
            settings["last_pp"] = clean
        elif key == "filial_path":
            settings["last_filial"] = clean
        save_settings(settings)
        refresh_file_cards()
        page.update()

    file_cards_container = ft.Column(spacing=15)

    def build_file_card(title: str, subtitle: str, state_key: str, picker: ft.FilePicker):
        info = get_file_info(state[state_key])
        status_color = ft.Colors.GREEN_400 if info["exists"] else ft.Colors.RED_400
        status_icon = ft.Icons.CHECK_CIRCLE if info["exists"] else ft.Icons.ERROR_OUTLINE

        path_input = ft.TextField(
            value=state[state_key],
            hint_text=r"Вставьте сетевой путь \\Fs01\... или выберите файл",
            expand=True,
            text_size=13,
            border_color="#334155",
            on_submit=lambda e: update_file_path(state_key, e.control.value),
            on_blur=lambda e: update_file_path(state_key, e.control.value)
        )

        return ft.Container(
            bgcolor="#1E293B",
            border_radius=12,
            padding=18,
            border=ft.border.all(1, "#334155"),
            content=ft.Column(
                spacing=10,
                controls=[
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=2,
                                controls=[
                                    ft.Text(title, size=16, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                                    ft.Text(subtitle, size=12, color=ft.Colors.WHITE_54),
                                ]
                            ),
                            ft.Row(
                                spacing=6,
                                controls=[
                                    ft.Icon(status_icon, color=status_color, size=18),
                                    ft.Text(info["status"], color=status_color, size=12, weight=ft.FontWeight.W_600)
                                ]
                            )
                        ]
                    ),
                    ft.Row(
                        spacing=10,
                        controls=[
                            path_input,
                            ft.Button(
                                content=ft.Row([ft.Icon(ft.Icons.FOLDER_OPEN, size=18), ft.Text("Обзор")], spacing=6),
                                on_click=lambda _: picker.pick_files(allowed_extensions=["xlsx", "xls"]),
                                style=ft.ButtonStyle(bgcolor="#2563EB", color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=8)),
                                height=44
                            ),
                            ft.IconButton(
                                icon=ft.Icons.CLEAR,
                                tooltip="Очистить путь",
                                icon_color=ft.Colors.RED_400,
                                on_click=lambda _: update_file_path(state_key, "")
                            )
                        ]
                    ),
                    ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Text(f"Файл: {info['name']}", size=12, color=ft.Colors.WHITE_70),
                            ft.Text(f"Изменён: {info['date']}", size=12, color=ft.Colors.WHITE_54),
                        ]
                    )
                ]
            )
        )

    def refresh_file_cards():
        file_cards_container.controls = [
            build_file_card("1. Накладная (Посылка)", "Основной файл с пустыми позициями в конце списка", "posylka_path", picker_posylka),
            build_file_card("2. Сводка ПП (Голова)", "База заказов головного офиса (ООО Печатные платы)", "pp_path", picker_pp),
            build_file_card("3. Сводка Филиала", "База заказов филиала (ЭЛЕКТРОконнект)", "filial_path", picker_filial)
        ]

    progress_bar = ft.ProgressBar(value=0, color="#3B82F6", bgcolor="#1E293B", visible=False)
    progress_label = ft.Text("", size=13, color=ft.Colors.BLUE_300, visible=False)
    
    results_summary_box = ft.Container(visible=False)
    warnings_column = ft.Column(spacing=5)

    def perform_save(out_path: str):
        try:
            saved_file = export_to_excel(
                state["result_df"],
                out_path,
                state["last_stats"]["row_statuses"],
                make_backup=settings.get("backup_results", True)
            )
            add_history_record({
                "posylka": os.path.basename(state["posylka_path"]),
                "pp": os.path.basename(state["pp_path"]),
                "filial": os.path.basename(state["filial_path"]),
                "total_rows": state["last_stats"]["total_rows"],
                "matched": state["last_stats"]["matched_head"] + state["last_stats"]["matched_filial"],
                "warnings_count": len(state["last_stats"]["warnings"]),
                "output_path": saved_file,
                "status": "Успешно"
            })
            notify(f"Итоговый Excel сохранён:\n{saved_file}", "success")
            refresh_dashboard()
            refresh_history()
        except Exception as ex:
            notify(f"Ошибка при сохранении Excel: {str(ex)}", "error")

    def on_save_click(e):
        if state["result_df"] is None:
            notify("Сначала выполните сведение накладных!", "warning")
            return
        default_name = f"Посылка_СВЕДЕНО_{datetime.now().strftime('%d.%m.%Y')}.xlsx"
        picker_save_file.save_file(
            dialog_title="Сохранить сведённую накладную",
            file_name=default_name,
            initial_directory=settings.get("save_dir", ""),
            allowed_extensions=["xlsx"]
        )

    btn_save_excel = ft.Button(
        content=ft.Row([ft.Icon(ft.Icons.DOWNLOAD, size=20), ft.Text("СОХРАНИТЬ ИТОГОВЫЙ EXCEL (.XLSX)", weight=ft.FontWeight.BOLD)], alignment=ft.MainAxisAlignment.CENTER),
        on_click=on_save_click,
        disabled=True,
        style=ft.ButtonStyle(bgcolor="#10B981", color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=10)),
        height=52,
        expand=True
    )

    def start_reconciliation_thread(e):
        if state["is_processing"]:
            return

        is_valid, msg = validate_excel_files(state["posylka_path"], state["pp_path"], state["filial_path"])
        if not is_valid:
            notify(msg, "error")
            return

        state["is_processing"] = True
        btn_reconcile.disabled = True
        btn_save_excel.disabled = True
        progress_bar.visible = True
        progress_bar.value = 0.05
        progress_label.visible = True
        progress_label.value = "Запуск проверки и чтения таблиц..."
        results_summary_box.visible = False
        warnings_column.controls.clear()
        page.update()

        def worker():
            try:
                def update_progress(val, text):
                    progress_bar.value = val
                    progress_label.value = text
                    page.update()

                df_res, stats = run_reconciliation(
                    state["posylka_path"],
                    state["pp_path"],
                    state["filial_path"],
                    progress_callback=update_progress
                )

                state["result_df"] = df_res
                state["last_stats"] = stats

                total_matched = stats["matched_head"] + stats["matched_filial"]
                results_summary_box.content = ft.Container(
                    bgcolor="#1E293B",
                    border_radius=12,
                    padding=20,
                    border=ft.border.all(1, "#334155"),
                    content=ft.Column(
                        spacing=15,
                        controls=[
                            ft.Row(
                                alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                controls=[
                                    ft.Text("Отчёт о сведении накладных", size=18, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                                    ft.Container(
                                        bgcolor="#065F46",
                                        padding=ft.padding.symmetric(horizontal=12, vertical=6),
                                        border_radius=8,
                                        content=ft.Text("ВЫПОЛНЕНО", color="#34D399", weight=ft.FontWeight.BOLD, size=12)
                                    )
                                ]
                            ),
                            ft.Row(
                                spacing=15,
                                controls=[
                                    stat_mini_card("Всего строк", str(stats["total_rows"]), ft.Colors.BLUE_400),
                                    stat_mini_card("Голова (ПП)", str(stats["matched_head"]), ft.Colors.GREEN_400),
                                    stat_mini_card("Филиал", str(stats["matched_filial"]), ft.Colors.CYAN_400),
                                    stat_mini_card("Не найдено", str(stats["not_found"]), ft.Colors.RED_400),
                                    stat_mini_card("Дубликаты", str(stats["duplicates_in_summaries"]), ft.Colors.AMBER_400),
                                ]
                            ),
                            warnings_column
                        ]
                    )
                )

                if stats["warnings"]:
                    warnings_column.controls.append(ft.Text("Предупреждения и расхождения:", weight=ft.FontWeight.BOLD, color=ft.Colors.AMBER_400, size=13))
                    for w in stats["warnings"]:
                        warnings_column.controls.append(ft.Text(f"• {w}", color=ft.Colors.AMBER_200, size=12))

                results_summary_box.visible = True
                btn_save_excel.disabled = False

                if stats["warnings"]:
                    notify(f"Сведение завершено с предупреждениями ({len(stats['warnings'])}). Найдено совпадений: {total_matched}.", "warning")
                else:
                    notify(f"Сведение успешно завершено! Заполнено позиций: {total_matched}.", "success")

            except Exception as ex:
                notify(f"Критическая ошибка при обработке: {str(ex)}", "error")
            finally:
                state["is_processing"] = False
                btn_reconcile.disabled = False
                progress_bar.visible = False
                progress_label.visible = False
                page.update()

        threading.Thread(target=worker, daemon=True).start()

    btn_reconcile = ft.Button(
        content=ft.Row([ft.Icon(ft.Icons.PLAY_ARROW, size=22), ft.Text("СВЕСТИ НАКЛАДНЫЕ", weight=ft.FontWeight.BOLD)], alignment=ft.MainAxisAlignment.CENTER),
        on_click=start_reconciliation_thread,
        style=ft.ButtonStyle(bgcolor="#2563EB", color=ft.Colors.WHITE, shape=ft.RoundedRectangleBorder(radius=10)),
        height=52,
        expand=True
    )

    def stat_mini_card(label: str, value: str, color):
        return ft.Container(
            expand=True,
            bgcolor="#0F172A",
            padding=12,
            border_radius=10,
            border=ft.border.all(1, "#334155"),
            content=ft.Column(
                spacing=4,
                controls=[
                    ft.Text(label, size=12, color=ft.Colors.WHITE_54),
                    ft.Text(value, size=22, weight=ft.FontWeight.BOLD, color=color)
                ]
            )
        )

    dashboard_cards = ft.Column(spacing=20)

    def refresh_dashboard():
        history = load_history()
        total_ops = len(history)
        total_rows_all = sum(h.get("total_rows", 0) for h in history)
        total_matched_all = sum(h.get("matched", 0) for h in history)
        total_warnings_all = sum(h.get("warnings_count", 0) for h in history)
        last_op = history[0] if history else None

        dashboard_cards.controls = [
            ft.Text("Панель управления", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Text("Сводная статистика работы системы автоматического сведения накладных", size=14, color=ft.Colors.WHITE_54),
            ft.Row(
                spacing=15,
                controls=[
                    stat_mini_card("Выполнено сведений", str(total_ops), ft.Colors.BLUE_400),
                    stat_mini_card("Всего обработано строк", str(total_rows_all), ft.Colors.WHITE),
                    stat_mini_card("Найдено совпадений", str(total_matched_all), ft.Colors.GREEN_400),
                    stat_mini_card("Предупреждений", str(total_warnings_all), ft.Colors.AMBER_400),
                ]
            ),
            ft.Container(
                bgcolor="#1E293B",
                border_radius=12,
                padding=20,
                border=ft.border.all(1, "#334155"),
                content=ft.Column(
                    spacing=10,
                    controls=[
                        ft.Text("Последнее сведение", size=18, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
                        ft.Divider(color="#334155"),
                        ft.Text(
                            f"Дата и время: {last_op['timestamp']}\n"
                            f"Исходный файл: {last_op['posylka']}\n"
                            f"Найдено совпадений: {last_op['matched']} из {last_op['total_rows']} строк\n"
                            f"Итоговый файл: {last_op['output_path']}"
                            if last_op else "Операций сведения ещё не проводилось. Перейдите в раздел «Сведение накладных».",
                            size=14,
                            color=ft.Colors.WHITE_70
                        )
                    ]
                )
            )
        ]

    refresh_file_cards()
    view_reconcile = ft.Column(
        scroll=ft.ScrollMode.AUTO,
        spacing=20,
        controls=[
            ft.Text("Сведение накладных", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Text("Укажите сетевые или локальные пути к таблицам Excel для автоматического заполнения", size=14, color=ft.Colors.WHITE_54),
            file_cards_container,
            ft.Row([btn_reconcile, btn_save_excel], spacing=15),
            progress_label,
            progress_bar,
            results_summary_box
        ]
    )

    preview_source_dropdown = ft.Dropdown(
        label="Выберите таблицу для просмотра",
        width=320,
        value="result",
        options=[
            ft.dropdown.Option("result", "Результат сведения (Итоговая таблица)"),
            ft.dropdown.Option("posylka", "Исходная Посылка"),
            ft.dropdown.Option("pp", "Сводка ПП (Голова)"),
            ft.dropdown.Option("filial", "Сводка Филиал")
        ]
    )
    preview_search = ft.TextField(label="Поиск по таблице (PS, фирма, доставка...)", expand=True)
    preview_table_container = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True)
    preview_info_label = ft.Text("Отображается до 100 строк для быстрого просмотра", size=12, color=ft.Colors.WHITE_54)

    def load_preview_data(e=None):
        src = preview_source_dropdown.value
        query = (preview_search.value or "").strip().lower()
        df = None

        try:
            if src == "result":
                df = state["result_df"]
                if df is None:
                    preview_table_container.controls = [ft.Text("Сначала выполните сведение накладных, чтобы увидеть результат.", color=ft.Colors.AMBER_400)]
                    page.update()
                    return
            elif src == "posylka" and os.path.exists(state["posylka_path"]):
                df = pd.read_excel(state["posylka_path"])
            elif src == "pp" and os.path.exists(state["pp_path"]):
                df = pd.read_excel(state["pp_path"])
            elif src == "filial" and os.path.exists(state["filial_path"]):
                df = pd.read_excel(state["filial_path"])
            else:
                preview_table_container.controls = [ft.Text("Выбранный файл не найден или не указан.", color=ft.Colors.RED_400)]
                page.update()
                return

            if query:
                mask = df.astype(str).apply(lambda col: col.str.lower().str.contains(query, na=False)).any(axis=1)
                df_filtered = df[mask]
            else:
                df_filtered = df

            preview_info_label.value = f"Всего строк в таблице: {len(df)} | Найдено по фильтру: {len(df_filtered)} (показаны первые 100)"
            df_show = df_filtered.head(100).fillna("")

            cols = [ft.DataColumn(ft.Text(str(c), weight=ft.FontWeight.BOLD, size=12)) for c in df_show.columns[:12]]
            rows = []
            for _, r in df_show.iterrows():
                cells = [ft.DataCell(ft.Text(str(val)[:40], size=12)) for val in list(r)[:12]]
                rows.append(ft.DataRow(cells=cells))

            table = ft.DataTable(
                columns=cols,
                rows=rows,
                border=ft.border.all(1, "#334155"),
                heading_row_color="#1E293B",
                data_row_min_height=36
            )
            preview_table_container.controls = [ft.Row([table], scroll=ft.ScrollMode.AUTO)]
        except Exception as ex:
            preview_table_container.controls = [ft.Text(f"Ошибка чтения таблицы: {ex}", color=ft.Colors.RED_400)]
        page.update()

    preview_source_dropdown.on_change = load_preview_data
    preview_search.on_submit = load_preview_data

    view_preview = ft.Column(
        expand=True,
        spacing=15,
        controls=[
            ft.Text("Предварительный просмотр таблиц", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Row(
                spacing=15,
                controls=[
                    preview_source_dropdown,
                    preview_search,
                    ft.Button(
                        content=ft.Text("Обновить / Найти"),
                        on_click=load_preview_data,
                        style=ft.ButtonStyle(bgcolor="#2563EB", color=ft.Colors.WHITE),
                        height=48
                    )
                ]
            ),
            preview_info_label,
            preview_table_container
        ]
    )

    history_list = ft.Column(scroll=ft.ScrollMode.AUTO, spacing=10, expand=True)

    def open_folder(file_path):
        folder = os.path.dirname(file_path)
        if os.path.exists(folder):
            subprocess.Popen(f'explorer "{folder}"')
        else:
            notify("Папка больше не существует.", "error")

    def refresh_history():
        history = load_history()
        history_list.controls.clear()
        if not history:
            history_list.controls.append(ft.Text("Журнал операций пуст.", color=ft.Colors.WHITE_54))
            return

        for item in history:
            out_p = item.get("output_path", "")
            history_list.controls.append(
                ft.Container(
                    bgcolor="#1E293B",
                    border_radius=10,
                    padding=15,
                    border=ft.border.all(1, "#334155"),
                    content=ft.Row(
                        alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                        controls=[
                            ft.Column(
                                spacing=4,
                                controls=[
                                    ft.Text(f"{item.get('timestamp')} — {item.get('posylka')}", weight=ft.FontWeight.BOLD, size=14, color=ft.Colors.WHITE),
                                    ft.Text(f"Совпадений: {item.get('matched')} из {item.get('total_rows')} | Предупреждений: {item.get('warnings_count')}", size=12, color=ft.Colors.GREEN_400),
                                    ft.Text(f"Сохранено: {out_p}", size=11, color=ft.Colors.WHITE_54)
                                ]
                            ),
                            ft.Button(
                                content=ft.Row([ft.Icon(ft.Icons.FOLDER, size=16), ft.Text("Открыть папку")]),
                                on_click=lambda _, p=out_p: open_folder(p),
                                style=ft.ButtonStyle(bgcolor="#334155", color=ft.Colors.WHITE)
                            )
                        ]
                    )
                )
            )

    view_history = ft.Column(
        expand=True,
        spacing=15,
        controls=[
            ft.Text("История операций", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Text("Журнал всех выполненных сведений и сформированных отчётов", size=14, color=ft.Colors.WHITE_54),
            history_list
        ]
    )

    txt_save_dir = ft.TextField(label="Папка по умолчанию для сохранения итоговых Excel", value=settings.get("save_dir", ""), expand=True)
    sw_autoload = ft.Switch(label="Автоматически подставлять последние файлы при запуске", value=settings.get("auto_load_files", True))
    sw_backup = ft.Switch(label="Создавать резервную копию (Backup) при перезаписи файла", value=settings.get("backup_results", True))

    def on_save_settings_click(e):
        settings["save_dir"] = txt_save_dir.value
        settings["auto_load_files"] = sw_autoload.value
        settings["backup_results"] = sw_backup.value
        save_settings(settings)
        notify("Настройки успешно сохранены!", "success")

    view_settings = ft.Column(
        spacing=20,
        controls=[
            ft.Text("Настройки приложения", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Container(
                bgcolor="#1E293B",
                border_radius=12,
                padding=20,
                border=ft.border.all(1, "#334155"),
                content=ft.Column(
                    spacing=15,
                    controls=[
                        ft.Row([
                            txt_save_dir,
                            ft.Button(content=ft.Text("Выбрать папку"), on_click=lambda _: picker_save_dir.get_directory_path(), height=48)
                        ]),
                        sw_autoload,
                        sw_backup,
                        ft.Button(
                            content=ft.Text("Сохранить настройки", weight=ft.FontWeight.BOLD),
                            on_click=on_save_settings_click,
                            style=ft.ButtonStyle(bgcolor="#2563EB", color=ft.Colors.WHITE),
                            height=45
                        )
                    ]
                )
            )
        ]
    )

    view_about = ft.Column(
        spacing=15,
        controls=[
            ft.Text("О программе", size=26, weight=ft.FontWeight.BOLD, color=ft.Colors.WHITE),
            ft.Container(
                bgcolor="#1E293B",
                border_radius=12,
                padding=25,
                border=ft.border.all(1, "#334155"),
                content=ft.Column(
                    spacing=10,
                    controls=[
                        ft.Text("PCB Auto Excel — Автоматическое сведение накладных", size=18, weight=ft.FontWeight.BOLD, color=ft.Colors.BLUE_400),
                        ft.Text("Версия: 2.0 Commercial Edition (Windows 10 / 11)", size=14, color=ft.Colors.WHITE_70),
                        ft.Divider(color="#334155"),
                        ft.Text(
                            "Программа предназначена для автоматического поиска и заполнения данных о количестве («По счету») "
                            "и способе доставки («Доставка») в накладных на основе сводных таблиц Головного офиса и Филиала.\n\n"
                            "Цветовая маркировка в итоговом Excel-файле:\n"
                            "• Светло-зелёный — данные заполнены из Сводки ПП (Голова).\n"
                            "• Светло-синий — данные заполнены из Сводки Филиала.\n"
                            "• Жёлтый — заказ найден в обеих сводках (приоритет отдан Голове).\n"
                            "• Светло-красный — заказ PS не найден ни в одной из сводок.",
                            size=14,
                            color=ft.Colors.WHITE_70
                        )
                    ]
                )
            )
        ]
    )

    refresh_dashboard()
    refresh_history()

    pages = [
        dashboard_cards,
        view_reconcile,
        view_preview,
        view_history,
        view_settings,
        view_about
    ]

    content_area = ft.Container(content=view_reconcile, expand=True, padding=30)

    def on_nav_change(e):
        idx = e.control.selected_index
        if idx == 0:
            refresh_dashboard()
        elif idx == 2:
            load_preview_data()
        elif idx == 3:
            refresh_history()
        content_area.content = pages[idx]
        page.update()

    rail = ft.NavigationRail(
        selected_index=1,
        label_type=ft.NavigationRailLabelType.ALL,
        min_width=110,
        min_extended_width=220,
        extended=True,
        bgcolor="#1E293B",
        on_change=on_nav_change,
        destinations=[
            ft.NavigationRailDestination(icon=ft.Icons.DASHBOARD_OUTLINED, selected_icon=ft.Icons.DASHBOARD, label="Главная"),
            ft.NavigationRailDestination(icon=ft.Icons.MERGE_TYPE_OUTLINED, selected_icon=ft.Icons.MERGE_TYPE, label="Сведение накладных"),
            ft.NavigationRailDestination(icon=ft.Icons.TABLE_VIEW_OUTLINED, selected_icon=ft.Icons.TABLE_VIEW, label="Предпросмотр"),
            ft.NavigationRailDestination(icon=ft.Icons.HISTORY_OUTLINED, selected_icon=ft.Icons.HISTORY, label="История операций"),
            ft.NavigationRailDestination(icon=ft.Icons.SETTINGS_OUTLINED, selected_icon=ft.Icons.SETTINGS, label="Настройки"),
            ft.NavigationRailDestination(icon=ft.Icons.INFO_OUTLINE, selected_icon=ft.Icons.INFO, label="О программе"),
        ]
    )

    page.add(
        ft.Row(
            expand=True,
            spacing=0,
            controls=[
                rail,
                ft.VerticalDivider(width=1, color="#334155"),
                content_area
            ]
        )
    )

ft.run(main)
