import customtkinter as ctk
from tkinter import filedialog, messagebox
from reconciler import process_waybills
import os

ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Автоматическое сведение накладных - pcb-auto-excel")
        self.geometry("1000x700")
        self.processed_df = None
        self.entries = {}
        self.setup_ui()

    def setup_ui(self):
        lbl_title = ctk.CTkLabel(self, text="Автоматическое сведение накладных", font=("Arial", 22, "bold"))
        lbl_title.pack(pady=(20, 5), padx=20, anchor="w")
        lbl_subtitle = ctk.CTkLabel(self, text="Выбери файлы через проводник или вставь прямой путь (включая сетевые папки \\\\Server\\...)", text_color="gray")
        lbl_subtitle.pack(pady=(0, 20), padx=20, anchor="w")

        self.create_file_input("1. Посылка (Обязательно)", "posylka")
        self.create_file_input("2. Сводка ПП (Голова)", "pp")
        self.create_file_input("3. Сводка Филиал (ЭЛЕКТРОконнект)", "filial")

        self.btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.btn_frame.pack(pady=20, padx=20, fill="x")
        
        self.btn_run = ctk.CTkButton(self.btn_frame, text="▶ СВЕСТИ НАКЛАДНУЮ", font=("Arial", 14, "bold"), height=40, command=self.run_process)
        self.btn_run.pack(side="left", fill="x", expand=True, padx=(0, 10))

        self.btn_download = ctk.CTkButton(self.btn_frame, text="📥 СКАЧАТЬ EXCEL (.XLSX)", font=("Arial", 14, "bold"), height=40, fg_color="#10b981", hover_color="#059669", state="disabled", command=self.save_excel)
        self.btn_download.pack(side="right", fill="x", expand=True, padx=(10, 0))

        self.result_frame = ctk.CTkFrame(self)
        self.result_frame.pack(fill="both", expand=True, padx=20, pady=(0, 20))
        
        header_frame = ctk.CTkFrame(self.result_frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(header_frame, text="Результат сведения:", font=("Arial", 14, "bold"), text_color="#10b981").pack(side="left")
        
        self.lbl_stats = ctk.CTkLabel(header_frame, text="Голова: 0  |  Филиал: 0", font=("Arial", 12, "bold"))
        self.lbl_stats.pack(side="right")
        
        self.txt_preview = ctk.CTkTextbox(self.result_frame, font=("Courier", 12))
        self.txt_preview.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def create_file_input(self, title, key):
        frame = ctk.CTkFrame(self)
        frame.pack(fill="x", padx=20, pady=5)
        
        lbl = ctk.CTkLabel(frame, text=title, text_color="#10b981", font=("Arial", 12, "bold"))
        lbl.pack(anchor="w", padx=10, pady=(10, 0))
        
        entry_frame = ctk.CTkFrame(frame, fg_color="transparent")
        entry_frame.pack(fill="x", padx=10, pady=(5, 10))
        
        entry = ctk.CTkEntry(entry_frame, placeholder_text="\\\\Fs01\\...")
        entry.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entries[key] = entry
        
        btn = ctk.CTkButton(entry_frame, text="Обзор", width=80, command=lambda: self.browse_file(entry))
        btn.pack(side="right")

    def browse_file(self, entry):
        filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if filepath:
            filepath = os.path.normpath(filepath)
            entry.delete(0, "end")
            entry.insert(0, filepath)

    def run_process(self):
        p_posylka = self.entries["posylka"].get().strip(' "')
        p_pp = self.entries["pp"].get().strip(' "')
        p_filial = self.entries["filial"].get().strip(' "')
        
        if not p_posylka or not p_pp or not p_filial:
            messagebox.showwarning("Внимание", "Укажите пути ко всем трем файлам.")
            return

        self.btn_run.configure(text="ОБРАБОТКА...", state="disabled")
        self.update()
        
        success, df, preview, c_head, c_filial = process_waybills(p_posylka, p_pp, p_filial)
        
        if success:
            self.processed_df = df
            self.lbl_stats.configure(text=f"Голова: {c_head}  |  Филиал: {c_filial}")
            self.txt_preview.delete("1.0", "end")
            self.txt_preview.insert("end", preview)
            self.btn_download.configure(state="normal")
            messagebox.showinfo("Готово", "Данные успешно сведены! Нажмите 'Скачать Excel'.")
        else:
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{preview}")
            
        self.btn_run.configure(text="▶ СВЕСТИ НАКЛАДНУЮ", state="normal")

    def save_excel(self):
        if self.processed_df is not None:
            save_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Посылка_СВЕДЕНО.xlsx",
                filetypes=[("Excel files", "*.xlsx")]
            )
            if save_path:
                try:
                    self.processed_df.to_excel(save_path, index=False)
                    messagebox.showinfo("Сохранено", f"Файл сохранен:\n{save_path}")
                except Exception as e:
                    messagebox.showerror("Ошибка сохранения", str(e))

if __name__ == "__main__":
    app = App()
    app.mainloop()
