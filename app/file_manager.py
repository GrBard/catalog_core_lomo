from tkinter import filedialog, ttk, Toplevel, messagebox
import customtkinter as ctk
from docx2pdf import convert
import os

class FileManager:
    def __init__(self):
        self.excel_path = None
        self.images_folder = None
        self.last_catalog_path = None
        self.main_file_columns = None  # Для хранения выбранных столбцов основного файла
        self.samples_file_columns = None  # Для хранения выбранных столбцов файла с образцами

    def select_excel(self):
        """Выбирает Excel-файл и запрашивает выбор столбцов."""
        self.excel_path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if self.excel_path:
            # Загружаем файл, чтобы получить список столбцов
            import pandas as pd
            try:
                df = pd.read_excel(self.excel_path)
                columns = list(df.columns)
                if not columns:
                    messagebox.showerror("Ошибка", "Excel-файл пуст или не содержит столбцов.")
                    self.excel_path = None
                    return None
                # Запрашиваем выбор столбцов
                self.main_file_columns = self.select_columns(
                    columns,
                    ["BOX (номер коробки)", "От (начало интервала)", "До (конец интервала)"],
                    "Выбор столбцов для основного файла"
                )
                if not self.main_file_columns:
                    self.excel_path = None
                    return None
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить Excel-файл: {str(e)}")
                self.excel_path = None
                return None
        return self.excel_path

    def select_folder(self):
        """Выбирает папку с изображениями."""
        self.images_folder = filedialog.askdirectory(
            title="Выберите папку с изображениями"
        )
        return self.images_folder

    def select_samples_file(self):
        """Выбирает файл с образцами и запрашивает выбор столбцов."""
        samples_file = filedialog.askopenfilename(
            title="Выберите файл с образцами",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if samples_file:
            # Загружаем файл, чтобы получить список столбцов
            import pandas as pd
            try:
                df = pd.read_excel(samples_file)
                columns = list(df.columns)
                if len(columns) < 3:
                    messagebox.showerror("Ошибка", "Файл с образцами должен содержать как минимум 3 столбца.")
                    return None
                # Запрашиваем выбор столбцов
                self.samples_file_columns = self.select_columns(
                    columns,
                    ["Номер образца", "Глубина"],
                    "Выбор столбцов для файла с образцами"
                )
                if not self.samples_file_columns:
                    return None
            except Exception as e:
                messagebox.showerror("Ошибка", f"Не удалось загрузить файл с образцами: {str(e)}")
                return None
        return samples_file

    def select_columns(self, columns, labels, title):
        """Создаёт окно для выбора столбцов из списка."""
        window = ctk.CTkToplevel()
        window.title(title)
        window.geometry("400x300")
        window.resizable(False, False)
        window.transient()  # Привязываем к главному окну
        window.grab_set()  # Блокируем взаимодействие с главным окном

        # Метка
        label = ctk.CTkLabel(window, text="Выберите соответствующие столбцы:", font=("Helvetica", 12))
        label.pack(pady=10)

        # Создаём выпадающие списки для каждого требуемого столбца
        selected_columns = {}
        for label_text in labels:
            frame = ctk.CTkFrame(window)
            frame.pack(fill="x", padx=20, pady=5)
            lbl = ctk.CTkLabel(frame, text=label_text, width=150, anchor="w")
            lbl.pack(side="left")
            combo = ctk.CTkComboBox(frame, values=columns, width=200)
            combo.pack(side="right")
            selected_columns[label_text] = combo

        # Кнопка подтверждения
        result = [None]  # Для хранения результата
        def confirm():
            selected = []
            for label_text in labels:
                value = selected_columns[label_text].get()
                if not value:
                    messagebox.showerror("Ошибка", f"Выберите столбец для '{label_text}'.")
                    return
                selected.append(value)
            # Проверяем, что все выбранные столбцы уникальны
            if len(set(selected)) != len(selected):
                messagebox.showerror("Ошибка", "Выбранные столбцы должны быть уникальными.")
                return
            result[0] = selected
            window.destroy()

        btn_confirm = ctk.CTkButton(window, text="Подтвердить", command=confirm, corner_radius=8)
        btn_confirm.pack(pady=20)

        # Кнопка отмены
        def cancel():
            result[0] = None
            window.destroy()

        btn_cancel = ctk.CTkButton(window, text="Отмена", command=cancel, corner_radius=8, fg_color="#FF5555")
        btn_cancel.pack(pady=5)

        window.protocol("WM_DELETE_WINDOW", cancel)
        window.wait_window()  # Ждём, пока окно не закроется
        return result[0]

    def save_dataframe(self, dataframe):
        """Сохраняет DataFrame в Excel-файл."""
        if dataframe is None:
            raise ValueError("Нет данных для сохранения.")
        result_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            title="Сохранить таблицу"
        )
        if not result_path:
            return None
        dataframe.to_excel(result_path, index=False, engine="openpyxl")
        return result_path

    def save_catalog(self):
        """Запрашивает путь для сохранения каталога в формате .docx."""
        save_path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word files", "*.docx")],
            title="Сохранить каталог"
        )
        if save_path:
            self.last_catalog_path = save_path
        return save_path

    def convert_to_pdf(self):
        """Конвертирует последний созданный каталог в PDF."""
        if self.last_catalog_path is None:
            raise ValueError("Сначала создайте каталог в формате Word.")
        pdf_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="Сохранить как PDF",
            initialfile=os.path.splitext(os.path.basename(self.last_catalog_path))[0]
        )
        if not pdf_path:
            return None
        convert(self.last_catalog_path, pdf_path)
        return pdf_path

    def get_excel_path(self):
        return self.excel_path

    def get_images_folder(self):
        return self.images_folder

    def get_last_catalog_path(self):
        return self.last_catalog_path

    def get_main_file_columns(self):
        return self.main_file_columns

    def get_samples_file_columns(self):
        return self.samples_file_columns