from tkinter import filedialog
import os
from docx2pdf import convert

class FileManager:
    def __init__(self):
        self.excel_path = None
        self.images_folder = None
        self.last_catalog_path = None

    def select_excel(self):
        """Выбирает Excel-файл."""
        self.excel_path = filedialog.askopenfilename(
            title="Выберите Excel-файл",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        return self.excel_path

    def select_folder(self):
        """Выбирает папку с изображениями."""
        self.images_folder = filedialog.askdirectory(
            title="Выберите папку с изображениями"
        )
        return self.images_folder

    def select_samples_file(self):
        """Выбирает файл с образцами."""
        samples_file = filedialog.askopenfilename(
            title="Выберите файл с образцами",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        return samples_file

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