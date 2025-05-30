import customtkinter as ctk
from customtkinter import CTkFrame, CTkButton, CTkCheckBox, CTkLabel, CTkScrollableFrame, CTkImage, CTkTabview
from tkinter import messagebox, ttk
from PIL import Image
import os
import platform
import subprocess
import re
import pandas as pd
from app.utils import find_continuous_intervals, resource_path
from app.data_processor import DataProcessor

class AppUI:
    def __init__(self, root, file_manager):
        """Инициализирует пользовательский интерфейс приложения."""
        self.root = root
        self.root.title("Каталог образцов и фотографий")
        self.root.geometry("1000x700")
        self.root.minsize(800, 600)
        self.file_manager = file_manager
        self.data_processor = None
        self.samples_file = None
        self.samples_dataframe = None  # DataFrame для образцов

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Основной фрейм
        self.main_frame = CTkFrame(self.root, corner_radius=10)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Верхний фрейм с кнопками
        self.top_frame = CTkFrame(self.main_frame, corner_radius=10)
        self.top_frame.pack(fill="x", pady=(0, 10))

        # Кнопки в верхнем фрейме
        self.btn_select_excel = CTkButton(self.top_frame, text="Загрузить Excel", command=self.select_excel,
                                          corner_radius=8, font=("Helvetica", 12))
        self.btn_select_excel.grid(row=0, column=0, padx=10, pady=10)
        self.btn_select_folder = CTkButton(self.top_frame, text="Папка с фото", command=self.select_folder,
                                           corner_radius=8, font=("Helvetica", 12))
        self.btn_select_folder.grid(row=0, column=1, padx=10, pady=10)
        self.btn_process_data = CTkButton(self.top_frame, text="Обработать", command=self.process_data,
                                          corner_radius=8, font=("Helvetica", 12))
        self.btn_process_data.grid(row=0, column=2, padx=10, pady=10)
        self.samples_var = ctk.BooleanVar()
        self.chk_samples = CTkCheckBox(self.top_frame, text="Образцы", variable=self.samples_var,
                                       command=self.toggle_samples_button, font=("Helvetica", 12))
        self.chk_samples.grid(row=0, column=3, padx=10, pady=10)
        self.btn_select_samples = CTkButton(self.top_frame, text="Файл с образцами...",
                                            command=self.select_samples_file,
                                            state="disabled", corner_radius=8, font=("Helvetica", 12))
        self.btn_select_samples.grid(row=0, column=4, padx=10, pady=10)
        self.btn_create_catalog = CTkButton(self.top_frame, text="Создать каталог (Word)", command=self.create_catalog,
                                            corner_radius=8, font=("Helvetica", 12))
        self.btn_create_catalog.grid(row=1, column=0, padx=10, pady=10)
        #self.btn_preview = CTkButton(self.top_frame, text="Предпросмотр", command=self.preview_catalog,
        #                             corner_radius=8, font=("Helvetica", 12))
        #self.btn_preview.grid(row=1, column=4, padx=10, pady=10)
        self.btn_convert_pdf = CTkButton(self.top_frame, text="Конвертировать в PDF", command=self.convert_to_pdf,
                                         corner_radius=8, font=("Helvetica", 12))
        self.btn_convert_pdf.grid(row=1, column=1, padx=10, pady=10)
        self.btn_save_data = CTkButton(self.top_frame, text="Сохранить таблицу", command=self.save_data,
                                       corner_radius=8, font=("Helvetica", 12))
        self.btn_save_data.grid(row=1, column=2, padx=10, pady=10)

        # Кнопка "Очистить данные"
        self.clear_button = ctk.CTkButton(self.top_frame, text="Очистить данные", command=self.clear_data,
                                          corner_radius=8, font=("Helvetica", 12))
        self.clear_button.grid(row=1, column=3, pady=10, padx=10)

        # Фрейм для вкладок
        self.tab_view = CTkTabview(self.main_frame, corner_radius=10)
        self.tab_view.pack(fill="both", expand=True)

        # Вкладка для основной таблицы
        self.tab_main = self.tab_view.add("Основные данные")
        self.table_frame = CTkFrame(self.tab_main, corner_radius=10)
        self.table_frame.pack(fill="both", expand=True)

        # Вкладка для образцов (изначально скрыта)
        self.tab_samples = None  # Изначально вкладка не создана
        self.samples_table_frame = None  # Фрейм для таблицы образцов тоже пока не создаём

        # Статусная строка
        self.status_var = ctk.StringVar(value="Готово")
        self.status_bar = CTkLabel(self.main_frame, textvariable=self.status_var, font=("Helvetica", 10),
                                   corner_radius=5, fg_color="#2A2A2A")
        self.status_bar.pack(side="bottom", fill="x")

        # Сохраняем пути к данным для повторной обработки
        self.last_excel_path = None
        self.last_images_folder = None
        self.last_samples_path = None

    def clear_data(self):
        response = messagebox.askyesno("Подтверждение", "Вы уверены, что хотите очистить все данные?")
        if response:
            self.data_processor = None
            self.last_excel_path = None
            self.last_images_folder = None
            self.last_samples_path = None
            self.samples_dataframe = None
            self.status_var.set("Данные очищены.")
            self.samples_var.set(False)
            print("Все данные очищены")
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            if self.tab_samples and "Образцы" in self.tab_view._tab_dict:
                self.tab_view.delete("Образцы")
                self.tab_samples = None
                self.samples_table_frame = None

    def toggle_samples_button(self):
        """Переключает состояние кнопки выбора файла с образцами."""
        state = "normal" if self.samples_var.get() else "disabled"
        self.btn_select_samples.configure(state=state)

    def select_excel(self):
        """Выбирает Excel-файл и сохраняет путь."""
        excel_path = self.file_manager.select_excel()
        if excel_path:
            self.last_excel_path = excel_path
            self.status_var.set(f"Выбран Excel: {os.path.basename(excel_path)}")
        else:
            self.status_var.set("Выбор Excel-файла отменён")

    def select_folder(self):
        """Выбирает папку с фотографиями и сохраняет путь."""
        images_folder = self.file_manager.select_folder()
        if images_folder:
            self.last_images_folder = images_folder
            self.status_var.set(f"Выбрана папка: {os.path.basename(images_folder)}")
        else:
            self.status_var.set("Выбор папки отменён")

    def select_samples_file(self):
        """Выбирает файл с образцами и сохраняет путь."""
        samples_path = self.file_manager.select_samples_file()
        if samples_path:
            self.last_samples_path = samples_path
            self.samples_dataframe = pd.read_excel(samples_path)
            self.status_var.set(f"Выбран файл с образцами: {os.path.basename(samples_path)}")
        else:
            self.status_var.set("Выбор файла с образцами отменён")
            self.last_samples_path = None
            self.samples_dataframe = None

    def process_data(self):
        """Пересчитывает данные на основе уже выбранных путей."""
        # Проверяем, что основные пути выбраны
        if not self.last_excel_path or not self.last_images_folder:
            messagebox.showerror("Ошибка", "Сначала выберите Excel-файл и папку с фотографиями.")
            return

        # Если данные уже обработаны, спрашиваем о пересчёте
        if self.data_processor and self.data_processor.get_current_dataframe() is not None:
            response = messagebox.askyesno("Пересчёт данных",
                                           "Данные уже обработаны. Хотите пересчитать их с новыми параметрами?")
            if not response:
                self.status_var.set("Действие отменено. Текущие данные сохранены.")
                print("Пересчёт отменён")
                return

        try:
            self.status_var.set("Обработка данных...")
            print(
                f"Excel: {self.last_excel_path}, Images: {self.last_images_folder}, Samples: {self.last_samples_path}")

            # Получаем выбранные столбцы из FileManager
            main_columns = self.file_manager.get_main_file_columns()
            print(f"Выбранные столбцы: {main_columns}")
            if main_columns and len(main_columns) == 4:
                box_column = main_columns[0]
                start_column = main_columns[1]
                end_column = main_columns[2]
                measurements_column = main_columns[3]
                print(
                    f"Используемые столбцы: box={box_column}, start={start_column}, end={end_column}, measurements={measurements_column}")
            else:
                box_column = "BOX"
                start_column = "от"
                end_column = "до"
                measurements_column = "замеры"
                print("Используются столбцы по умолчанию")

            # Пересоздаём DataProcessor с текущими путями
            self.data_processor = DataProcessor(
                self.last_excel_path,
                self.last_images_folder,
                box_column=box_column,
                start_column=start_column,
                end_column=end_column,
                measurements_column=measurements_column
            )

            # Пересчитываем основные данные
            df = self.data_processor.process_data()

            # Отображаем основной DataFrame
            print(f"Столбцы в DataFrame: {list(df.columns)}")
            self.display_dataframe(df)

            # Если выбраны образцы, обрабатываем их
            if self.samples_var.get() and self.last_samples_path:
                self.samples_file = self.last_samples_path  # Устанавливаем путь к файлу образцов
                self.process_samples()  # Обрабатываем и отображаем данные образцов

            self.status_var.set(f"Данные обработаны: {len(df)} строк")

        except Exception as e:
            self.status_var.set(f"Ошибка обработки: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось обработать данные: {str(e)}")
            raise

    def process_samples(self):
        """Обрабатывает данные образцов и создаёт DataFrame для второй вкладки."""
        samples_df = pd.read_excel(self.samples_file, sheet_name=0)

        if samples_df.shape[1] < 3:
            raise ValueError("Файл образцов должен содержать как минимум 3 столбца.")

        sample_numbers = samples_df.iloc[:, 1]  # 2-й столбец: номер образца
        absolute_depths = samples_df.iloc[:, 2]  # 3-й столбец: абсолютная глубина

        temp_data = []
        # Получаем имя столбца для коробок из DataProcessor
        box_column = self.data_processor.box_column
        for idx, (sample_num, abs_depth) in enumerate(zip(sample_numbers, absolute_depths)):
            try:
                sample_num = float(sample_num)
                box_num = int(sample_num)  # Целая часть — номер коробки
                research = []
                for col in samples_df.columns[3:]:
                    if samples_df.iloc[idx][col] == "+":
                        research.append(str(col))
                research_str = ", ".join(research) if research else "Нет исследований"
                temp_data.append({
                    box_column: box_num,  # Используем динамическое имя столбца
                    "Номер образца": sample_num,
                    "Глубина": round(float(abs_depth), 2) if pd.notna(abs_depth) else abs_depth,
                    "Исследования": research_str
                })
            except (ValueError, TypeError):
                continue

        temp_df = pd.DataFrame(temp_data)
        if not temp_df.empty:
            grouped = temp_df.groupby("Номер образца").agg({
                box_column: "first",  # Используем динамическое имя столбца
                "Глубина": "first",
                "Исследования": lambda x: ", ".join(sorted(set(
                    item for sublist in x for item in sublist.split(", ") if item != "Нет исследований"
                )))
            }).reset_index()
            grouped["Исследования"] = grouped["Исследования"].apply(
                lambda x: x if x else "Нет исследований"
            )
            self.samples_dataframe = grouped[[box_column, "Номер образца", "Глубина", "Исследования"]]
        else:
            self.samples_dataframe = pd.DataFrame(columns=[box_column, "Номер образца", "Глубина", "Исследования"])

        self.samples_dataframe = self.samples_dataframe.sort_values(by="Номер образца")

        if self.samples_dataframe.empty:
            messagebox.showwarning("Предупреждение", "Нет данных об образцах для отображения.")
            return

        if "Образцы" not in self.tab_view._tab_dict:
            self.tab_samples = self.tab_view.add("Образцы")
            self.samples_table_frame = CTkFrame(self.tab_samples, corner_radius=10)
            self.samples_table_frame.pack(fill="both", expand=True)

        self.display_samples_dataframe(self.samples_dataframe)

    def check_samples_issues(self):
        """Проверяет образцы на отсутствие '+' и одинаковые номера образцов."""
        if self.samples_dataframe is None:
            return []

        issues = []
        samples_df = pd.read_excel(self.samples_file, sheet_name=0)

        # Проверяем отсутствие "+"
        for idx, row in samples_df.iterrows():
            has_plus = False
            for col in samples_df.columns[3:]:  # Начинаем с 4-го столбца
                if row[col] == "+":
                    has_plus = True
                    break
            if not has_plus:
                sample_num = row.iloc[1]  # 2-й столбец
                issues.append(f"Образец {sample_num}: Данные об исследованиях не найдены.")

        # Проверяем одинаковые номера образцов (2-й столбец)
        sample_numbers = samples_df.iloc[:, 1]
        duplicates = sample_numbers[sample_numbers.duplicated(keep=False)]
        if not duplicates.empty:
            for sample_num in duplicates.unique():
                issues.append(f"Номера образцов совпадают: {sample_num}.")

        return issues

    def display_dataframe(self, dataframe):
        """Отображает DataFrame в виде таблицы с выделением строк, где вынос > 100%."""
        for widget in self.table_frame.winfo_children():
            widget.destroy()

        if dataframe.empty:
            messagebox.showwarning("Предупреждение", "Нет данных для отображения.")
            return

        columns = list(dataframe.columns)
        self.tree = ttk.Treeview(self.table_frame, columns=columns, show="headings", height=len(dataframe))
        self.tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_x = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        scrollbar_y = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.tree.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        self.table_frame.grid_rowconfigure(0, weight=1)
        self.table_frame.grid_columnconfigure(0, weight=1)

        from tkinter.font import Font
        font = Font(family="Helvetica", size=12)

        # Настраиваем заголовки и ширину столбцов
        for col in columns:
            self.tree.heading(col, text=col)
            max_length = len(str(col))
            col_data = dataframe[col].astype(str).apply(len)
            if not col_data.empty:
                max_length = max(max_length, col_data.max())
            width = max(50, min(300, max_length * 10))
            self.tree.column(col, width=width, anchor="w", stretch=False)

        # Определяем индекс столбца "Вынос"
        vynos_col_index = columns.index("Вынос") if "Вынос" in columns else -1

        # Вставляем строки и выделяем те, где вынос > 100%
        for index, row in dataframe.iterrows():
            row_values = list(row)
            row_id = self.tree.insert("", "end", iid=str(index), values=row_values)

            # Проверяем значение в столбце "Вынос"
            if vynos_col_index != -1:
                vynos_value = row["Вынос"]
                if isinstance(vynos_value, str) and vynos_value != "N/A":
                    # Извлекаем процент из строки вида "X м (Y %)"
                    match = re.search(r"\((\d+\.\d+|\d+) %\)", vynos_value)
                    if match:
                        percentage = float(match.group(1))
                        if percentage > 100:
                            # Выделяем строку красным цветом
                            self.tree.item(row_id, tags=("highlight",))

        # Настраиваем тег для выделения
        self.tree.tag_configure("highlight", background="#FF6666")  # Красный фон

        self.tree.bind("<Double-1>", self.on_double_click)

    def display_samples_dataframe(self, dataframe):
        """Отображает DataFrame образцов во второй вкладке."""
        for widget in self.samples_table_frame.winfo_children():
            widget.destroy()

        if dataframe.empty:
            messagebox.showwarning("Предупреждение", "Нет данных об образцах для отображения.")
            return

        columns = list(dataframe.columns)
        self.samples_tree = ttk.Treeview(self.samples_table_frame, columns=columns, show="headings",
                                         height=len(dataframe))
        self.samples_tree.grid(row=0, column=0, sticky="nsew")

        scrollbar_x = ttk.Scrollbar(self.samples_table_frame, orient="horizontal", command=self.samples_tree.xview)
        scrollbar_x.grid(row=1, column=0, sticky="ew")
        scrollbar_y = ttk.Scrollbar(self.samples_table_frame, orient="vertical", command=self.samples_tree.yview)
        scrollbar_y.grid(row=0, column=1, sticky="ns")
        self.samples_tree.configure(xscrollcommand=scrollbar_x.set, yscrollcommand=scrollbar_y.set)

        self.samples_table_frame.grid_rowconfigure(0, weight=1)
        self.samples_table_frame.grid_columnconfigure(0, weight=1)

        from tkinter.font import Font
        font = Font(family="Helvetica", size=12)

        # Настраиваем заголовки и ширину столбцов
        for col in columns:
            self.samples_tree.heading(col, text=col)
            max_length = len(str(col))
            col_data = dataframe[col].astype(str).apply(len)
            if not col_data.empty:
                max_length = max(max_length, col_data.max())
            width = max(50, min(300, max_length * 10))
            self.samples_tree.column(col, width=width, anchor="w", stretch=False)

        # Вставляем строки и выделяем те, где "Нет исследований"
        for index, row in dataframe.iterrows():
            row_values = list(row)
            row_id = self.samples_tree.insert("", "end", iid=str(index), values=row_values)
            if row["Исследования"] == "Нет исследований":
                self.samples_tree.item(row_id, tags=("no_research",))

        # Настраиваем тег для выделения
        self.samples_tree.tag_configure("no_research", background="#FF6666")  # Красный фон

        # Привязываем обработчик двойного клика
        self.samples_tree.bind("<Double-1>", self.on_double_click_samples)

    def on_double_click_samples(self, event):
        """Обрабатывает двойной клик по ячейке таблицы образцов для редактирования."""
        region = self.samples_tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.samples_tree.identify_row(event.y)
        column = self.samples_tree.identify_column(event.x)
        x, y, width, height = self.samples_tree.bbox(row_id, column)
        cell_value = self.samples_tree.item(row_id, "values")[int(column.replace('#', '')) - 1]

        # Создаём CTkEntry для редактирования ячейки
        entry = ctk.CTkEntry(self.samples_tree, width=width, height=height, font=("Helvetica", 12))
        entry.place(x=x, y=y)
        entry.insert(0, cell_value)
        entry.focus_set()

        def on_focus_out(event):
            new_value = entry.get()
            self.samples_tree.set(row_id, column, new_value)
            # Обновляем DataFrame образцов
            col_index = int(column.replace('#', '')) - 1
            self.samples_dataframe.iat[int(row_id), col_index] = new_value
            entry.destroy()

            # Перепроверяем строку на наличие исследований
            if col_index == self.samples_dataframe.columns.get_loc("Исследования"):
                if new_value == "Нет исследований":
                    self.samples_tree.item(row_id, tags=("no_research",))
                else:
                    self.samples_tree.item(row_id, tags=())

        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", on_focus_out)

    def on_double_click(self, event):
        """Обрабатывает двойной клик по ячейке для редактирования."""
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.tree.identify_row(event.y)
        column = self.tree.identify_column(event.x)
        x, y, width, height = self.tree.bbox(row_id, column)
        cell_value = self.tree.item(row_id, "values")[int(column.replace('#', '')) - 1]

        # Создаём CTkEntry для редактирования ячейки
        entry = ctk.CTkEntry(self.tree, width=width, height=height, font=("Helvetica", 12))
        entry.place(x=x, y=y)
        entry.insert(0, cell_value)
        entry.focus_set()

        def on_focus_out(event):
            new_value = entry.get()
            self.tree.set(row_id, column, new_value)
            # Обновляем DataFrame в data_processor
            col_index = int(column.replace('#', '')) - 1
            df = self.data_processor.get_current_dataframe()
            df.iat[int(row_id), col_index] = new_value
            entry.destroy()

            # Перепроверяем строку на условие выноса > 100%
            if col_index == df.columns.get_loc("Вынос"):
                vynos_value = new_value
                if isinstance(vynos_value, str) and vynos_value != "N/A":
                    match = re.search(r"\((\d+\.\d+|\d+) %\)", vynos_value)
                    if match:
                        percentage = float(match.group(1))
                        if percentage > 100:
                            self.tree.item(row_id, tags=("highlight",))
                        else:
                            self.tree.item(row_id, tags=())

        entry.bind("<FocusOut>", on_focus_out)
        entry.bind("<Return>", on_focus_out)

    def save_data(self):
        """Сохраняет активный DataFrame в файл в зависимости от текущей вкладки."""
        # Определяем текущую вкладку
        current_tab = self.tab_view.get()

        if current_tab == "Основные данные":
            if not self.data_processor or self.data_processor.get_current_dataframe() is None:
                messagebox.showerror("Ошибка", "Нет основных данных для сохранения.")
                return
            df = self.data_processor.get_current_dataframe()
        elif current_tab == "Образцы":
            if self.samples_dataframe is None or self.samples_dataframe.empty:
                messagebox.showerror("Ошибка", "Нет данных об образцах для сохранения.")
                return
            df = self.samples_dataframe
        else:
            messagebox.showerror("Ошибка", "Неизвестная вкладка.")
            return

        try:
            result_path = self.file_manager.save_dataframe(df)
            if result_path:
                self.status_var.set(f"Таблица сохранена: {result_path}")
                # Показываем уведомление с вопросом об открытии
                response = messagebox.askyesno("Успех", f"Таблица сохранена: {result_path}\nХотите открыть её?")
                if response:  # Если пользователь согласен
                    self.open_file(result_path)
        except Exception as e:
            messagebox.showerror("Ошибка сохранения", str(e))

    def create_catalog(self):
        """Создаёт каталог на основе обработанных данных с отображением прогресса."""
        print("Начало метода create_catalog в AppUI")
        if not self.data_processor or self.data_processor.get_current_dataframe() is None:
            print("Ошибка: нет данных для создания каталога")
            messagebox.showerror("Ошибка", "Нет данных для создания каталога.")
            return

        try:
            save_path = self.file_manager.save_catalog()
            print(f"Выбран путь сохранения: {save_path}")
            if save_path:
                print(f"Текущий box_column в data_processor: '{self.data_processor.box_column}'")
                print(f"Столбцы в current_dataframe: {list(self.data_processor.get_current_dataframe().columns)}")

                # Создаём окно прогресса
                progress_window = ctk.CTkToplevel(self.root)
                progress_window.title("Создание каталога")
                progress_window.geometry("300x150")
                progress_window.resizable(False, False)
                progress_window.transient(self.root)
                progress_window.grab_set()

                label = ctk.CTkLabel(progress_window, text="Создание каталога...", font=("Helvetica", 12))
                label.pack(pady=10)

                progress_bar = ctk.CTkProgressBar(progress_window, width=250)
                progress_bar.pack(pady=10)
                progress_bar.set(0)

                total_boxes = len(self.data_processor.get_current_dataframe().groupby(self.data_processor.box_column))
                progress_step = 1.0 / total_boxes if total_boxes > 0 else 1.0
                print(f"Количество коробок: {total_boxes}, шаг прогресса: {progress_step}")

                print("Вызов DataProcessor.create_catalog")
                self.data_processor.create_catalog(
                    save_path,
                    self.samples_dataframe if self.samples_var.get() else None,
                    progress_bar,
                    progress_step
                )

                progress_window.destroy()

                self.status_var.set(f"Каталог создан: {save_path}")
                response = messagebox.askyesno("Успех", f"Каталог создан: {save_path}\nХотите открыть его?")
                if response:
                    self.open_file(save_path)
        except Exception as e:
            if 'progress_window' in locals():
                progress_window.destroy()
            print(f"Ошибка в create_catalog: {str(e)}")
            messagebox.showerror("Ошибка создания каталога", str(e))
            raise  # Добавляем raise для полной трассировки

    def open_file(self, file_path):
        """Открывает файл в зависимости от операционной системы."""
        try:
            if platform.system() == "Windows":
                os.startfile(file_path)  # Для Windows
            elif platform.system() == "Darwin":  # macOS
                subprocess.run(["open", file_path])
            else:  # Linux и другие
                subprocess.run(["xdg-open", file_path])
        except Exception as e:
            print(f"Ошибка открытия файла: {str(e)}")
            messagebox.showerror("Ошибка", f"Не удалось открыть файл: {str(e)}")

    #def preview_catalog(self):
    #    """Показывает предпросмотр каталога, аналогичный Word-документу."""
    #    if not self.data_processor or self.data_processor.get_current_dataframe() is None:
    #        messagebox.showerror("Ошибка", "Нет данных для предпросмотра.")
    #        return
#
    #    preview_window = ctk.CTkToplevel(self.root)
    #    preview_window.title("Предпросмотр каталога")
    #    preview_window.geometry("1000x700")
    #    preview_window.attributes("-topmost", True)
#
    #    preview_frame = CTkScrollableFrame(preview_window, corner_radius=10)
    #    preview_frame.pack(fill="both", expand=True, padx=20, pady=20)
#
    #    title_label = CTkLabel(preview_frame,
    #                           text=f"Фотографии керна по скважине {self.data_processor.get_current_dataframe()['Скважина'].iloc[0]}",
    #                           font=("Helvetica", 18, "bold"))
    #    title_label.pack(pady=10)
#
    #    intro_label = CTkLabel(preview_frame,
    #                           text="Глубины даны по керну. Номера образцов расположены напротив точек выбуривания. "
    #                                "Номер образца состоит из двух цифр, разделённых точкой. "
    #                                "Первая цифра соответствует номеру коробки, вторая – расстояние в сантиметрах от низа коробки до точки отбора образца.",
    #                           font=("Helvetica", 12), wraplength=900, justify="left")
    #    intro_label.pack(pady=10)
#
    #    df = self.data_processor.get_current_dataframe()
    #    grouped = df.groupby(self.data_processor.box_column)
#
    #    scale_image_path = resource_path('resources/scale.jpg')
    #    shkala_image_path = resource_path('resources/shkala.jpg')
#
    #    target_height = 400
    #    shkala_height = 50
#
    #    for box_number, group in grouped:
    #        separator = CTkFrame(preview_frame, height=2, fg_color="#555555")
    #        separator.pack(fill="x", pady=20)
#
    #        table_frame = CTkFrame(preview_frame, corner_radius=5, fg_color="#333333")
    #        table_frame.pack(fill="x", padx=10, pady=10)
#
    #        width_samles_col = 100
    #        width_photo_col = 400
    #        width_samles_right = 300
    #        width_scale_col = 50
#
    #        cell_0_0 = CTkFrame(table_frame, width=width_samles_col, height=50, fg_color="#444444")
    #        cell_0_0.grid(row=0, column=0, padx=1, pady=1, sticky="nsew")
    #        label_0_0 = CTkLabel(cell_0_0, text=f"Коробка {int(box_number)}", font=("Helvetica", 12))
    #        label_0_0.pack(pady=5)
#
    #        cell_0_1 = CTkFrame(table_frame, width=width_photo_col, height=50, fg_color="#444444")
    #        cell_0_1.grid(row=0, column=1, padx=1, pady=1, sticky="nsew")
    #        ts = 'Интервал бурения: '
    #        for idx, row in group.iterrows():
    #            ts += f"{row['Начало интервала']}-{row['Конец интервала']}\nвынос: {row['Вынос']}\n"
    #        label_0_1 = CTkLabel(cell_0_1, text=ts.strip(), font=("Helvetica", 12), justify="left")
    #        label_0_1.pack(pady=5)
#
    #        cell_0_2 = CTkFrame(table_frame, width=width_samles_right, height=50, fg_color="#444444")
    #        cell_0_2.grid(row=0, column=2, padx=1, pady=1, sticky="nsew")
#
    #        cell_0_3 = CTkFrame(table_frame, width=width_scale_col, height=50, fg_color="#444444")
    #        cell_0_3.grid(row=0, column=3, padx=1, pady=1, sticky="nsew")
#
    #        cell_1_0 = CTkFrame(table_frame, width=width_samles_col, height=30, fg_color="#444444")
    #        cell_1_0.grid(row=1, column=0, padx=1, pady=1, sticky="nsew")
    #        label_1_0 = CTkLabel(cell_1_0, text="Номера образцов:", font=("Helvetica", 12))
    #        label_1_0.pack(pady=5)
#
    #        cell_1_1 = CTkFrame(table_frame, width=width_photo_col, height=30, fg_color="#444444")
    #        cell_1_1.grid(row=1, column=1, padx=1, pady=1, sticky="nsew")
    #        start_value = group[self.data_processor.start_column].iloc[
    #            0] if self.data_processor.start_column in group.columns else "N/A"
    #        label_1_1 = CTkLabel(cell_1_1, text=f"[{start_value}]", font=("Helvetica", 12))
    #        label_1_1.pack(pady=5)
#
    #        cell_1_2 = CTkFrame(table_frame, width=width_samles_right, height=30, fg_color="#444444")
    #        cell_1_2.grid(row=1, column=2, padx=1, pady=1, sticky="nsew")
    #        label_1_2 = CTkLabel(cell_1_2, text="Исследования:", font=("Helvetica", 12))
    #        label_1_2.pack(pady=5)
#
    #        cell_1_3 = CTkFrame(table_frame, width=width_scale_col, height=30, fg_color="#444444")
    #        cell_1_3.grid(row=1, column=3, padx=1, pady=1, sticky="nsew")
#
    #        cell_2_1 = CTkFrame(table_frame, width=width_photo_col, height=450, fg_color="#444444")
    #        cell_2_1.grid(row=2, column=1, padx=1, pady=1, sticky="nsew")
    #        images_frame = CTkFrame(cell_2_1, fg_color="#444444")
    #        images_frame.pack(pady=5)
#
    #        try:
    #            top_depth = float(group[self.data_processor.start_column].iloc[
    #                                  0]) if self.data_processor.start_column in group.columns else 0
    #            bottom_depth = float(group[self.data_processor.end_column].iloc[
    #                                     0]) if self.data_processor.end_column in group.columns else 0
    #            core_count = 1
    #            img_d, img_d2 = self.data_processor.generate_depth_scale(top_depth, bottom_depth, core_count)
    #            img = Image.open(img_d)
    #            img.thumbnail((50, target_height))
    #            depth_scale = CTkImage(light_image=img, dark_image=img, size=(50, target_height))
    #            depth_scale_label = CTkLabel(images_frame, image=depth_scale, text="")
    #            depth_scale_label.pack(side="left", padx=5)
    #            if img_d2:
    #                img2 = Image.open(img_d2)
    #                img2.thumbnail((50, target_height))
    #                depth_scale2 = CTkImage(light_image=img2, dark_image=img2, size=(50, target_height))
    #                depth_scale_label2 = CTkLabel(images_frame, image=depth_scale2, text="")
    #                depth_scale_label2.pack(side="left", padx=5)
    #        except Exception as e:
    #            error_label = CTkLabel(images_frame, text=f"Ошибка шкалы: {str(e)}", font=("Helvetica", 10))
    #            error_label.pack(side="left", padx=5)
#
    #        # ... (далее без изменений до конца метода)
#
    #        cell_3_1 = CTkFrame(table_frame, width=width_photo_col, height=30, fg_color="#444444")
    #        cell_3_1.grid(row=3, column=1, padx=1, pady=1, sticky="nsew")
    #        end_value = group[self.data_processor.end_column].iloc[
    #            0] if self.data_processor.end_column in group.columns else "N/A"
    #        label_3_1 = CTkLabel(cell_3_1, text=f"[{end_value}]", font=("Helvetica", 12))
    #        label_3_1.pack(pady=5)
#
    #    close_button = CTkButton(preview_frame, text="Закрыть", command=preview_window.destroy,
    #                             corner_radius=8, font=("Helvetica", 12))
    #    close_button.pack(pady=10)
    def convert_to_pdf(self):
        """Конвертирует каталог в PDF."""
        try:
            pdf_path = self.file_manager.convert_to_pdf()
            if pdf_path:
                self.status_var.set(f"Конвертировано в PDF: {pdf_path}")
                messagebox.showinfo("Успех", f"Конвертировано в PDF: {pdf_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))