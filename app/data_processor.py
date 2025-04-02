import os
import pandas as pd
import re
from pathlib import Path
from app.utils import find_continuous_intervals, resource_path
from docx import Document
from docx.shared import Inches, Cm
from PIL import Image, ImageDraw, ImageFont
import io
import math

class DataProcessor:
    def __init__(self, excel_path, images_folder):
        self.excel_path = Path(excel_path).resolve()
        self.images_folder = Path(images_folder).resolve()
        self.data = None
        self.all_image_files = []
        self.current_dataframe = None

    def load_excel(self):
        """Загружает данные из Excel в DataFrame."""
        self.data = pd.read_excel(self.excel_path)

    def load_image_files(self):
        """Получает список всех изображений из выбранной папки."""
        valid_extensions = (".jpg", ".jpeg", ".png", ".tif", ".tiff", ".bmp")
        self.all_image_files = [f.resolve() for f in self.images_folder.iterdir() if f.suffix.lower() in valid_extensions]

    def add_photo_columns(self):
        """Добавляет столбцы с путями к фотографиям и названием скважины."""
        if not self.all_image_files:
            self.load_image_files()

        if "BOX" not in self.data.columns:
            raise ValueError("В данных отсутствует столбец 'BOX' для сопоставления с фото.")

        def find_matching_file(box, uf=False):
            box_str = str(box)
            pattern = re.compile(rf"^(.+?)_{box_str}(_uf)?\.(jpg|jpeg|png|tif|tiff|bmp)$", re.IGNORECASE)
            for file_path in self.all_image_files:
                file_name = file_path.name.lower()
                match = pattern.match(file_name)
                if match:
                    if uf and "_uf" in file_name:
                        return str(file_path)
                    elif not uf and "_uf" not in file_name:
                        return str(file_path)
            return None

        def extract_well_name(photo_path):
            if not photo_path:
                return None
            file_name = Path(photo_path).stem
            match = re.match(r"^(.+?)_\d+(_uf)?$", file_name)
            if match:
                well_name = match.group(1)
                return well_name.replace("скв.", "").strip()
            return None

        self.data["Фото"] = self.data["BOX"].apply(lambda box: find_matching_file(box, uf=False))
        self.data["Фото УФ"] = self.data["BOX"].apply(lambda box: find_matching_file(box, uf=True))
        self.data["Скважина"] = self.data["Фото"].apply(extract_well_name)
        cols = list(self.data.columns)
        cols.insert(cols.index("BOX"), cols.pop(cols.index("Скважина")))
        self.data = self.data[cols]

    def compute_intervals(self):
        """Вычисляет непрерывные интервалы и добавляет их в DataFrame как 'Начало интервала' и 'Конец интервала'."""
        cols_lower = {col.lower(): col for col in self.data.columns}
        if "от" not in cols_lower or "до" not in cols_lower:
            raise ValueError("В данных отсутствуют столбцы 'от' или 'до' для вычисления интервалов.")

        start_col = cols_lower["от"]
        end_col = cols_lower["до"]

        _, interval_dict = find_continuous_intervals(self.data, start_col, end_col)

        def get_interval_start(row):
            start = float(row[start_col])
            return interval_dict.get(start, (None, None))[0]

        def get_interval_end(row):
            start = float(row[start_col])
            return interval_dict.get(start, (None, None))[1]

        self.data["Начало интервала"] = self.data.apply(get_interval_start, axis=1)
        self.data["Конец интервала"] = self.data.apply(get_interval_end, axis=1)

        cols = list(self.data.columns)
        do_col = cols_lower["до"]
        do_index = cols.index(do_col)
        cols.pop(cols.index("Начало интервала"))
        cols.pop(cols.index("Конец интервала"))
        cols.insert(do_index + 1, "Начало интервала")
        cols.insert(do_index + 2, "Конец интервала")
        self.data = self.data[cols]

    def process_data(self, samples_file=None):
        """Обрабатывает данные: загружает Excel, добавляет фото, вычисляет интервалы, объединяет с образцами."""
        self.load_excel()
        self.add_photo_columns()
        self.compute_intervals()

        if samples_file:
            self.merge_samples_data(samples_file)

        # Вычисляем "Вынос"
        cols_lower = {col.lower(): col for col in self.data.columns}
        required_cols_lower = ["замеры", "до", "от"]
        missing_cols = [col for col in required_cols_lower if col not in cols_lower]

        if not missing_cols:
            zamery_col = cols_lower["замеры"]
            do_col = cols_lower["до"]
            ot_col = cols_lower["от"]

            self.data[zamery_col] = pd.to_numeric(self.data[zamery_col], errors="coerce")
            self.data[do_col] = pd.to_numeric(self.data[do_col], errors="coerce")
            self.data[ot_col] = pd.to_numeric(self.data[ot_col], errors="coerce")

            self.data["Вынос"] = self.data.apply(
                lambda row: f"{row[zamery_col]} м ({(row[zamery_col] / (row[do_col] - row[ot_col]) * 100):.1f} %)"
                if pd.notna(row[zamery_col]) and pd.notna(row[do_col]) and pd.notna(row[ot_col]) and (row[do_col] - row[ot_col]) != 0
                else "N/A",
                axis=1
            )
            cols = list(self.data.columns)
            cols.insert(cols.index(zamery_col) + 1, cols.pop(cols.index("Вынос")))
            self.data = self.data[cols]

        self.current_dataframe = self.data.copy()
        return self.current_dataframe

    def merge_samples_data(self, samples_file):
        """Объединяет данные с образцами."""
        samples_df = pd.read_excel(samples_file)
        if "BOX" in samples_df.columns and "BOX" in self.data.columns:
            self.data = self.data.merge(samples_df, on="BOX", how="left")
        else:
            raise ValueError("Не удалось объединить данные об образцах: нет общего столбца 'BOX'")

    def generate_depth_scale(self, top_depth, bottom_depth, core_count, box_length=1.0):
        """Генерирует шкалу глубин для коробки."""
        # Параметры шкалы
        shift_up = 20  # Сдвиг сверху в пикселях
        shift_btm = 20  # Сдвиг снизу в пикселях
        image_height = 1100  # Высота изображения шкалы (пиксели)
        image_width = 50  # Ширина изображения шкалы
        step_rul = 0.1  # Шаг линейки (0.1 м)

        # Создаём изображение для первой шкалы
        im_d = Image.new('RGB', (image_width, image_height), (255, 255, 255))
        draw = ImageDraw.Draw(im_d)

        # Если есть вторая шкала (core_count == 2), создаём её
        if core_count == 2:
            im_d2 = Image.new('RGB', (image_width, image_height), (255, 255, 255))
            draw2 = ImageDraw.Draw(im_d2)
        else:
            draw2 = None

        # Загружаем шрифт
        try:
            font = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 14)
            font_small = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 10)
        except IOError:
            font = ImageFont.load_default()
            font_small = ImageFont.load_default()

        # Определяем направление шкалы (сверху вниз)
        rul_direct = 1  # 1: сверху вниз, -1: снизу вверх
        if rul_direct == 1:
            shift = shift_up
            h = (image_height - shift_up - shift_btm) / (box_length * 10)  # Масштаб в пикселях на дециметр
            dy_text = 15  # Смещение текста для первой отметки
        else:
            shift = image_height - shift_btm
            h = -(image_height - shift_up - shift_btm) / (box_length * 10)
            dy_text = -15

        # Рисуем отметки на шкале
        for depth in [top_depth, bottom_depth]:
            for ll in range(round((bottom_depth - top_depth) / step_rul) + 1):
                current_depth = top_depth + ll * step_rul
                # Вычисляем позицию на шкале
                dz = (current_depth - top_depth) / box_length  # Нормализованная позиция
                hy = (image_height - shift_up - shift_btm) * dz
                tsh = dy_text if ll == 0 else 0  # Смещение текста для первой отметки

                # Определяем, на какую шкалу рисовать (первая или вторая)
                target_draw = draw if ll < (round((bottom_depth - top_depth) / step_rul) + 1) // 2 or core_count == 1 else draw2

                # Рисуем отметки
                if current_depth % 1 < 0.00001:  # Каждые 1 м
                    target_draw.line((0, shift_up + hy, 50, shift_up + hy), fill="green", width=4)
                    target_draw.text((0, shift_up + hy + tsh), str(round(current_depth, 1)), (0, 0, 0), font=font)
                elif current_depth % 0.5 < 0.00001:  # Каждые 0.5 м
                    target_draw.line((0, shift_up + hy, 50, shift_up + hy), fill="green", width=2)
                    target_draw.text((0, shift_up + hy + tsh), str(round(current_depth, 1)), (0, 0, 0), font=font)
                else:  # Каждые 0.1 м
                    target_draw.line((0, shift_up + hy, 25, shift_up + hy), fill="green", width=2)

        # Сохраняем первую шкалу в поток
        img_d = io.BytesIO()
        im_d.save(img_d, 'JPEG')
        img_d.seek(0)

        # Если есть вторая шкала, сохраняем её
        if core_count == 2:
            img_d2 = io.BytesIO()
            im_d2.save(img_d2, 'JPEG')
            img_d2.seek(0)
            return img_d, img_d2
        return img_d, None

    def create_catalog(self, save_path):
        """Создаёт каталог в формате .docx."""
        if self.current_dataframe is None:
            raise ValueError("Нет данных для создания каталога.")

        # Путь к scale.jpg и shkala.jpg
        scale_image_path = resource_path('resources/scale.jpg')
        shkala_image_path = resource_path('resources/shkala.jpg')

        # Проверяем, существуют ли файлы
        if not os.path.exists(scale_image_path):
            raise FileNotFoundError(f"Файл {scale_image_path} не найден.")
        if not os.path.exists(shkala_image_path):
            raise FileNotFoundError(f"Файл {shkala_image_path} не найден.")

        # Высота scale.jpg = 21,88 см = 8,614 дюйма
        target_height = Inches(8.614)  # 21,88 см
        shkala_height = Inches(1)  # Высота shkala.jpg = 1 дюйм

        doc = Document()

        # Настройка отступов страницы
        sections = doc.sections
        for section in sections:
            section.top_margin = Cm(1)
            section.bottom_margin = Cm(1)
            section.left_margin = Cm(1)
            section.right_margin = Cm(1)

        # Заголовок документа
        doc.add_heading(f'Фотографии керна по скважине {self.current_dataframe["Скважина"].iloc[0]}', 0)

        # Вводный текст
        p = doc.add_paragraph('Глубины даны по керну. ')
        p.add_run('Номера образцов расположены напротив точек выбуривания. ')
        p.add_run('Номер образца состоит из двух цифр, разделённых точкой. ')
        p.add_run('Первая цифра соответствует номеру коробки, вторая – расстояние в сантиметрах от низа коробки до точки отбора образца. ')

        # Определяем размеры столбцов (увеличиваем второй столбец на 0,5 дюйма)
        width_samles_col = 0.9
        width_photo_col = 3.5  # Было 3, увеличиваем до 3,5
        width_samles_right = 4.0  # Было 4,5, уменьшаем до 4,0

        # Группируем данные по коробкам (BOX)
        grouped = self.current_dataframe.groupby('BOX')

        # Ищем столбцы "От" и "До" без учёта регистра
        cols_lower = {col.lower(): col for col in self.current_dataframe.columns}
        ot_col = cols_lower.get('от')
        do_col = cols_lower.get('до')

        if not ot_col or not do_col:
            raise ValueError("Не найдены столбцы 'От' и/или 'До' в DataFrame.")

        for box_number, group in grouped:
            doc.add_page_break()

            # Создаём таблицу: 4 строки, 4 столбца
            table = doc.add_table(rows=4, cols=4, style='Table Grid')

            # Устанавливаем ширину столбцов
            for cell in table.columns[0].cells:
                cell.width = Inches(width_samles_col)
            for cell in table.columns[1].cells:
                cell.width = Inches(width_photo_col)
            for cell in table.columns[2].cells:
                cell.width = Inches(width_samles_right)
            for cell in table.columns[3].cells:
                cell.width = Inches(0.5)

            # 1 строка, 1 столбец: "Коробка"
            cell = table.cell(0, 0)
            cell.text = f'Коробка {int(box_number)}'

            # 1 строка, 2 столбец: Интервал бурения
            cell = table.cell(0, 1)
            ts = 'Интервал бурения: '
            for idx, row in group.iterrows():
                ts += f"{row['Начало интервала']}-{row['Конец интервала']}\nвынос: {row['Вынос']}\n"
            cell.text = ts.strip()

            # 2 строка, 1 столбец: "Номера образцов"
            cell = table.cell(1, 0)
            cell.text = 'Номера образцов:'

            # 2 строка, 2 столбец: Значение "от"
            cell = table.cell(1, 1)
            ot_value = group[ot_col].iloc[0]
            cell.text = f'[{ot_value}]'

            # 4 строка, 2 столбец: Значение "до"
            cell = table.cell(3, 1)
            do_value = group[do_col].iloc[0]
            cell.text = f'[{do_value}]'

            # 2 строка, 3 столбец: "Исследования"
            cell = table.cell(1, 2)
            cell.text = 'Исследования:'

            # 3 строка, 2 столбец: Шкала и фотографии (горизонтально в ряд)
            cell = table.cell(2, 1)
            paragraph = cell.paragraphs[0]

            # Генерируем шкалу
            top_depth = float(group[ot_col].iloc[0])
            bottom_depth = float(group[do_col].iloc[0])
            core_count = 1  # Предполагаем 1 кусок керна
            img_d, img_d2 = self.generate_depth_scale(top_depth, bottom_depth, core_count)

            # Вставляем сгенерированную шкалу
            run = paragraph.add_run()
            run.add_picture(img_d, height=target_height)  # Устанавливаем высоту 21,88 см
            if img_d2:  # Если есть вторая шкала
                run = paragraph.add_run()
                run.add_picture(img_d2, height=target_height)

            # Вставляем основное фото коробки
            photo_path = group["Фото"].iloc[0]
            if pd.notna(photo_path) and os.path.exists(photo_path):
                run = paragraph.add_run()
                run.add_picture(photo_path, height=target_height)  # Устанавливаем высоту 21,88 см

            # Вставляем shkala.jpg (высота 1 дюйм)
            run = paragraph.add_run()
            run.add_picture(shkala_image_path, height=shkala_height)  # Устанавливаем высоту 1 дюйм

            # Вставляем УФ-фото коробки
            photo_uf_path = group["Фото УФ"].iloc[0]
            if pd.notna(photo_uf_path) and os.path.exists(photo_uf_path):
                run = paragraph.add_run()
                run.add_picture(photo_uf_path, height=target_height)  # Устанавливаем высоту 21,88 см

            # 3 строка, 4 столбец: Вставляем scale.jpg
            cell = table.cell(2, 3)
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(scale_image_path, height=target_height)  # Устанавливаем высоту 21,88 см

        # Сохраняем документ
        doc.save(save_path)
        return save_path

    def get_current_dataframe(self):
        """Возвращает текущий DataFrame."""
        return self.current_dataframe
