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
import time

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
        """Обрабатывает данные: загружает Excel, добавляет фото, вычисляет интервалы."""
        self.load_excel()
        self.add_photo_columns()
        self.compute_intervals()

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
                if pd.notna(row[zamery_col]) and pd.notna(row[do_col]) and pd.notna(row[ot_col]) and (
                            row[do_col] - row[ot_col]) != 0
                else "N/A",
                axis=1
            )
            cols = list(self.data.columns)
            cols.insert(cols.index(zamery_col) + 1, cols.pop(cols.index("Вынос")))
            self.data = self.data[cols]

        self.current_dataframe = self.data.copy()
        # Сортируем по столбцу "от"
        if "от" in self.current_dataframe.columns:
            self.current_dataframe = self.current_dataframe.sort_values(by="от")
        return self.current_dataframe

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
            font_path = resource_path('resources/arial.ttf')
            font = ImageFont.truetype(font_path, 14)
            font_small = ImageFont.truetype(font_path, 10)
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

    def draw_sample_circles(self, photo_path, samples_in_box, core_count, suffix='_with_circles'):
        """Рисует кружки на копии фото керна в местах отбора образцов, не изменяя оригинал."""
        # Открываем изображение
        img = Image.open(photo_path)
        # Создаём копию изображения
        img_copy = Image.new('RGB', img.size)
        img_copy.paste(img)  # Копируем содержимое, чтобы гарантированно не изменять оригинал
        draw = ImageDraw.Draw(img_copy)

        # Параметры изображения
        img_width, img_height = img_copy.size
        shift_up = 0  # Отступ сверху (можно настроить, если нужно)
        shift_btm = 0  # Отступ снизу (можно настроить, если нужно)

        # Радиус кружка (12% от ширины изображения)
        r = img_width * 0.12

        # Вычисляем позицию по горизонтали (центр для одного ядра)
        hx = img_width / (2 * core_count)

        # Настраиваем шрифт для подписи
        try:
            font_path = resource_path('resources/arial.ttf')
            font = ImageFont.truetype(font_path, int(r * 1))
        except IOError:
            font = ImageFont.load_default()

        # Отступы для текста
        sht_txt_g = r  # Горизонтальный отступ текста
        sht_txt_v = r  # Вертикальный отступ текста

        # Обрабатываем каждый образец в коробке
        for _, sample in samples_in_box.iterrows():
            sample_num = sample['Номер образца']
            # Извлекаем десятичную часть номера образца (глубина в метрах от верха коробки)
            depth_in_box = sample_num - int(sample_num)  # Например, для 3.15 это 0.15 м

            # Вычисляем вертикальную позицию кружка
            # Коробка 1 м, высота изображения (за вычетом отступов) соответствует 1 м
            hy = (img_height - shift_up - shift_btm) * depth_in_box

            # Корректируем положение, если кружок близко к краю
            if shift_up + hy + r > 0.95 * img_height:  # Близко к нижнему краю
                hy = 0.95 * img_height - shift_up - r
            elif shift_up + hy - r < 0.05 * img_height:  # Близко к верхнему краю
                hy = 0.05 * img_height - shift_up + r

            # Рисуем кружок
            draw.arc(
                (hx - r, shift_up + hy - r, hx + r, shift_up + hy + r),
                0, 360, fill=(255, 255, 0), width=int(r / 10)
            )

            # Добавляем подпись (номер образца)
            if shift_up + hy + r > 0.95 * img_height:
                # Если кружок внизу, подпись выше
                draw.text(
                    (hx - sht_txt_g, shift_up + hy - r - sht_txt_v),
                    str(sample_num),
                    (255, 255, 0),
                    font=font
                )
            else:
                # Иначе подпись ниже
                draw.text(
                    (hx - sht_txt_g, shift_up + hy + r),
                    str(sample_num),
                    (255, 255, 0),
                    font=font
                )

        # Закрываем исходное изображение
        img.close()

        # Сохраняем изменённое изображение во временный файл
        temp_img = io.BytesIO()
        img_copy.save(temp_img, 'PNG')
        temp_img.seek(0)

        # Создаём уникальный временный путь, чтобы избежать перезаписи
        base, ext = os.path.splitext(photo_path)
        temp_path = f"{base}{suffix}_{int(time.time())}{ext}"  # Добавляем временную метку для уникальности
        with open(temp_path, 'wb') as f:
            f.write(temp_img.getvalue())

        # Проверяем, что исходный файл не изменился
        if os.path.getsize(photo_path) != os.path.getsize(temp_path):
            print(f"Исходный файл {photo_path} не изменён.")
        else:
            print(f"Внимание: исходный файл {photo_path} может быть изменён!")

        return temp_path

    def create_catalog(self, save_path, samples_df=None, progress_bar=None, progress_step=1.0):
        """Создаёт каталог в формате .docx, с учётом образцов, если они переданы."""
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
        p.add_run(
            'Первая цифра соответствует номеру коробки, вторая – расстояние в сантиметрах от низа коробки до точки отбора образца. ')

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

        # Инициализируем прогресс
        current_progress = 0.0

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

            # 3 строка, 1 столбец: Сами номера образцов
            cell = table.cell(2, 0)
            if samples_df is not None:
                samples_in_box = samples_df[samples_df['BOX'] == int(box_number)]
                sample_numbers = samples_in_box['Номер образца'].tolist()
                cell.text = '\n'.join(map(str, sample_numbers))
            else:
                cell.text = ''

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

            # 3 строка, 3 столбец: Сами исследования в формате "номер образца \tab исследования"
            cell = table.cell(2, 2)
            if samples_df is not None:
                samples_in_box = samples_df[samples_df['BOX'] == int(box_number)]
                if not samples_in_box.empty:
                    # Вычисляем высоту ячейки (примерно, в пикселях или условных единицах)
                    # Предполагаем, что высота ячейки соответствует высоте фото (21,88 см = 8,614 дюйма)
                    cell_height = 8.614  # Высота в дюймах (примерно соответствует высоте фото)
                    img_height = 8.614 * 96  # Примерно 96 пикселей на дюйм (для расчётов)

                    # Параметры, аналогичные draw_sample_circles
                    shift_up = 0
                    shift_btm = 0

                    # Отслеживаем текущую позицию (в строках)
                    current_lines = 0

                    # Сортируем образцы по номеру, чтобы подписи шли сверху вниз
                    samples_in_box = samples_in_box.sort_values(by='Номер образца')

                    for idx, (_, sample) in enumerate(samples_in_box.iterrows()):
                        sample_num = sample['Номер образца']
                        # Извлекаем десятичную часть номера образца (глубина в метрах от верха коробки)
                        depth_in_box = sample_num - int(sample_num)  # Например, для 8.71 это 0.71 м

                        # Вычисляем вертикальную позицию кружка (аналогично draw_sample_circles)
                        hy = (img_height - shift_up - shift_btm) * depth_in_box

                        # Корректируем положение, если кружок близко к краю (аналогично draw_sample_circles)
                        r = img_height * 0.12  # Радиус кружка (12% от высоты изображения, как в draw_sample_circles)
                        if shift_up + hy + r > 0.95 * img_height:  # Близко к нижнему краю
                            hy = 0.95 * img_height - shift_up - r
                        elif shift_up + hy - r < 0.05 * img_height:  # Близко к верхнему краю
                            hy = 0.05 * img_height - shift_up + r

                        # Вычисляем, сколько строк должно быть до этой подписи
                        # Калибровочный коэффициент для высоты строки
                        line_height = 0.19 * 96  # Уменьшаем высоту строки для более точного соответствия
                        desired_lines = int(hy / line_height)

                        # Вычисляем, сколько пустых строк нужно добавить
                        lines_to_skip = max(0, desired_lines - current_lines)

                        # Добавляем пустые строки
                        for _ in range(lines_to_skip):
                            cell.add_paragraph()
                            current_lines += 1

                        # Добавляем подпись
                        paragraph = cell.add_paragraph()
                        run = paragraph.add_run()
                        run.text = f"{sample['Номер образца']}"
                        run.add_tab()
                        run = paragraph.add_run()
                        run.text = f"{sample['Исследования']}"
                        current_lines += 1  # Учитываем строку с подписью

                        # Добавляем минимальный отступ (1 пустая строка) перед следующей подписью
                        cell.add_paragraph()
                        current_lines += 1
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

            # Вставляем основное фото коробки с кружками, если есть образцы
            photo_path = group["Фото"].iloc[0]
            if pd.notna(photo_path) and os.path.exists(photo_path):
                run = paragraph.add_run()
                if samples_df is not None:
                    samples_in_box = samples_df[samples_df['BOX'] == int(box_number)]
                    if not samples_in_box.empty:
                        # Рисуем кружки на копии фото
                        print(f"Обрабатываем основное фото: {photo_path}")
                        modified_photo_path = self.draw_sample_circles(photo_path, samples_in_box, core_count,
                                                                       suffix='_with_circles')
                        print(f"Сохранено с кружками (основное): {modified_photo_path}")
                        run.add_picture(modified_photo_path, height=target_height)
                        # Удаляем временный файл после использования
                        try:
                            os.remove(modified_photo_path)
                            print(f"Удалён временный файл: {modified_photo_path}")
                        except Exception as e:
                            print(f"Ошибка удаления временного файла {modified_photo_path}: {e}")
                    else:
                        print(f"Используем оригинальное основное фото: {photo_path}")
                        run.add_picture(photo_path, height=target_height)
                else:
                    print(f"Используем оригинальное основное фото: {photo_path}")
                    run.add_picture(photo_path, height=target_height)

            # Вставляем shkala.jpg (высота 1 дюйм)
            run = paragraph.add_run()
            run.add_picture(shkala_image_path, height=shkala_height)  # Устанавливаем высоту 1 дюйм

            # Вставляем УФ-фото коробки с кружками, если есть образцы
            photo_uf_path = group["Фото УФ"].iloc[0]
            if pd.notna(photo_uf_path) and os.path.exists(photo_uf_path):
                run = paragraph.add_run()
                if samples_df is not None:
                    samples_in_box = samples_df[samples_df['BOX'] == int(box_number)]
                    if not samples_in_box.empty:
                        # Рисуем кружки на копии УФ-фото
                        print(f"Обрабатываем УФ-фото: {photo_uf_path}")
                        modified_uf_photo_path = self.draw_sample_circles(photo_uf_path, samples_in_box, core_count,
                                                                          suffix='_uf_with_circles')
                        print(f"Сохранено с кружками (УФ): {modified_uf_photo_path}")
                        run.add_picture(modified_uf_photo_path, height=target_height)
                        # Удаляем временный файл после использования
                        try:
                            os.remove(modified_uf_photo_path)
                            print(f"Удалён временный файл: {modified_uf_photo_path}")
                        except Exception as e:
                            print(f"Ошибка удаления временного файла {modified_uf_photo_path}: {e}")
                    else:
                        print(f"Используем оригинальное УФ-фото: {photo_uf_path}")
                        run.add_picture(photo_uf_path, height=target_height)
                else:
                    print(f"Используем оригинальное УФ-фото: {photo_uf_path}")
                    run.add_picture(photo_uf_path, height=target_height)

            # 3 строка, 4 столбец: Вставляем scale.jpg
            cell = table.cell(2, 3)
            paragraph = cell.paragraphs[0]
            run = paragraph.add_run()
            run.add_picture(scale_image_path, height=target_height)  # Устанавливаем высоту 21,88 см

            # Обновляем прогресс
            if progress_bar is not None:
                current_progress += progress_step
                progress_bar.set(min(current_progress, 1.0))
                progress_bar.update()

        # Сохраняем документ
        doc.save(save_path)
        return save_path

    def get_current_dataframe(self):
        """Возвращает текущий DataFrame."""
        return self.current_dataframe
