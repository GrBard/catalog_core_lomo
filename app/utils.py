import os  # Модуль для работы с файлами и папками
import sys  # Модуль для работы с системными параметрами
import pandas as pd  # Библиотека для работы с таблицами

# Функция для получения пути к ресурсам (работает и в .exe)
def resource_path(relative_path):
    """Возвращает абсолютный путь к ресурсу, работает как в разработке, так и в .exe"""
    if hasattr(sys, '_MEIPASS'):  # Если программа запущена как .exe (PyInstaller)
        return os.path.join(sys._MEIPASS, relative_path)  # Путь из временной папки PyInstaller
    else:  # Если запущена как обычный скрипт
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', relative_path)  # Путь относительно файла

# Функция для поиска непрерывных интервалов в таблице
def find_continuous_intervals(df, start_col, end_col):
    """Вычисляет непрерывные интервалы на основе столбцов start_col и end_col."""
    # Преобразуем столбцы в числа, заменяя ошибки на NaN
    df[start_col] = pd.to_numeric(df[start_col], errors='coerce')
    df[end_col] = pd.to_numeric(df[end_col], errors='coerce')

    # Сортируем данные по столбцу начала
    sorted_data = df[[start_col, end_col]].dropna().sort_values(by=start_col)

    # Если данных нет, возвращаем пустые списки
    if sorted_data.empty:
        return [], {}

    intervals = []  # Список интервалов
    interval_dict = {}  # Словарь для сопоставления начала интервала с его началом и концом
    current_start = sorted_data.iloc[0][start_col]  # Начало первого интервала
    current_end = sorted_data.iloc[0][end_col]  # Конец первого интервала

    # Проходим по всем строкам отсортированных данных
    for _, row in sorted_data.iterrows():
        start = row[start_col]  # Начало текущей строки
        end = row[end_col]  # Конец текущей строки

        if pd.isna(start) or pd.isna(end):  # Если данные пустые, пропускаем
            continue

        # Если начало текущей строки близко к концу предыдущего интервала
        if start <= current_end + 0.1:  # Допускаем разрыв 0.1 м
            current_end = max(current_end, end)  # Увеличиваем конец интервала
        else:  # Если начинается новый интервал
            intervals.append((current_start, current_end))  # Добавляем предыдущий интервал
            # Записываем начало и конец для всех точек в интервале
            for s in sorted_data[(sorted_data[start_col] >= current_start) & (sorted_data[start_col] <= current_end)][start_col]:
                interval_dict[s] = (current_start, current_end)
            current_start = start  # Начинаем новый интервал
            current_end = end

    # Добавляем последний интервал
    intervals.append((current_start, current_end))
    for s in sorted_data[(sorted_data[start_col] >= current_start) & (sorted_data[start_col] <= current_end)][start_col]:
        interval_dict[s] = (current_start, current_end)

    return intervals, interval_dict  # Возвращаем список интервалов и словарь