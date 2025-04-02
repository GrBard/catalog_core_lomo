import os
import sys
import pandas as pd

def resource_path(relative_path):
    """Возвращает абсолютный путь к ресурсу, работает как в разработке, так и в .exe"""
    if hasattr(sys, '_MEIPASS'):
        # Если запущено как .exe, используем временную папку PyInstaller
        return os.path.join(sys._MEIPASS, relative_path)
    else:
        # Если запущено как скрипт, используем путь относительно корня проекта
        return os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', relative_path)

def find_continuous_intervals(df, start_col, end_col):
    """Вычисляет непрерывные интервалы на основе столбцов start_col и end_col."""
    # Преобразуем столбцы в числовой формат
    df[start_col] = pd.to_numeric(df[start_col], errors='coerce')
    df[end_col] = pd.to_numeric(df[end_col], errors='coerce')

    # Сортируем данные по началу интервала
    sorted_data = df[[start_col, end_col]].dropna().sort_values(by=start_col)

    if sorted_data.empty:
        return [], {}

    intervals = []
    interval_dict = {}
    current_start = sorted_data.iloc[0][start_col]
    current_end = sorted_data.iloc[0][end_col]

    for _, row in sorted_data.iterrows():
        start = row[start_col]
        end = row[end_col]

        if pd.isna(start) or pd.isna(end):
            continue

        if start <= current_end + 0.1:  # Допустим небольшой разрыв (0.1 м)
            current_end = max(current_end, end)
        else:
            intervals.append((current_start, current_end))
            for s in sorted_data[(sorted_data[start_col] >= current_start) & (sorted_data[start_col] <= current_end)][start_col]:
                interval_dict[s] = (current_start, current_end)
            current_start = start
            current_end = end

    # Добавляем последний интервал
    intervals.append((current_start, current_end))
    for s in sorted_data[(sorted_data[start_col] >= current_start) & (sorted_data[start_col] <= current_end)][start_col]:
        interval_dict[s] = (current_start, current_end)

    return intervals, interval_dict