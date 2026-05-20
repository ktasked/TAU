#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для преобразования DAT файлов в Excel с данными для построения 
круговых областей показателей качества САУ
"""

import re
import pandas as pd
from pathlib import Path
import numpy as np

def parse_dat_file(filepath):
    """Парсинг DAT файла и извлечение данных"""
    data = []
    
    # Пробуем разные кодировки
    encodings = ['utf-8', 'cp1251', 'latin-1', 'iso-8859-1']
    content = None
    
    for encoding in encodings:
        try:
            with open(filepath, 'r', encoding=encoding) as f:
                content = f.read()
            break
        except UnicodeDecodeError:
            continue
    
    if content is None:
        print(f"Не удалось прочитать файл {filepath} ни в одной кодировке")
        return pd.DataFrame()
    
    # Поиск строк с данными формата: | N | Ko | To | Tau | Treg | Rd |
    pattern = r'\|\s*(\d+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|'
    
    matches = re.findall(pattern, content)
    
    for match in matches:
        N = int(match[0])
        Ko = float(match[1])
        To = float(match[2])
        Tau = float(match[3])
        Treg = float(match[4])
        Rd = float(match[5])
        
        data.append({
            'N': N,
            'Ko': Ko,
            'To': To,
            'Tau': Tau,
            'Treg': Treg,
            'Rd': Rd
        })
    
    return pd.DataFrame(data)

def order_points_as_circle(df, x_col='Treg', y_col='Rd'):
    """
    Упорядочивает точки так, чтобы они образовывали замкнутый контур (круг)
    Начинает с минимального значения X, проходит против часовой стрелки и заканчивается там же
    """
    if len(df) == 0:
        return df
    
    # Вычисляем центр масс точек
    x_center = df[x_col].mean()
    y_center = df[y_col].mean()
    
    # Вычисляем угол для каждой точки относительно центра
    df = df.copy()
    df['angle'] = np.arctan2(df[y_col] - y_center, df[x_col] - x_center)
    
    # Находим точку с минимальным X (или ближайшую к минимальному)
    min_x_idx = df[x_col].idxmin()
    min_x_angle = df.loc[min_x_idx, 'angle']
    
    # Сортируем по углу, начиная с точки с минимальным X
    # Сдвигаем углы так, чтобы минимальный X был началом
    df['angle_shifted'] = (df['angle'] - min_x_angle + 2 * np.pi) % (2 * np.pi)
    
    # Сортируем по смещенному углу
    df_sorted = df.sort_values('angle_shifted').copy()
    
    # Добавляем первую точку в конец для замыкания круга
    first_row = df_sorted.iloc[[0]].copy()
    df_closed = pd.concat([df_sorted, first_row], ignore_index=True)
    
    # Удаляем вспомогательные колонки
    cols_to_drop = ['angle', 'angle_shifted']
    df_closed = df_closed.drop(columns=cols_to_drop, errors='ignore')
    
    return df_closed

def smooth_circle_points(df, x_col='Treg', y_col='Rd', alpha=0.3):
    """
    Сглаживает точки для создания более круглой формы
    alpha - коэффициент сглаживания (0 - нет сглаживания, 1 - сильное сглаживание)
    """
    if len(df) < 3:
        return df
    
    df = df.copy()
    
    # Вычисляем центр
    x_center = df[x_col].mean()
    y_center = df[y_col].mean()
    
    # Вычисляем среднее расстояние от центра
    distances = np.sqrt((df[x_col] - x_center)**2 + (df[y_col] - y_center)**2)
    avg_distance = distances.mean()
    
    # Сглаживаем расстояния, стремясь к среднему
    df[x_col] = x_center + (df[x_col] - x_center) * (1 - alpha) + \
                (df[x_col] - x_center) / distances * avg_distance * alpha
    df[y_col] = y_center + (df[y_col] - y_center) * (1 - alpha) + \
                (df[y_col] - y_center) / distances * avg_distance * alpha
    
    return df

def process_file(filepath, output_writer):
    """Обработка одного файла и запись в Excel"""
    filename = Path(filepath).stem
    
    # Парсинг данных
    df = parse_dat_file(filepath)
    
    if df.empty:
        print(f"Файл {filepath} не содержит данных")
        return
    
    # Группировка по Ko (настройкам регулятора)
    ko_groups = df.groupby('Ko')
    
    all_data_cycles = []
    
    for ko_value, group in ko_groups:
        # Упорядочиваем точки как круг
        group_circle = order_points_as_circle(group)
        
        # Добавляем информацию о цикле
        group_circle['Cycle_Ko'] = ko_value
        group_circle['Cycle_Start'] = group_circle['Treg'].iloc[0]
        group_circle['Cycle_End'] = group_circle['Treg'].iloc[-1]
        
        all_data_cycles.append(group_circle)
    
    # Объединяем все циклы
    if all_data_cycles:
        df_cycles = pd.concat(all_data_cycles, ignore_index=True)
        
        # Записываем подробные данные на отдельный лист
        sheet_name = filename[:31]  # Excel ограничивает имя листа 31 символом
        df_cycles.to_excel(output_writer, sheet_name=sheet_name, index=False, float_format='%.3f')
        
        print(f"\nФайл: {filename}")
        print(f"Найдено настроек регулятора (Ko): {len(ko_groups)}")
        
        # Выводим пример данных для первого цикла
        for ko_value, group in list(ko_groups)[:1]:
            group_circle = order_points_as_circle(group)
            print(f"\nПример цикла для Ko = {ko_value}:")
            print("Treg\tRd")
            for _, row in group_circle.iterrows():
                print(f"{row['Treg']:.3f}\t{row['Rd']:.3f}")

def main():
    # Поиск всех DAT файлов
    dat_files = list(Path('/workspace').glob('*.dat'))
    
    if not dat_files:
        print("DAT файлы не найдены в /workspace")
        return
    
    print(f"Найдено DAT файлов: {len(dat_files)}")
    
    # Создание Excel файла
    output_file = '/workspace/результаты_круги.xlsx'
    
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for dat_file in dat_files:
            process_file(dat_file, writer)
    
    print(f"\n✅ Данные сохранены в файл: {output_file}")
    print("\nДанные организованы следующим образом:")
    print("- Каждый лист соответствует одному DAT файлу")
    print("- Данные отсортированы так, чтобы образовывать замкнутый круг")
    print("- Цикл начинается с минимального значения Treg и заканчивается тем же значением")
    print("- Для каждого цикла указана настройка регулятора Ko")

if __name__ == '__main__':
    main()
