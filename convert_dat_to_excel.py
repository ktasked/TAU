#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Скрипт для преобразования данных из DAT файлов в формат Excel.
Извлекает параметры Treg и Rd, сортирует по возрастанию Rd,
группирует по Ko и определяет диапазоны номеров экспериментов.
"""

import re
import openpyxl
from openpyxl import Workbook


def parse_dat_file(filepath):
    """
    Парсит DAT файл и извлекает данные из таблицы.
    
    Args:
        filepath: Путь к DAT файлу
        
    Returns:
        Список словарей с данными экспериментов
    """
    data = []
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as f:
        content = f.read()
    
    # Ищем строки таблицы вида: |  N|  Ko  |   To,с.   | Tau,с.|  Treg, с. |   Rd   |
    pattern = r'\|\s*(\d+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|\s*([\d.]+)\s*\|'
    matches = re.findall(pattern, content)
    
    for match in matches:
        n, ko, to, tau, treg, rd = match
        data.append({
            'N': int(n),
            'Ko': float(ko),
            'To': float(to),
            'Tau': float(tau),
            'Treg': float(treg),
            'Rd': float(rd)
        })
    
    return data


def remove_duplicates(data):
    """Убирает дубликаты по номеру эксперимента N."""
    unique_data = []
    seen_n = set()
    for d in data:
        if d['N'] not in seen_n:
            seen_n.add(d['N'])
            unique_data.append(d)
    return unique_data


def group_by_ko(data):
    """
    Группирует данные по коэффициенту Ko.
    
    Args:
        data: Список словарей с данными
        
    Returns:
        Словарь {Ko: [список экспериментов]}
    """
    groups = {}
    for d in data:
        ko = d['Ko']
        if ko not in groups:
            groups[ko] = []
        groups[ko].append(d)
    
    # Сортируем каждую группу по возрастанию Rd
    for ko in groups:
        groups[ko].sort(key=lambda x: x['Rd'])
    
    return groups


def get_cycle_ranges(groups):
    """
    Определяет диапазоны номеров экспериментов для каждого цикла (группы Ko).
    
    Args:
        groups: Словарь групп по Ko
        
    Returns:
        Словарь {Ko: {'start': мин. N, 'end': макс. N, 'count': количество}}
    """
    ranges = {}
    for ko, experiments in groups.items():
        n_values = [e['N'] for e in experiments]
        ranges[ko] = {
            'start': min(n_values),
            'end': max(n_values),
            'count': len(n_values),
            'n_list': sorted(n_values)
        }
    return ranges


def create_excel_output(all_data, output_path):
    """
    Создает Excel файл с отформатированными данными.
    
    Args:
        all_data: Словарь {имя_регулятора: данные}
        output_path: Путь для сохранения Excel файла
    """
    wb = Workbook()
    
    # Создаем сводный лист
    summary_ws = wb.active
    summary_ws.title = "Сводные данные"
    
    # Заголовки для сводного листа
    headers = ["Регулятор", "Ko", "Начало цикла (N)", "Конец цикла (N)", 
               "Количество экспериментов", "Диапазон Rd", "Диапазон Treg"]
    summary_ws.append(headers)
    
    row_num = 2
    for reg_name, data in all_data.items():
        groups = group_by_ko(data)
        ranges = get_cycle_ranges(groups)
        
        for ko in sorted(ranges.keys()):
            range_info = ranges[ko]
            rd_values = [e['Rd'] for e in groups[ko]]
            treg_values = [e['Treg'] for e in groups[ko]]
            
            summary_ws.append([
                reg_name,
                ko,
                range_info['start'],
                range_info['end'],
                range_info['count'],
                f"{min(rd_values):.3f} - {max(rd_values):.3f}",
                f"{min(treg_values):.1f} - {max(treg_values):.1f}"
            ])
            row_num += 1
    
    # Создаем листы для каждого регулятора с детальными данными
    for reg_name, data in all_data.items():
        detail_ws = wb.create_sheet(title=reg_name[:31])  # Ограничение на имя листа
        groups = group_by_ko(data)
        
        # Заголовки
        detail_headers = ["Цикл (Ko)", "№ п/п", "N", "Ko", "To", "Tau", "Treg", "Rd"]
        detail_ws.append(detail_headers)
        
        row_num = 2
        for ko in sorted(groups.keys()):
            experiments = groups[ko]
            for i, exp in enumerate(experiments, 1):
                detail_ws.append([
                    ko,
                    i,
                    exp['N'],
                    exp['Ko'],
                    exp['To'],
                    exp['Tau'],
                    exp['Treg'],
                    exp['Rd']
                ])
                row_num += 1
    
    wb.save(output_path)
    print(f"Excel файл сохранен: {output_path}")


def print_summary(all_data):
    """Выводит сводную информацию в консоль."""
    print("\n" + "="*80)
    print("СВОДНАЯ ИНФОРМАЦИЯ ПО ДАННЫМ ИЗ DAT ФАЙЛОВ")
    print("="*80)
    
    for reg_name, data in all_data.items():
        print(f"\n{reg_name}:")
        print("-" * 60)
        
        groups = group_by_ko(data)
        ranges = get_cycle_ranges(groups)
        
        for ko in sorted(ranges.keys()):
            range_info = ranges[ko]
            experiments = groups[ko]
            rd_values = [e['Rd'] for e in experiments]
            treg_values = [e['Treg'] for e in experiments]
            
            print(f"\n  Цикл Ko={ko}:")
            print(f"    Номера экспериментов: с {range_info['start']} по {range_info['end']}")
            print(f"    Всего экспериментов: {range_info['count']}")
            print(f"    Диапазон Rd: {min(rd_values):.3f} - {max(rd_values):.3f}")
            print(f"    Диапазон Treg: {min(treg_values):.1f} - {max(treg_values):.1f} сек")
            print(f"    Отсортировано по возрастанию Rd")


def main():
    """Основная функция."""
    # Пути к файлам
    files = {
        'MOM_PI': '/workspace/MOM_PI.dat',
        'MOM_PID': '/workspace/MOM_PID.dat',
        'LAMBDA_PI': '/workspace/LAMBDA_PI.dat',
        'LAMBDA_PID': '/workspace/LAMBDA_PID.dat'
    }
    
    # Парсим все файлы
    all_data = {}
    for name, filepath in files.items():
        print(f"Обработка файла: {filepath}")
        data = parse_dat_file(filepath)
        unique_data = remove_duplicates(data)
        all_data[name] = unique_data
        print(f"  Найдено {len(unique_data)} уникальных экспериментов")
    
    # Выводим сводку
    print_summary(all_data)
    
    # Создаем Excel файл
    output_path = '/workspace/результаты_преобразования.xlsx'
    create_excel_output(all_data, output_path)
    
    # Дополнительный вывод с подробными данными
    print("\n" + "="*80)
    print("ПОДРОБНЫЕ ДАННЫЕ ПО КАЖДОМУ РЕГУЛЯТОРУ")
    print("="*80)
    
    for reg_name, data in all_data.items():
        print(f"\n{'='*60}")
        print(f"{reg_name} - данные отсортированные по Ko и Rd")
        print('='*60)
        
        groups = group_by_ko(data)
        
        for ko in sorted(groups.keys()):
            experiments = groups[ko]
            n_values = [e['N'] for e in experiments]
            
            print(f"\n  >>> ЦИКЛ Ko={ko} (N с {min(n_values)} по {max(n_values)}) <<<")
            print(f"  {'N':>3} | {'Ko':>5} | {'To':>7} | {'Tau':>7} | {'Treg':>7} | {'Rd':>7}")
            print(f"  {'-'*3}|{'-'*7}|{'-'*9}|{'-'*9}|{'-'*9}|{'-'*9}")
            
            for exp in experiments:
                print(f"  {exp['N']:>3} | {exp['Ko']:>5.2f} | {exp['To']:>7.2f} | {exp['Tau']:>7.2f} | {exp['Treg']:>7.1f} | {exp['Rd']:>7.3f}")


if __name__ == '__main__':
    main()
