import os
import json
import pandas as pd
from typing import List, Dict
from openpyxl import load_workbook

# Путь к папке с JSON-файлами
json_directory = 'C:/PowerBI/test_1/Model/tables'

# Список файлов JSON для каждой таблицы
json_files: Dict[str, List[str]] = {
    'Животные': ['Котики.json', 'Собачки.json'],
    'Таблица 1 (TemplateImportOU)': ['Внешний идентификатор для импорта.json', 'Вышестоящий отдел.json', 'Название департамента.json'],
    'Таблица 2': ['ВажнаяМера.json', 'ДатаDate.json'],
    'Таблица (3)': ['Столбец1.json', 'Столбец2.json']
}

def parse_json_files(json_directory: str, json_files: Dict[str, List[str]]) -> List[List[str]]:
    """
    Парсит JSON-файлы для извлечения информации о таблицах и колонках.

    Args:
        json_directory (str): Путь к корневой папке с JSON-файлами.
        json_files (Dict[str, List[str]]): Словарь, где ключи - названия таблиц, а значения - списки файлов JSON для этих таблиц.

    Returns:
        List[List[str]]: Список строк с данными для каждой колонки (таблица, колонка, репрезентативная колонка, категория).
    """
    data = []
    for table_name, files in json_files.items():
        # Парсим информацию о таблице
        table_info_path = os.path.join(json_directory, table_name, 'table.json')
        if os.path.exists(table_info_path):
            with open(table_info_path, 'r', encoding='utf-8') as f:
                content = json.load(f)
                table_name = content['name']
                # Парсим информацию о колонках
                for file_name in files:
                    column_info_path = os.path.join(json_directory, table_name, 'columns', file_name)
                    if os.path.exists(column_info_path):
                        with open(column_info_path, 'r', encoding='utf-8') as f:
                            column_content = json.load(f)
                            column_name = column_content['name']
                            rep_column = f"{table_name}.{column_name}"
                            if "measures" in column_info_path.lower():
                                category = "Measure"
                            elif "columns" in column_info_path.lower():
                                category = "Column"
                            elif "hierarchies" in column_info_path.lower():
                                category = "Hierarchy"
                            else:
                                category = "Unknown"
                            data.append([table_name, column_name, rep_column, category])
    return data

# Парсинг JSON-файлов и создание DataFrame
data = parse_json_files(json_directory, json_files)

# Создание DataFrame и сохранение в Excel
df = pd.DataFrame(data, columns=['Table', 'Column', 'RepColumn', 'Category'])
excel_path = 'parsed_table_columns.xlsx'
df.to_excel(excel_path, index=False)

# Настройка ширины столбцов
def adjust_column_width(excel_path: str):
    """
    Настраивает ширину столбцов в Excel-файле на основе содержимого.

    Args:
        excel_path (str): Путь к Excel-файлу.
    """
    workbook = load_workbook(excel_path)
    worksheet = workbook.active

    for col in worksheet.columns:
        max_length = 0
        column = col[0].column_letter # Получаем букву столбца
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(cell.value)
            except:
                pass
        adjusted_width = (max_length + 2)
        worksheet.column_dimensions[column].width = adjusted_width

    workbook.save(excel_path)

# Применение настройки ширины столбцов
adjust_column_width(excel_path)

print("Парсинг завершен. Данные сохранены в 'parsed_table_columns.xlsx'")
