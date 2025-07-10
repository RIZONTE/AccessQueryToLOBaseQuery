import get_access_meta
import get_csv_excel
import access_queries

import os

import os

# Получаем абсолютный путь к каталогу со скриптом
script_dir = os.path.dirname(os.path.abspath(__file__))

# Конфигурация
access_db_path = os.path.join(script_dir, r"input\Сверка — копия.accdb")                 # путь до файла access
meta_output_file = os.path.join(script_dir, r"output\access_meta.txt")                   # путь до выходного файла с метаданными таблиц
queries_out_file = os.path.join(script_dir, r"output\queries.txt")                       # путь до выходного файла с запросами
csv_files_path = os.path.join(script_dir, r"output\\")                                   # каталог куда сохраняются csv

excel_files = [
    os.path.join(script_dir, r"input\0503124 — копия.xls"),                              # список excel-файлов
    os.path.join(script_dir, r"input\0506604 — копия.xls"),
]
sheet_names = [
    '2',                                                                                 # список листов которые импортируются
    'Лист1'
]


# Основная часть

# Работа с базой MS Access
if not os.path.exists(access_db_path):
    print(f"Файл базы данных не найден: {access_db_path}")
else:
    # Извлечение и конвертация запросов
    access_queries.export_access_queries(access_db_path, queries_out_file)
    print(f"Запросы сконвертированы и извлечены в {queries_out_file} ")

    # Извлечение метаданных из базы Access
    get_access_meta.get_access_metadata(access_db_path, meta_output_file)
    print(f"Метаданные успешно экспортированы в {meta_output_file} ")

# Формирование csv файлов для импорта в LO Base
table_names, columns, data_types = get_csv_excel.parse_schema_file(meta_output_file)

for i in range(len(excel_files)):
    csv_output = csv_files_path +"/" + table_names[i] + ".csv"
    if not get_csv_excel.process_excel_to_csv(excel_files[i], sheet_names[i], csv_output):
        print("Ошибка формирования csv")

    print(f"Обрабатывается таблица: {table_names[i]}")
    print(f"Столбцы: {columns[i]}")
    print(f"Типы данных: {data_types[i]}")

    get_csv_excel.clean_numeric_columns(csv_output, columns[i], data_types[i])