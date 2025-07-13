import pandas as pd
import re
import os

def process_excel_to_csv(excel_file, sheet_name, output_csv):
    """Читает указанный лист из Excel файла и сохраняет в CSV"""
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
        df.to_csv(output_csv, index=False, encoding='utf-8')
        print(f"Файл CSV успешно создан: {output_csv}")
        return True
    except Exception as e:
        print(f"Ошибка при обработке Excel файла: {e}")
        return False

def parse_schema_file(schema_file):
    """Парсит файл схемы и возвращает информацию о таблице"""
    table_name = []
    columns = []
    data_types = []
    try:
        with open(schema_file, 'r', encoding='utf-8') as f:
            lines = [line.strip() for line in f.readlines() if line.strip()]
        for i in range(0,len(lines), 3):
            table_name.append(lines[i])
            columns.append(lines[i+1].split())
            data_types.append(lines[i+2].split())
        
        
        return table_name, columns, data_types
    except Exception as e:
        print(f"Ошибка при чтении файла схемы: {e}")
        return None, None, None

def clean_numeric_columns(csv_file, columns, data_types):
    """Очищает числовые столбцы в CSV файле"""
    try:
        df = pd.read_csv(csv_file, index_col=None, dtype=str)
        df[" "] = "empty"
        df.columns = columns
        #print(df['F4'])

        # Создаем словарь {column: type} только для числовых столбцов
        numeric_cols = {
            col: dtype 
            for col, dtype in zip(columns, data_types) 
            if col in df.columns and dtype.lower() in ['int', 'integer', 'float', 'number', 'numeric', 'double', 'real']
        }
        for col, _ in numeric_cols.items():

            # Заменяем запятые на точки в числах
            df[col] = df[col].astype(str).apply(lambda x: re.sub(r',(\d+)', r'.\1', x))
            
            # Преобразуем в числовой тип, нечисловые значения станут NaN
            df[col] = pd.to_numeric(df[col], errors='coerce')
        
        # Сохраняем обратно в тот же файл
        df.to_csv(csv_file, index=False, encoding='utf-8')

        content = ''
        with open(csv_file, 'r', encoding='utf-8') as f_in:
            content = f_in.read()

        with open(csv_file, 'w', encoding='cp1251') as f_out:
                f_out.write(content)

        print(f"Числовые столбцы очищены в файле: {csv_file}")
        return True
    except Exception as e:
        print(f"Ошибка при обработке CSV файла: {e}")
        return False

def main():
    # Конфигурация
    # Получаем абсолютный путь к каталогу со скриптом
    script_dir = os.path.dirname(os.path.abspath(__file__))

    schema_file = os.path.join(script_dir, r"output\access_meta.txt")                               # Файл с описанием структуры
    csv_files_path = os.path.join(script_dir, r"output\\")                                          # Выходной CSV файл
    
    excel_files = [
        os.path.join(script_dir, r"input\0503124 — копия.xls"),                                     # список excel-файлов
        os.path.join(script_dir, r"input\0506604 — копия.xls"),
    ]
    sheet_names = [
        '2',                                                                                        # список листов которые импортируются
        'Лист1'
    ]


    # Формирование csv файлов для импорта в LO Base
    table_names, columns, data_types = parse_schema_file(schema_file)

    for i in range(len(excel_files)):
        csv_output = csv_files_path +"/" + table_names[i] + ".csv"
        if not process_excel_to_csv(excel_files[i], sheet_names[i], csv_output):
            print("Ошибка формирования csv")

        print(f"Обрабатывается таблица: {table_names[i]}")
        print(f"Столбцы: {columns[i]}")
        print(f"Типы данных: {data_types[i]}")

        clean_numeric_columns(csv_output, columns[i], data_types[i])

if __name__ == "__main__":
    main()