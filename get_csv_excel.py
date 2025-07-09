import pandas as pd
import re

def process_excel_to_csv(excel_file, sheet_name, output_csv):
    """Читает указанный лист из Excel файла и сохраняет в CSV"""
    try:
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
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
        df = pd.read_csv(csv_file, index_col=None)
        df.columns = columns[0:-1]

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
        print(f"Числовые столбцы очищены в файле: {csv_file}")
        return True
    except Exception as e:
        print(f"Ошибка при обработке CSV файла: {e}")
        return False

def main():
    # Конфигурация (можно заменить на аргументы командной строки)
    excel_file = "C:/Users/danch/Desktop/Практика_07-2025/job1/604 со 124/0503124 — копия.xls"      # Путь к исходному Excel файлу
    sheet_name = '2'                                                                                # Имя листа для обработки
    schema_file = "C:/Users/danch/Desktop/Практика_07-2025/access_meta.txt"                         # Файл с описанием структуры
    output_csv = 'C:/Users/danch/Desktop/Практика_07-2025/output124.csv'                            # Выходной CSV файл
    
    # 1. Конвертируем Excel в CSV
    if not process_excel_to_csv(excel_file, sheet_name, output_csv):
        return
    
    # 2. Читаем файл схемы
    table_name, columns, data_types = parse_schema_file(schema_file)
    if not table_name:
        return
    
    print(f"Обрабатывается таблица: {table_name}")
    print(f"Столбцы: {columns}")
    print(f"Типы данных: {data_types}")
    # 3. Очищаем числовые столбцы
    clean_numeric_columns(output_csv, columns, data_types)

if __name__ == "__main__":
    main()