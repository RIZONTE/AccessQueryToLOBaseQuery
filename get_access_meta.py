import win32com.client

def get_access_metadata(access_db_path, output_file):
    # Создаем подключение к Access
    access = win32com.client.Dispatch("Access.Application")
    db = access.DBEngine.OpenDatabase(access_db_path)
    
    with open(output_file, 'w', encoding='utf-8') as f:
        # Получаем все таблицы
        tables = db.TableDefs
        for table in tables:
            # Пропускаем системные таблицы
            if table.Name.startswith('MSys') or table.Name.startswith('~'):
                continue
                
            # Записываем имя таблицы
            f.write(f"{table.Name}\n")
            
            # Получаем и записываем имена полей
            fields = table.Fields
            field_names = [field.Name for field in fields]
            f.write(" ".join(field_names) + "\n")
            
            # Получаем и записываем типы данных полей
            field_types = []
            for field in fields:
                # Преобразуем числовой тип в строковое представление
                type_name = get_field_type_name(field.Type)
                field_types.append(type_name)
            f.write(" ".join(field_types) + "\n\n")
    
    db.Close()
    access.Quit()

def get_field_type_name(field_type):
    # Сопоставление типов данных Access с их строковыми представлениями
    type_map = {
        1: "Boolean",
        2: "Byte",
        3: "Integer",
        4: "Long",
        5: "Currency",
        6: "Single",
        7: "Double",
        8: "DateTime",
        9: "Binary",
        10: "Text",
        11: "LongBinary",
        12: "Memo",
        15: "GUID",
        16: "BigInt",
        17: "VarBinary",
        18: "Char",
        19: "Numeric",
        20: "Decimal",
        21: "Float",
        22: "Time",
        23: "TimeStamp",
    }
    return type_map.get(field_type, f"Unknown({field_type})")
