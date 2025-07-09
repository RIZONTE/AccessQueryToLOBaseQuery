import win32com.client

def convert_access_to_base(access_query):
    # замена всех квадратных скобок на кавычки
    access_query = access_query.replace('[', '\"')
    access_query = access_query.replace(']', '\"')

    # замена ! на .
    access_query = access_query.replace('!', '.')

    # замена конкатенации
    access_query = access_query.replace('&', '||')

    print(f"Сконвертированная строка: {access_query}")
    return access_query

def export_access_queries(access_db_path, output_file):
    """Извлекает запросы из MS Access, конвертирует их в запросы LO Base и записывает в файл"""

    dao_db = win32com.client.Dispatch("DAO.DBEngine.120").OpenDatabase(access_db_path)
    try:
        with open(output_file, 'w', encoding='utf-8') as f:

            # Получаем все запросы через QueryDefs
            for query_def in dao_db.QueryDefs:
                f.write(f"=== Запрос: {query_def.Name} ===\n")
                converted = convert_access_to_base(query_def.SQL)
                f.write(f"{converted}\n\n")
        print(f"Все запросы сохранены в файл: {output_file}")
        dao_db.Close()
        return True
    except:
        dao_db.Close()
        return False
    

