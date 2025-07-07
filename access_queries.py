import win32com.client

def export_access_queries(access_db_path, output_file):
    
    dao_db = win32com.client.Dispatch("DAO.DBEngine.120").OpenDatabase(access_db_path)
    
    with open(output_file, 'w', encoding='utf-8') as f:

        # Получаем все запросы через QueryDefs
        for query_def in dao_db.QueryDefs:
            f.write(f"=== Запрос: {query_def.Name} ===\n")
            f.write(f"{query_def.SQL}\n\n")
    
    dao_db.Close()
    print(f"Все запросы сохранены в файл: {output_file}")
    

export_access_queries('C:/Users/danch/Desktop/Практика_07-2025/Сверка.accdb', 
                     'C:/Users/danch/Desktop/Практика_07-2025/queries_output.txt')
