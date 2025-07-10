import win32com.client
import re

def convert_access_to_base(access_query):
    # замена всех квадратных скобок на кавычки
    access_query = access_query.replace('[', '\"')
    access_query = access_query.replace(']', '\"')

    # замена ! на .
    access_query = access_query.replace('!', '.')

    # замена конкатенации
    access_query = access_query.replace('&', '||')

    # Замена дат #YYYY-MM-DD# → 'YYYY-MM-DD'
    access_query = re.sub(r'#(\d{4}-\d{2}-\d{2})#', r"'\1'", access_query)
    
    # Замена * и ? в LIKE на % и _
    access_query = re.sub(r'(?i)LIKE\s+[\'"]([^\'"]*)\*([^\'"]*)[\'"]', r'LIKE \'%\2\'', access_query)
    access_query = re.sub(r'(?i)LIKE\s+[\'"]([^\'"]*)\?([^\'"]*)[\'"]', r'LIKE \'_\2\'', access_query)

    # Замена IIF(условие, да, нет) → CASE WHEN условие THEN да ELSE нет END
    
    access_query = re.sub(
        r'IIF\(([^,]+),\s*([^,]+),\s*([^)]+)\)', 
        r'CASE WHEN \1 THEN \2 ELSE \3 END', 
        access_query, 
        flags=re.IGNORECASE
    )
    
    # Замена TOP N → LIMIT N (если не в подзапросе)
    if "LIMIT" not in access_query.upper():
        access_query = re.sub(
            r'SELECT\s+(TOP\s+\d+\s+)(.*?)\s+FROM', 
            r'SELECT \2 LIMIT \1', 
            access_query, 
            flags=re.IGNORECASE
        )
        access_query = access_query.replace("TOP", "").strip()

    print(f"Сконвертированная строка: {access_query}")
    return access_query

def export_access_queries(access_db_path, output_file):
    """Извлекает запросы из MS Access, конвертирует их в запросы LO Base и записывает в файл"""

    dao_db = win32com.client.Dispatch("DAO.DBEngine.120").OpenDatabase(access_db_path)
    try:
        with open(output_file, 'w', encoding='utf-8') as f:

            # Получаем все запросы через QueryDefs
            for query_def in dao_db.QueryDefs:
                f.write(f"{query_def.Name}\n")
                converted = convert_access_to_base(query_def.SQL)
                f.write(f"{converted.replace("\r\n", ' ')}\n\n")
        print(f"Все запросы сохранены в файл: {output_file}")
        dao_db.Close()
        return True
    except:
        dao_db.Close()
        return False
    

