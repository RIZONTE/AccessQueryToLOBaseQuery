import uno
from com.sun.star.beans import PropertyValue
from com.sun.star.uno import Exception as UnoException

def test_libreoffice_base_connection(odb_path):
    """
    Тестирование подключения к файлу LibreOffice Base (.odb)
    
    :param odb_path: Путь к файлу базы данных (например: 'C:/path/to/database.odb')
    """
    print("\n=== Тест подключения к LibreOffice Base ===")
    
    try:
        # 1. Получаем контекст компонентов
        local_context = uno.getComponentContext()
        print("✓ Контекст компонентов получен")
        
        # 2. Создаем резолвер для подключения
        resolver = local_context.ServiceManager.createInstanceWithContext(
            "com.sun.star.bridge.UnoUrlResolver", local_context)
        print("✓ Резолвер создан")
        
        # 3. Подключаемся к запущенному LibreOffice
        connection_string = "uno:socket,host=localhost,port=2002;urp;StarOffice.ComponentContext"
        context = resolver.resolve(connection_string)
        print("✓ Подключение к LibreOffice установлено")
        
        # 4. Получаем объект Desktop
        desktop = context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", context)
        print("✓ Desktop объект получен")
        
        # 5. Открываем базу данных
        db_url = uno.systemPathToFileUrl(odb_path)
        db_properties = (
            PropertyValue(Name="Hidden", Value=True),
            PropertyValue(Name="ReadOnly", Value=False),
        )
        
        print(f"\nПопытка открыть базу данных: {odb_path}")
        database = desktop.loadComponentFromURL(
            db_url, "_blank", 0, db_properties)
        print("✓ База данных успешно открыта")
        
        # Проверка доступных сервисов
        db_context = desktop.getCurrentComponent().getDatabaseContext()
        print("Доступные базы:", [name for name in db_context.getElementNames()])

        # Проверка состояния базы
        database = desktop.getCurrentComponent()
        print("Документ загружен?", database.isLoaded())
        print("Тип документа:", database.ImplementationName)
        
        # 6. Проверяем подключение к данным
        connection = database.getConnection("", "")
        print("✓ Подключение к данным установлено")
        
        # 7. Получаем список таблиц
        tables = database.getTables()
        table_names = tables.getElementNames()
        print(f"\nНайдено таблиц: {len(table_names)}")
        print("Первые 5 таблиц:" if len(table_names) > 5 else "Таблицы:")
        for name in table_names[:5]:
            print(f" - {name}")
        
        # 8. Получаем список запросов
        queries = database.getQueryDefinitions()
        query_names = queries.getElementNames()
        print(f"\nНайдено запросов: {len(query_names)}")
        print("Первые 5 запросов:" if len(query_names) > 5 else "Запросы:")
        for name in query_names[:5]:
            print(f" - {name}")
        
        # 9. Закрываем соединение
        database.close(True)
        print("\n✓ Соединение закрыто")
        print("=== Тест пройден успешно ===")
        return True
        
    except UnoException as e:
        print(f"\n!!! Ошибка UNO: {e.Message}")
        if "Connection refused" in e.Message:
            print("Возможно, LibreOffice не запущен в серверном режиме")
            print("Запустите LibreOffice командой:")
            print('soffice.exe --headless --accept="socket,host=localhost,port=2002;urp;StarOffice.ServiceManager"')
        return False
    except Exception as e:
        print(f"\n!!! Общая ошибка: {str(e)}")
        return False

if __name__ == "__main__":
    # Укажите путь к вашему файлу .odb
    database_path = "C:/Users/danch/Desktop/Практика_07-2025/job1/604 со 124/Новая база данных3.odb"  # Замените на реальный путь
    
    # Запуск теста
    success = test_libreoffice_base_connection(database_path)
