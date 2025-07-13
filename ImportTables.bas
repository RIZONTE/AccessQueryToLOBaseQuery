
Sub ImportTablesFromCsv()
	Dim sDirectory As String
	Dim csvFiles As Variant
	
	' Каталог из которого берутся CSV
	sDirectory = "file:///C:/Users/danch/Desktop/YaIMP/Python/AccessQueryToLOBaseQuery/output/"
	
	' Считывание файлов из каталога
	csvFiles = GetCsvFilesArray(sDirectory)
	
	' Импорт файлов в таблицы
	For i = LBound(csvFiles) To UBound(csvFiles)
            ImportCSVToTable(csvFiles(i))
    Next i
	
End Sub


Function GetCsvFilesArray(sDirectory As String) As Variant
    Dim sFile As String
    Dim iCount As Integer
    Dim aFiles() As String  ' Массив для хранения имен файлов
    
    ' Проверяем, что путь заканчивается на "/" или "\"
    If Right(sDirectory, 1) <> "/" And Right(sDirectory, 1) <> "\" Then
        sDirectory = sDirectory & "/"
    End If
    
    ' Ищем CSV-файлы
    sFile = Dir(sDirectory & "*.csv")
    iCount = 0
    
    ' Собираем файлы в массив
    While sFile <> ""
        ReDim Preserve aFiles(iCount)  ' Расширяем массив
        aFiles(iCount) = sDirectory & sFile  ' Добавляем файл в массив
        iCount = iCount + 1
        sFile = Dir()  ' Получаем следующий файл
    Wend
    
    GetCsvFilesArray = aFiles   ' Возвращаем массив с файлами
End Function

Sub ImportCSVToTable(sFilePath As String)
    Dim sTableName As String
    Dim oStatement As Object
    Dim oConnection As Object
    Dim oDocument As Object
    Dim oService As Object
    Dim oArgs(3) As New com.sun.star.beans.PropertyValue
    
    ' Получаем имя таблицы из имени файла (без расширения)
    sTableName = Left(right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "/")), _
                 InStrRev(right(sFilePath, Len(sFilePath) - InStrRev(sFilePath, "/")), ".") - 1)
    
    If sTableName = "" Then
        MsgBox "Не удалось определить имя таблицы из пути к файлу", vbExclamation
        Exit Sub
    End If
    
    ' Получаем текущий документ Base
    oDocument = ThisComponent
    
    If IsNull(oDocument) Then
        MsgBox "Нет открытого документа!", vbExclamation
        Exit Sub
    End If
    
    ' Получаем соединение с базой данных
    Dim oDBContext As Object
    Dim oDataSource As Object
    
    oDBContext = CreateUnoService("com.sun.star.sdb.DatabaseContext")
    oDataSource = oDBContext.getByName("Новая база данных4") ' Например, "Безымянная"
    oConnection = oDataSource.getConnection("", "") ' Логин/пароль, если требуется
    
    ' Проверяем, существует ли таблица
    If Not TableExists(oConnection, sTableName) Then
        MsgBox "Таблица " & sTableName & " не существует в базе данных", vbExclamation
        Exit Sub
    End If
    
    ' Чтение CSV файла
    Dim iFile As Integer, sLine As String, aValues() As String
    Dim sValues As String, sSQL As String, i As Integer, lRowCount As Long
    Dim lImported As Long, lFailed As Long
    Dim sDelimiter As String, sTextQualifier As String
    
    ' Параметры CSV (можно изменить)
    sDelimiter = ","     ' Разделитель полей
    sTextQualifier = """" ' Обрамляющий символ
    
    On Error Resume Next
    iFile = FreeFile()
    Open sFilePath For Input Access Read As #iFile
    
    ' Пропускаем заголовок (первую строку)
    Line Input #iFile, sLine
    lRowCount = 0
    
    ' Очистка перед вставкой данных
    sSQL = "DELETE FROM """ & sTableName & """"
    oStatement = oConnection.createStatement()
    oStatement.executeUpdate(sSQL)
    
    ' Основной цикл обработки
    Do Until EOF(iFile)
        Line Input #iFile, sLine
        lRowCount = lRowCount + 1
        
        ' Улучшенный парсинг строки CSV с учетом кавычек
        aValues = ParseCSVLine(sLine, sDelimiter, sTextQualifier)
        
        ' Формируем SQL запрос
        sValues = ""
        For i = LBound(aValues) To UBound(aValues)
            If sValues <> "" Then sValues = sValues & ", "
            
            ' Автоматическое определение типа значения
            Dim sValue As String
            sValue = Trim(aValues(i))
            
            ' Обработка специальных случаев
            If sValue = "" Then
                sValues = sValues & "NULL"
            'ElseIf IsNumeric(sValue) Then
                'sValues = sValues & sValue
            Else
                sValues = sValues & "'" & Replace(sValue, "'", "''") & "'"
            End If
        Next i
        
        ' Выполняем INSERT
        sSQL = "INSERT INTO """ & sTableName & """ VALUES (" & sValues & ")"
        
        On Error Resume Next
        oStatement = oConnection.createStatement()
        oStatement.executeUpdate(sSQL)
        If Err.Number = 0 Then
            lImported = lImported + 1
        Else
            lFailed = lFailed + 1
            Debug.Print "Ошибка в строке " & lRowCount & ": " & Err.Description
            Debug.Print "SQL: " & sSQL
            Err.Clear
        End If
        On Error GoTo 0
    Loop
    
    Close #iFile
    On Error GoTo 0
    
    MsgBox "Импорт завершен. Успешно: " & lImported & ", с ошибками: " & lFailed & vbCrLf & _
           "Данные импортированы в таблицу " & sTableName, vbInformation
End Sub

Function ParseCSVLine(sLine As String, sDelimiter As String, sTextQualifier As String) As Variant
    Dim aResults() As String
    Dim iPos As Integer, iStart As Integer, iEnd As Integer
    Dim bInQuotes As Boolean
    Dim sChar As String
    Dim i As Integer
    Dim sValue As String
    Dim iResultCount As Integer
    
    ReDim aResults(0)
    iResultCount = 0
    bInQuotes = False
    sValue = ""
    
    For i = 1 To Len(sLine)
        sChar = Mid(sLine, i, 1)
        
        If sChar = sTextQualifier Then
            ' Проверяем, не является ли это экранированной кавычкой
            If i < Len(sLine) And Mid(sLine, i + 1, 1) = sTextQualifier Then
                ' Это экранированная кавычка - добавляем одну кавычку
                sValue = sValue & sTextQualifier
                i = i + 1 ' Пропускаем следующую кавычку
            Else
                ' Это начало или конец кавычек
                bInQuotes = Not bInQuotes
            End If
        ElseIf sChar = sDelimiter And Not bInQuotes Then
            ' Нашли разделитель вне кавычек - добавляем значение в массив
            ReDim Preserve aResults(iResultCount)
            aResults(iResultCount) = sValue
            iResultCount = iResultCount + 1
            sValue = ""
        Else
            ' Обычный символ - добавляем к текущему значению
            sValue = sValue & sChar
        End If
    Next i
    
    ' Добавляем последнее значение
    ReDim Preserve aResults(iResultCount)
    aResults(iResultCount) = sValue
    
    ParseCSVLine = aResults
End Function

Function TableExists(oConnection As Object, sTableName As String) As Boolean
    Dim oTables As Object
    Dim sName As String
    
    oTables = oConnection.getTables()
    TableExists = oTables.hasByName(sTableName)
End Function

Function InStrRev(sText As String, sChar As String) As Integer
    If IsMissing(sText) Or IsMissing(sChar) Then
        InStrRev = 0
        Exit Function
    End If

    If sText = "" Or sChar = "" Then
        InStrRev = 0
        Exit Function
    End If

    Dim i As Integer
    For i = Len(sText) To 1 Step -1
        If Mid(sText, i, 1) = sChar Then
            InStrRev = i
            Exit Function
        End If
    Next i
    InStrRev = 0
End Function