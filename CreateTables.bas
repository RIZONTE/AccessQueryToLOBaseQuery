REM  *****  BASIC  *****

Sub CreateTablesFromFile()
    Dim sFilePath As String
    Dim iFile As Integer
    Dim sFileContent As String
    Dim sLine As String
    Dim aTables() As String
    Dim i As Integer
    Dim oConn As Object
    
    ' Укажите путь к вашему текстовому файлу
    sFilePath = "C:\Users\danch\Desktop\Практика_07-2025\access_meta.txt"
    
    ' Соединение с бд
    oConn = ThisComponent.getCurrentController().ActiveConnection
    
    
    ' Открываем файл для чтения
    iFile = FreeFile()
    On Error GoTo FileError
    Open sFilePath For Input As #iFile
    On Error GoTo 0
    
    ' Читаем содержимое файла
    sFileContent = ""
    Do Until EOF(iFile)
        Line Input #iFile, sLine
        sFileContent = sFileContent & sLine & Chr(10)
    Loop
    Close #iFile
    
    ' Разделяем на таблицы (разделитель - пустая строка)
    aTables = Split(sFileContent, Chr(10) & Chr(10))
    
    ' Обрабатываем каждую таблицу
    For i = LBound(aTables) To UBound(aTables)
        If Trim(aTables(i)) <> "" Then
            CreateTable(oConn, aTables(i))
        End If
    Next i
    
    MsgBox "Таблицы успешно созданы!", vbInformation
    Exit Sub
    
FileError:
    MsgBox "Ошибка при открытии файла: " & sFilePath, vbExclamation
    Exit Sub
End Sub

Sub CreateTable(oConn As Object, sTableDef As String)
    Dim aLines() As String
    Dim sTableName As String
    Dim aFields() As String
    Dim aTypes() As String
    Dim sSQL As String
    Dim i As Integer
    Dim oStatement As Object
    
    ' Разделяем определение таблицы на строки
    aLines = Split(sTableDef, Chr(10))
    
    ' Проверяем, что есть достаточно строк
    If UBound(aLines) < 2 Then
        MsgBox "Неверный формат определения таблицы: " & sTableDef, vbExclamation
        Exit Sub
    End If
    
    ' Первая строка - имя таблицы
    sTableName = Trim(aLines(0))
    
    ' Вторая строка - имена полей
    aFields = Split(Trim(aLines(1)), " ")
    
    ' Третья строка - типы данных
    aTypes = Split(Trim(aLines(2)), " ")
    
    ' Проверяем, что количество полей и типов совпадает
    If UBound(aFields) <> UBound(aTypes) Then
        MsgBox "Ошибка в определении таблицы " & sTableName & ": количество полей и типов не совпадает", vbExclamation
        Exit Sub
    End If
    
    ' Формируем SQL запрос для создания таблицы
    sSQL = "CREATE TABLE """ & sTableName & """ ("
    
    For i = LBound(aFields) To UBound(aFields)
        If aFields(i) <> "" Then
            sSQL = sSQL & """" & aFields(i) & """ " & ConvertType(aTypes(i))
            If i < UBound(aFields) Then sSQL = sSQL & ", "
        End If
    Next i
    
    sSQL = sSQL & ")"
    
    
    ' Выполняем SQL запрос
    On Error Resume Next
    Set oStatement = oConn.createStatement()
    oStatement.executeUpdate(sSQL)
    If Err <> 0 Then
        MsgBox "Ошибка при создании таблицы " & sTableName & ": " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0
End Sub

Function ConvertType(sType As String) As String
    ' Конвертируем типы данных в соответствующие для Base
    Select Case LCase(Trim(sType))
        Case "int", "integer"
            ConvertType = "INTEGER"
        Case "text", "string", "varchar"
            ConvertType = "VARCHAR(255)"
        Case "date"
            ConvertType = "DATE"
        Case "datetime", "timestamp"
            ConvertType = "TIMESTAMP"
        Case "boolean", "bool"
            ConvertType = "BOOLEAN"
        Case "float", "double", "real"
            ConvertType = "DOUBLE"
        Case "blob", "binary"
            ConvertType = "BLOB"
        Case Else
            ConvertType = sType ' Используем как есть, если тип не распознан
    End Select
End Function