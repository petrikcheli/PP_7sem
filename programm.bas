Attribute VB_Name = "Module3"
Sub SummarizeByMaster()
    Dim ws As Worksheet, reportWs As Worksheet
    Dim lastRow As Long, reportLastRow As Long
    Dim dataRange As Range
    Dim summary As Object
    Dim masterName As Variant
    Dim hours As Double
    Dim total As Double
    Dim columnIndex As Long
    
    columnIndex = 92 'Range("SQ4").Column ' Получаем номер столбца, где находятся данные
    
    'Set ws = ThisWorkbook.Sheets("Лист1")  можно подставить свой лист либо удалить эту строку и везде ws чтобы лист брался тот на котором вы находитесь
    lastRow = Cells(Rows.Count, "H").End(xlUp).Row ' Находим последнюю строку

    ' Создаем коллекцию для хранения сумм
    Set summary = CreateObject("Scripting.Dictionary")
    
    ' Проходимся по нужным нам строкам
    For Each dataRange In Range("H4:H" & lastRow) ' Столбец с "ФИО Мастера"
        masterName = dataRange.value
        hours = Cells(dataRange.Row, columnIndex).value ' Столбец с последним значением
       
        ' суммируем часы по ФИО Мастера
        If summary.exists(masterName) Then
            summary(masterName) = summary(masterName) + hours
        Else
            summary.Add masterName, hours
        End If
    Next dataRange

    ' создаем новый лист для отчета
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("Отчет")
    On Error GoTo 0

    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.name = "Отчет"
    Else
        reportWs.Cells.Clear
    End If

    ' заголовки для нового листа
    reportWs.Range("A1").value = "ФИО Мастера"
    reportWs.Range("B1").value = "Сумма часов"

    reportLastRow = 2
    total = 0
    ' заполняем таблицу
    For Each masterName In summary.Keys
        reportWs.Cells(reportLastRow, 1).value = masterName
        reportWs.Cells(reportLastRow, 2).value = summary(masterName)
        total = total + summary(masterName)
        reportLastRow = reportLastRow + 1
    Next masterName

    ' Итоговая строка
    reportWs.Cells(reportLastRow, 1).value = "Общий итог"
    reportWs.Cells(reportLastRow, 2).value = total

    reportWs.Columns("A:B").AutoFit

    MsgBox "Отчет успешно создан!", vbInformation
End Sub
