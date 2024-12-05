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
    
    columnIndex = 92 'Range("SQ4").Column ' �������� ����� �������, ��� ��������� ������
    
    'Set ws = ThisWorkbook.Sheets("����1")  ����� ���������� ���� ���� ���� ������� ��� ������ � ����� ws ����� ���� ������ ��� �� ������� �� ����������
    lastRow = Cells(Rows.Count, "H").End(xlUp).Row ' ������� ��������� ������

    ' ������� ��������� ��� �������� ����
    Set summary = CreateObject("Scripting.Dictionary")
    
    ' ���������� �� ������ ��� �������
    For Each dataRange In Range("H4:H" & lastRow) ' ������� � "��� �������"
        masterName = dataRange.value
        hours = Cells(dataRange.Row, columnIndex).value ' ������� � ��������� ���������
       
        ' ��������� ���� �� ��� �������
        If summary.exists(masterName) Then
            summary(masterName) = summary(masterName) + hours
        Else
            summary.Add masterName, hours
        End If
    Next dataRange

    ' ������� ����� ���� ��� ������
    On Error Resume Next
    Set reportWs = ThisWorkbook.Sheets("�����")
    On Error GoTo 0

    If reportWs Is Nothing Then
        Set reportWs = ThisWorkbook.Sheets.Add
        reportWs.name = "�����"
    Else
        reportWs.Cells.Clear
    End If

    ' ��������� ��� ������ �����
    reportWs.Range("A1").value = "��� �������"
    reportWs.Range("B1").value = "����� �����"

    reportLastRow = 2
    total = 0
    ' ��������� �������
    For Each masterName In summary.Keys
        reportWs.Cells(reportLastRow, 1).value = masterName
        reportWs.Cells(reportLastRow, 2).value = summary(masterName)
        total = total + summary(masterName)
        reportLastRow = reportLastRow + 1
    Next masterName

    ' �������� ������
    reportWs.Cells(reportLastRow, 1).value = "����� ����"
    reportWs.Cells(reportLastRow, 2).value = total

    reportWs.Columns("A:B").AutoFit

    MsgBox "����� ������� ������!", vbInformation
End Sub
