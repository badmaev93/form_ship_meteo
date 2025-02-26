Attribute VB_Name = "ModuleMain"
Public Sub AddRecord()
    ' Вызываем форму для добавления записи
    Dim frm As UserForm1
    Set frm = New UserForm1
    frm.Tag = "New"
    frm.Show
End Sub

Public Sub EditRecord()
    On Error GoTo ErrorHandler
    
    ' Активируем лист Data перед запросом номера строки
    ThisWorkbook.Sheets("Data").Activate
    
    Dim rowNum As String
    rowNum = InputBox("Введите номер строки для редактирования:", "Редактирование записи")
    
    If rowNum = "" Then Exit Sub
    If Not IsNumeric(rowNum) Then
        MsgBox "Введите корректный номер строки!", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' Проверяем существование строки
    If CLng(rowNum) < 2 Or CLng(rowNum) > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "Строка с таким номером не существует!", vbExclamation
        Exit Sub
    End If
    
    ' Выделяем строку для наглядности
    ws.Rows(CLng(rowNum)).Select
    
    ' Создаем и показываем форму
    With New UserForm1
        .Caption = "Редактировать запись"
        .Tag = CStr(rowNum)
        .Show vbModal
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Ошибка при открытии формы: " & vbNewLine & Err.Description, vbCritical
End Sub



