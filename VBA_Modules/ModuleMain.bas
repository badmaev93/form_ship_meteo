Attribute VB_Name = "ModuleMain"
Public Sub AddRecord()
    ' �������� ����� ��� ���������� ������
    Dim frm As UserForm1
    Set frm = New UserForm1
    frm.Tag = "New"
    frm.Show
End Sub

Public Sub EditRecord()
    On Error GoTo ErrorHandler
    
    ' ���������� ���� Data ����� �������� ������ ������
    ThisWorkbook.Sheets("Data").Activate
    
    Dim rowNum As String
    rowNum = InputBox("������� ����� ������ ��� ��������������:", "�������������� ������")
    
    If rowNum = "" Then Exit Sub
    If Not IsNumeric(rowNum) Then
        MsgBox "������� ���������� ����� ������!", vbExclamation
        Exit Sub
    End If
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' ��������� ������������� ������
    If CLng(rowNum) < 2 Or CLng(rowNum) > ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Then
        MsgBox "������ � ����� ������� �� ����������!", vbExclamation
        Exit Sub
    End If
    
    ' �������� ������ ��� �����������
    ws.Rows(CLng(rowNum)).Select
    
    ' ������� � ���������� �����
    With New UserForm1
        .Caption = "������������� ������"
        .Tag = CStr(rowNum)
        .Show vbModal
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "������ ��� �������� �����: " & vbNewLine & Err.Description, vbCritical
End Sub



