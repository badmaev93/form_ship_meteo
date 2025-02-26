VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ёта нига"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    On Error Resume Next
    
    ' —начала попробуем сн¤ть защиту, если она уже установлена
    ThisWorkbook.Sheets("Data").Unprotect Password:="yourpassword" ' ”кажите ваш пароль здесь
    
    ' “еперь устанавливаем защиту с паролем (дл¤ работы с VBA оставл¤ем UserInterfaceOnly:=True)
    ThisWorkbook.Sheets("Data").Protect Password:="3timitimi3", UserInterfaceOnly:=True ' ”кажите тот же пароль
    
    On Error GoTo 0
End Sub
