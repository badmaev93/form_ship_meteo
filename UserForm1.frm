VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Add Entry / �������� ������"
   ClientHeight    =   6510
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   13095
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" _
    (ByVal hwnd As LongPtr, ByVal wMsg As Long, _
     ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr

Private Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" _
    (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, _
     ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr

Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const CB_SHOWDROPDOWN As Long = &H14F

' ��������� ��� ���������
Private Const COORD_FORMAT_DECIMAL As Boolean = False
Private Const COORD_FORMAT_DEGREES As Boolean = True

' ��� ��� ������ � �����������
Private Type CoordInput
    degrees As MSForms.TextBox
    minutes As MSForms.TextBox
    direction As MSForms.ComboBox
End Type
Private Type ReasonData
    Field As String
    Reason As String
    DateTime As Date
End Type
' ���������� ��� �������� ��������� �����
Private mCoordFormat As Boolean
Private mIsCalm As Boolean
Private mIsPort As Boolean
Private LatitudeInput As CoordInput
Private LongitudeInput As CoordInput
Private mIsIceNotated As Boolean  ' ���� ��� Sea Ice Conditions noted

Private Sub lblLatitude_Click()

End Sub
Private Sub UserForm_Activate()
    Debug.Print "Form Activated. Tag = " & Me.Tag
    
    If IsNumeric(Me.Tag) And Me.Tag <> "New" Then
        LoadExistingData CLng(Me.Tag)
    End If
End Sub
Private Sub UserForm_Initialize()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Starting UserForm_Initialize ==="
    Debug.Print "Form Tag: " & Me.Tag
    
    ' ������������� ������� ���������
    InitializeCoordinateFields
    InitializeControls
    InitializeIceControls
    
    ' ��������� ��������
    mCoordFormat = COORD_FORMAT_DEGREES
    UpdateCoordinateControls
    
    If Me.Tag = "" Then
        Debug.Print "Empty tag - setting to New"
        Me.Tag = "New"
        SetDefaultValues
    End If
    
    Debug.Print "=== UserForm_Initialize completed ==="
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in UserForm_Initialize: " & Err.Description
    MsgBox "������ ������������� �����: " & vbNewLine & Err.Description, vbCritical
End Sub


Private Sub LoadExistingData(ByVal rowNum As Long)
    On Error GoTo ErrorHandler
    
    Debug.Print "=== Starting LoadExistingData for row " & rowNum & " ==="
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    If ws Is Nothing Then
        Debug.Print "ERROR: Data sheet not found"
        Exit Sub
    End If
    
    If ws.Cells(rowNum, 1).value = "" Then
        Debug.Print "ERROR: Row " & rowNum & " is empty"
        Exit Sub
    End If
    
    ' �������� �������� ����� ���������
    Debug.Print "Reading values from row " & rowNum & ":"
    Debug.Print "Date/Time: " & ws.Cells(rowNum, 1).value
    Debug.Print "Latitude: " & ws.Cells(rowNum, 2).value
    Debug.Print "Longitude: " & ws.Cells(rowNum, 3).value
    
    With Me
        ' ������� ��� ���� ����� ���������
        ClearAllFields
        
        ' Date/Time
        .txtDateTime1.value = Format(ws.Cells(rowNum, 1).value, "dd.mm.yyyy hh:00")
        
        ' Coordinates
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            .fraMain.fraCoordinates.txtLatitude.Text = FormatNumber(ws.Cells(rowNum, 2).value, 4)
            .fraMain.fraCoordinates.txtLongitude.Text = FormatNumber(ws.Cells(rowNum, 3).value, 4)
        Else
            ConvertToDegreesMinutes CDbl(ws.Cells(rowNum, 2).value), _
                                  .fraMain.fraCoordinates.txtLatDegrees, _
                                  .fraMain.fraCoordinates.txtLatMinutes, _
                                  .fraMain.fraCoordinates.cboLatDirection, _
                                  True
                                  
            ConvertToDegreesMinutes CDbl(ws.Cells(rowNum, 3).value), _
                                  .fraMain.fraCoordinates.txtLonDegrees, _
                                  .fraMain.fraCoordinates.txtLonMinutes, _
                                  .fraMain.fraCoordinates.cboLonDirection, _
                                  False
        End If
        
        ' ��������� ����
        .txtTemp.Text = ws.Cells(rowNum, 4).Text
        .txtBarometer.Text = ws.Cells(rowNum, 5).Text
        .txtVisibility.Text = ws.Cells(rowNum, 6).Text
        .txtWindDirection.Text = ws.Cells(rowNum, 7).Text
        .txtWindSpeed.Text = ws.Cells(rowNum, 8).Text
        .txtSeaSwellDirection.Text = ws.Cells(rowNum, 9).Text
        .txtSeaSwell.Text = ws.Cells(rowNum, 10).Text
        .txtWindWaveDirection.Text = ws.Cells(rowNum, 11).Text
        .txtWindWaveHeight.Text = ws.Cells(rowNum, 12).Text
        
        ' Ice data
        If ws.Cells(rowNum, 13).Text = "������ ����" Then
            .chkIceNotated.value = False
        Else
            .chkIceNotated.value = True
            FindAndSelectComboValue .cboIceScore, ws.Cells(rowNum, 13).Text
            FindAndSelectComboValue .cboIceType, ws.Cells(rowNum, 14).Text
            FindAndSelectComboValue .cboIceShape, ws.Cells(rowNum, 15).Text
        End If
        
        ' Update controls
        UpdateSeaControls
        UpdateCoordinateControls
        
        Debug.Print "Data loaded successfully"
    End With
    
    Exit Sub

ErrorHandler:
    Debug.Print "ERROR in LoadExistingData: " & Err.Description
    Debug.Print "Error Line: " & Erl
    Resume Next
End Sub








Private Sub InitializeCoordinateFields()
    With Me.fraMain.fraCoordinates
        '�������� ����� ���������
        Set LatitudeInput.degrees = .txtLatDegrees
        Set LatitudeInput.minutes = .txtLatMinutes
        Set LatitudeInput.direction = .cboLatDirection
        
        Set LongitudeInput.degrees = .txtLonDegrees
        Set LongitudeInput.minutes = .txtLonMinutes
        Set LongitudeInput.direction = .cboLonDirection
    End With
End Sub
'����������� ������
' ��� Decimal Degrees
Private Sub fraMain_fraCoordinates_txtLatitude_Click()
    If Not Me.fraMain.fraCoordinates.txtLatitude.Enabled Then
        Me.fraMain.fraCoordinates.optDecimalCoords.value = True
    End If
End Sub
Private Sub fraMain_fraCoordinates_txtLongitude_Click()
    If Not Me.fraMain.fraCoordinates.txtLongitude.Enabled Then
        Me.fraMain.fraCoordinates.optDecimalCoords.value = True
    End If
End Sub
' ��� Latitude Degrees/Minutes
Private Sub fraMain_fraCoordinates_txtLatDegrees_Click()
    If Not Me.fraMain.fraCoordinates.txtLatDegrees.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
Private Sub fraMain_fraCoordinates_txtLatMinutes_Click()
    If Not Me.fraMain.fraCoordinates.txtLatMinutes.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
Private Sub fraMain_fraCoordinates_cboLatDirection_Click()
    If Not Me.fraMain.fraCoordinates.cboLatDirection.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
' ��� Longitude Degrees/Minutes
Private Sub fraMain_fraCoordinates_txtLonDegrees_Click()
    If Not Me.fraMain.fraCoordinates.txtLonDegrees.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
Private Sub fraMain_fraCoordinates_txtLonMinutes_Click()
    If Not Me.fraMain.fraCoordinates.txtLonMinutes.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
Private Sub fraMain_fraCoordinates_cboLonDirection_Click()
    If Not Me.fraMain.fraCoordinates.cboLonDirection.Enabled Then
        Me.fraMain.fraCoordinates.optDegreeCoords.value = True
    End If
End Sub
Private Sub InitializeControls()
    On Error GoTo ErrorHandler
    
    With Me
        ' ���. ��������
        .optDecimalCoords.value = False
        .optDegreeCoords.value = True
        
        ' ������� �����
        .txtLongitude.Text = ""
        .txtLatitude.Text = ""
        .txtTemp.Text = ""
        .txtBarometer.Text = ""
        .txtWindDirection.Text = ""
        .txtWindSpeed.Text = ""
        .txtVisibility.Text = ""
        .txtSeaSwell.Text = ""
        .txtSeaSwellDirection.Text = ""
        .txtWindWaveDirection.Text = ""
        .txtWindWaveHeight.Text = ""
        
        ' ����. �����������
        InitializeIceControls
        InitializeDirectionControls
        
        ' ���. ���������
        .chkIceNotated = False
        .chkSeaSwell.value = True
                
    End With
    
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitializeControls: " & Err.Description
    Err.Raise Err.Number, "InitializeControls", _
              "Oshibka inicializacii elementov upravleniya."
End Sub
Private Sub InitializeDirectionControls()
    ' ����� ��������� ����.
    With LatitudeInput.direction
        .Clear
        .AddItem "N"
        .AddItem "S"
        .Text = "N"
    End With
    
    With LongitudeInput.direction
        .Clear
        .AddItem "E"
        .AddItem "W"
        .Text = "E"
    End With
End Sub
Private Function FormControlsExist() As Boolean
    On Error Resume Next
    Dim controlExists As Boolean
    Dim msg As String
    
    controlExists = True
    msg = ""
    
    With Me
        ' ����. ����. �����.
        If .optDecimalCoords Is Nothing Then
            msg = msg & "optDecimalCoords" & vbNewLine
            controlExists = False
        End If
        If .optDegreeCoords Is Nothing Then
            msg = msg & "optDegreeCoords" & vbNewLine
            controlExists = False
        End If
        If .txtLongitude Is Nothing Then
            msg = msg & "txtLongitude" & vbNewLine
            controlExists = False
        End If
        If .txtLatitude Is Nothing Then
            msg = msg & "txtLatitude" & vbNewLine
            controlExists = False
        End If
        If .fraLatitude Is Nothing Then
            msg = msg & "fraLatitude" & vbNewLine
            controlExists = False
        End If
        If .fraLongitude Is Nothing Then
            msg = msg & "fraLongitude" & vbNewLine
            controlExists = False
        End If
        
        ' ����. ����. � frame Latitude
        If .fraLatitude.Controls("txtLatDegrees") Is Nothing Then
            msg = msg & "txtLatDegrees" & vbNewLine
            controlExists = False
        End If
        If .fraLatitude.Controls("txtLatMinutes") Is Nothing Then
            msg = msg & "txtLatMinutes" & vbNewLine
            controlExists = False
        End If
        If .fraLatitude.Controls("cboLatDirection") Is Nothing Then
            msg = msg & "cboLatDirection" & vbNewLine
            controlExists = False
        End If
        
        ' � Longitude
        If .fraLongitude.Controls("txtLonDegrees") Is Nothing Then
            msg = msg & "txtLonDegrees" & vbNewLine
            controlExists = False
        End If
        If .fraLongitude.Controls("txtLonMinutes") Is Nothing Then
            msg = msg & "txtLonMinutes" & vbNewLine
            controlExists = False
        End If
        If .fraLongitude.Controls("cboLonDirection") Is Nothing Then
            msg = msg & "cboLonDirection" & vbNewLine
            controlExists = False
        End If
        
        ' ����. ���. �����
        If .txtDateTime1 Is Nothing Then
            msg = msg & "txtDateTime1" & vbNewLine
            controlExists = False
        End If
        If .txtTemp Is Nothing Then
            msg = msg & "txtTemp" & vbNewLine
            controlExists = False
        End If
        If .txtBarometer Is Nothing Then
            msg = msg & "txtBarometer" & vbNewLine
            controlExists = False
        End If
        If .txtVisibility Is Nothing Then
            msg = msg & "txtVisibility" & vbNewLine
            controlExists = False
        End If
        
        '����. ����. �����
        If .txtWindDirection Is Nothing Then
            msg = msg & "txtWindDirection" & vbNewLine
            controlExists = False
        End If
        If .txtWindSpeed Is Nothing Then
            msg = msg & "txtWindSpeed" & vbNewLine
            controlExists = False
        End If
        If .chkCalm Is Nothing Then
            msg = msg & "chkCalm" & vbNewLine
            controlExists = False
        End If
        
        ' ����. ���� ����.����
        If .chkSeaSwell Is Nothing Then
            msg = msg & "chkSeaSwell" & vbNewLine
            controlExists = False
        End If
        If .chkPort Is Nothing Then
            msg = msg & "chkPort" & vbNewLine
            controlExists = False
        End If
        If .txtSeaSwell Is Nothing Then
            msg = msg & "txtSeaSwell" & vbNewLine
            controlExists = False
        End If
        If .txtSeaSwellDirection Is Nothing Then
            msg = msg & "txtSeaSwellDirection" & vbNewLine
            controlExists = False
        End If
        If .txtWindWaveDirection Is Nothing Then
            msg = msg & "txtWindWaveDirection" & vbNewLine
            controlExists = False
        End If
        If .txtWindWaveHeight Is Nothing Then
            msg = msg & "txtWindWaveHeight" & vbNewLine
            controlExists = False
        End If
        
        ' ����. ��������� ����
        If .cboIceType Is Nothing Then
            msg = msg & "cboIceType" & vbNewLine
            controlExists = False
        End If
        If .cboIceScore Is Nothing Then
            msg = msg & "cboIceScore" & vbNewLine
            controlExists = False
        End If
        
        ' ���� ����� labels
        If .lblSeaSwell Is Nothing Then
            msg = msg & "lblSeaSwell" & vbNewLine
            controlExists = False
        End If
        If .lblSeaSwellDirection Is Nothing Then
            msg = msg & "lblSeaSwellDirection" & vbNewLine
            controlExists = False
        End If
        If .lblWindWaveDirection Is Nothing Then
            msg = msg & "lblWindWaveDirection" & vbNewLine
            controlExists = False
        End If
        If .lblWindWaveHeight Is Nothing Then
            msg = msg & "lblWindWaveHeight" & vbNewLine
            controlExists = False
        End If
        If .lblIceType Is Nothing Then
            msg = msg & "lblIceType" & vbNewLine
            controlExists = False
        End If
        If .lblIceScore Is Nothing Then
            msg = msg & "lblIceScore" & vbNewLine
            controlExists = False
        End If
        
        '����. ������
        If .cmdSave Is Nothing Then
            msg = msg & "cmdSave" & vbNewLine
            controlExists = False
        End If
        If .cmdCancel Is Nothing Then
            msg = msg & "cmdCancel" & vbNewLine
            controlExists = False
        End If
    End With
    
    If Not controlExists Then
        MsgBox "Otsutstvuyut elementy:" & vbNewLine & msg, vbCritical
    End If
    
    FormControlsExist = controlExists
End Function
Private Sub InitCoordinateFields()
    On Error GoTo ErrorHandler
    
    ' ���� ������
    With Me.fraLatitude
        Set LatitudeInput.degrees = .Controls("txtLatDegrees")
        Set LatitudeInput.minutes = .Controls("txtLatMinutes")
        Set LatitudeInput.direction = .Controls("cboLatDirection")
    End With
    
    ' �������
    With Me.fraLongitude
        Set LongitudeInput.degrees = .Controls("txtLonDegrees")
        Set LongitudeInput.minutes = .Controls("txtLonMinutes")
        Set LongitudeInput.direction = .Controls("cboLonDirection")
    End With
    
    ' ����� ��������� �����������
    With LatitudeInput.direction
        .Clear
        .AddItem "N"
        .AddItem "S"
        .Text = "N"
    End With
    
    With LongitudeInput.direction
        .Clear
        .AddItem "E"
        .AddItem "W"
        .Text = "E"
    End With
    
    '������� ����� �����
    LatitudeInput.degrees.Text = ""
    LatitudeInput.minutes.Text = ""
    LongitudeInput.degrees.Text = ""
    LongitudeInput.minutes.Text = ""
    
    Exit Sub

ErrorHandler:
    Debug.Print "Error in InitCoordinateFields: " & Err.Description
    Err.Raise Err.Number, "InitCoordinateFields", "Oshibka inicializacii poley koordinat: " & Err.Description
End Sub
' ���� ������� ����
Private Sub InitializeIceControls()
    On Error GoTo ErrorHandler
    
    Dim wsIceScore As Worksheet
    Dim wsIceType As Worksheet
    Dim wsIceShape As Worksheet
    Set wsIceScore = ThisWorkbook.Sheets("IceScore")
    Set wsIceType = ThisWorkbook.Sheets("IceType")
    Set wsIceShape = ThisWorkbook.Sheets("IceShape")
    
    ' �����. ice score
    With Me.cboIceScore
        .Clear
        LoadComboBoxData wsIceScore, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    ' �����. Ice Type
    With Me.cboIceType
        .Clear
        LoadComboBoxData wsIceType, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    ' ����� Ice Shape
    With Me.cboIceShape
        .Clear
        LoadComboBoxData wsIceShape, .Name
        .TextColumn = 1
        .BoundColumn = 2
        .ColumnWidths = "200;0"
        .Style = fmStyleDropDownList
    End With
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Oshibka pri inicializacii dannyh l'da: " & vbNewLine & Err.Description, vbCritical
End Sub
Private Sub LoadComboBoxData(ws As Worksheet, comboName As String)
    ' ����� ����. ������
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    If lastRow < 2 Then Exit Sub  '���� ��� ������ ����� ���������
    
    ' ������������� ��� �������� (�������))
    Dim dataRange As Range
    Set dataRange = ws.Range("A2:B" & lastRow)
    
    ' ����. � ComboBox
    Select Case comboName
        Case "cboIceScore"
            Me.cboIceScore.List = dataRange.value
        Case "cboIceType"
            Me.cboIceType.List = dataRange.value
        Case "cboIceShape"
            Me.cboIceShape.List = dataRange.value
    End Select
End Sub
' ���������� �� 2 �������
Private Function GetTwoColumnValues(ws As Worksheet) As Variant
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    ' ������ ��� ���������������� (������� �������)
    Dim dataArray() As Variant
    ReDim dataArray(1 To lastRow - 1, 1 To 2)
    
    '������ ������ A i B (����. ���������)
    Dim i As Long
    For i = 2 To lastRow
        dataArray(i - 1, 1) = ws.Cells(i, "A").value
        dataArray(i - 1, 2) = ws.Cells(i, "B").value
    Next i
    
    GetTwoColumnValues = dataArray
End Function
Private Sub ClearAllFields()
    With Me
        .txtLongitude.Text = ""
        .txtLatitude.Text = ""
        .txtTemp.Text = ""
        .txtBarometer.Text = ""
        .txtWindDirection.Text = ""
        .txtWindSpeed.Text = ""
        .txtVisibility.Text = ""
        .txtSeaSwell.Text = ""
        .txtSeaSwellDirection.Text = ""
        .txtWindWaveDirection.Text = ""
        .txtWindWaveHeight.Text = ""
        
        If Not LatitudeInput.degrees Is Nothing Then LatitudeInput.degrees.Text = ""
        If Not LatitudeInput.minutes Is Nothing Then LatitudeInput.minutes.Text = ""
        If Not LongitudeInput.degrees Is Nothing Then LongitudeInput.degrees.Text = ""
        If Not LongitudeInput.minutes Is Nothing Then LongitudeInput.minutes.Text = ""
    End With
End Sub
' ��������� ���� ������ ������ �����������
Private Function ValidateData() As Boolean
    ' ������. �������� �� "n/d"
    If Not ValidateNoDataFields Then Exit Function
    
    ' ���� �� ������ ������ ����
    If Not ValidateRequiredFields Then Exit Function
    
    ' �������� �� ������ � �����
    If Not ValidateFieldErrors Then Exit Function
    
    ' �������� ���������
    If Not ValidateCoordinates Then
        MsgBox "Incorrect coordinate format!" & Chr(13) & "�������� ������ ���������!", vbExclamation
        Exit Function
    End If
    
    ValidateData = True
End Function
Private Function ValidateNoDataFields() As Boolean
    On Error GoTo ErrorHandler
    
    Dim fieldsToCheck As Variant
    fieldsToCheck = Array("txtTemp", "txtBarometer", "txtVisibility", "txtWindDirection", _
                         "txtWindSpeed", "txtSeaSwellDirection", "txtSeaSwell", _
                         "txtWindWaveDirection", "txtWindWaveHeight")
    
    Dim FieldName As Variant
    For Each FieldName In fieldsToCheck
        If Me.Controls(FieldName).Text = "n/d" Then
            If Not HasReason(CStr(FieldName)) Then
                MsgBox "���������� ������� ������� ���������� ������ ��� ���� " & FieldName, vbExclamation
                Me.Controls(FieldName).SetFocus
                ValidateNoDataFields = False
                Exit Function
            End If
        End If
    Next FieldName
    
    ValidateNoDataFields = True
    Exit Function

ErrorHandler:
    MsgBox "������ ��� �������� �����: " & Err.Description, vbExclamation
    ValidateNoDataFields = False
End Function
Private Function HasReason(FieldName As String) As Boolean
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Reasons")
    
    If ws Is Nothing Then
        HasReason = False
        Exit Function
    End If
    
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim i As Long
    For i = lastRow To 2 Step -1
        If ws.Cells(i, 2).value = FieldName Then
            HasReason = True
            Exit Function
        End If
    Next i
    
    HasReason = False
    Exit Function

ErrorHandler:
    HasReason = False
End Function

Private Function ValidateCoordinates() As Boolean
    If mCoordFormat = COORD_FORMAT_DECIMAL Then
        ' ���� ��������� ���������
        If Me.txtLongitude.Text = "" Or Me.txtLatitude.Text = "" Then Exit Function
        
        If Not IsNumeric(Replace(Me.txtLongitude.Text, ".", ",")) Or _
           Not IsNumeric(Replace(Me.txtLatitude.Text, ".", ",")) Then Exit Function
        
        Dim lon As Double, lat As Double
        lon = CDbl(Replace(Me.txtLongitude.Text, ".", ","))
        lat = CDbl(Replace(Me.txtLatitude.Text, ".", ","))
        
        If Abs(lon) > 180 Or Abs(lat) > 90 Then Exit Function
    Else
        ' ���� ��������� � ������� ��������-�����
        With Me
            If LatitudeInput.degrees.Text = "" Or _
               LatitudeInput.minutes.Text = "" Or _
               LatitudeInput.direction.Text = "" Then Exit Function
               
            If LongitudeInput.degrees.Text = "" Or _
               LongitudeInput.minutes.Text = "" Or _
               LongitudeInput.direction.Text = "" Then Exit Function
        End With
    End If
    
    ValidateCoordinates = True
End Function
Private Function ValidateRequiredFields() As Boolean
    ' ����. ���. �����
    If Me.txtDateTime1.value = "" Or _
       Me.txtTemp.value = "" Or _
       Me.txtBarometer.value = "" Or _
       Me.txtVisibility.value = "" Then
        MsgBox "Fill in all required fields!" & Chr(13) & "��������� ��� ������������ ����!", vbExclamation
        Exit Function
    End If
    
    '����� ��������� � ����� �� �������
    If mCoordFormat = COORD_FORMAT_DECIMAL Then
        If Me.txtLongitude.value = "" Or Me.txtLatitude.value = "" Then
            MsgBox "Enter coordinates!" & Chr(13) & "������� ����������!", vbExclamation
            Exit Function
        End If
    Else
        If LatitudeInput.degrees.Text = "" Or LatitudeInput.minutes.Text = "" Or _
           LongitudeInput.degrees.Text = "" Or LongitudeInput.minutes.Text = "" Then
            MsgBox "Enter coordinates!" & Chr(13) & "������� ����������!", vbExclamation
            Exit Function
        End If
    End If
    
    ' ����. ������ � �����
    If Me.txtWindDirection.value = "" Or Me.txtWindSpeed.value = "" Then
        If Me.txtWindDirection.value <> "0" And Me.txtWindSpeed.value <> "0" And _
           Me.txtWindDirection.value <> "n/d" And Me.txtWindSpeed.value <> "n/d" Then
            MsgBox "Enter wind data! For calm conditions enter 0." & Chr(13) & _
                  "������� ������ � �����! ��� ����� ������� 0.", vbExclamation
            Exit Function
        End If
    End If
    
    ' ����. ������ � ��������
    If Me.chkSeaSwell.value Then
        If Me.txtSeaSwell.value = "" Or Me.txtSeaSwellDirection.value = "" Or _
           Me.txtWindWaveDirection.value = "" Or Me.txtWindWaveHeight.value = "" Then
            ' ������. ��� ��� ���� "n/d"
            If Not (Me.txtSeaSwell.value = "n/d" And Me.txtSeaSwellDirection.value = "n/d" And _
                   Me.txtWindWaveDirection.value = "n/d" And Me.txtWindWaveHeight.value = "n/d") Then
                MsgBox "Fill in all wave activity fields!" & Chr(13) & "��������� ��� ���� ��������!", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    ' ���� ������ � ����
    If Me.chkIceNotated.value Then
        If Me.cboIceType.Text = "" Or Me.cboIceScore.Text = "" Or Me.cboIceShape.Text = "" Then
            MsgBox "Fill in all ice condition fields!" & Chr(13) & "��������� ��� ���� ��������� ����!", vbExclamation
            Exit Function
        End If
    End If
    
    ValidateRequiredFields = True
End Function
Private Function ValidateFieldErrors() As Boolean
    ' ���� �� ������ � ���. �����
    If Me.txtTemp.Text <> "n/d" And Me.txtTemp.ForeColor = RGB(255, 0, 0) Or _
       Me.txtBarometer.Text <> "n/d" And Me.txtBarometer.ForeColor = RGB(255, 0, 0) Or _
       Me.txtVisibility.Text <> "n/d" And Me.txtVisibility.ForeColor = RGB(255, 0, 0) Then
        MsgBox "Correct the invalid values!" & Chr(13) & "��������� ������������ ��������!", vbExclamation
        Exit Function
    End If
    
    ' ���� ����.
    If mCoordFormat = COORD_FORMAT_DECIMAL Then
        If Me.txtLongitude.ForeColor = RGB(255, 0, 0) Or _
           Me.txtLatitude.ForeColor = RGB(255, 0, 0) Then
            MsgBox "Correct the wrong coordinate values!" & Chr(13) & "��������� ������������ �������� ���������!", vbExclamation
            Exit Function
        End If
    End If
    
    ' ���� �����
    If (Me.txtWindDirection.Text <> "n/d" And Me.txtWindDirection.Text <> "0" And Me.txtWindDirection.ForeColor = RGB(255, 0, 0)) Or _
       (Me.txtWindSpeed.Text <> "n/d" And Me.txtWindSpeed.Text <> "0" And Me.txtWindSpeed.ForeColor = RGB(255, 0, 0)) Then
        MsgBox "Correct the wrong wind values!" & Chr(13) & "��������� ������������ �������� �����!", vbExclamation
        Exit Function
    End If
    
    ' ���� ����� ����.
    If Me.chkSeaSwell.value Then
        If (Me.txtSeaSwell.Text <> "n/d" And Me.txtSeaSwell.Text <> "0" And Me.txtSeaSwell.ForeColor = RGB(255, 0, 0)) Or _
           (Me.txtSeaSwellDirection.Text <> "n/d" And Me.txtSeaSwellDirection.Text <> "0" And Me.txtSeaSwellDirection.ForeColor = RGB(255, 0, 0)) Or _
           (Me.txtWindWaveDirection.Text <> "n/d" And Me.txtWindWaveDirection.Text <> "0" And Me.txtWindWaveDirection.ForeColor = RGB(255, 0, 0)) Or _
           (Me.txtWindWaveHeight.Text <> "n/d" And Me.txtWindWaveHeight.Text <> "0" And Me.txtWindWaveHeight.ForeColor = RGB(255, 0, 0)) Then
            MsgBox "Correct the wrong wave values!" & Chr(13) & "��������� ������������ �������� ��������!", vbExclamation
            Exit Function
        End If
    End If
    
    ValidateFieldErrors = True
End Function

Private Function ValidateSeaData() As Boolean
    If Not Me.chkPort.value Then
        If Me.chkSeaSwell.value Then
            ' ���� ������ � ����.
            If Me.txtSeaSwell.value = "" Or _
               Me.txtSeaSwellDirection.value = "" Or _
               Me.txtWindWaveDirection.value = "" Or _
               Me.txtWindWaveHeight.value = "" Then
                MsgBox "Zapolnite vse polya volneniya!", vbExclamation
                Exit Function
            End If
            
            If Me.txtSeaSwell.ForeColor = RGB(255, 0, 0) Or _
               Me.txtWindWaveHeight.ForeColor = RGB(255, 0, 0) Then
                MsgBox "Isprav'te nepravil'nye znacheniya volneniya!", vbExclamation
                Exit Function
            End If
        Else
            ' ����. ������ � ����
            If Me.cboIceType.value = "" Or Me.cboIceScore.value = "" Then
                MsgBox "Vyberite tip i ball l'da!", vbExclamation
                Exit Function
            End If
        End If
    End If
    
    ValidateSeaData = True
End Function
Private Sub txtDateTime1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' ����. ������ �����, �����, �������� � �����������
    Select Case KeyAscii
        Case 8  ' Backspace
            ' ������.
        Case 48 To 57  ' ����� 0-9
            ' ������.
        Case 46, 44  ' ����� � �����.
            If InStr(Me.txtDateTime1.Text, ".") > 0 Then
                KeyAscii = 0 ' ���� ���� �����
            Else
                KeyAscii = 46 ' ������ �����
            End If
        Case 32  ' ������
            ' ������. ������ ���� ������
            If InStr(Me.txtDateTime1.Text, " ") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0  ' ��� �����. ������� ���������
    End Select
End Sub

Private Sub txtDateTime1_Change()
    ValidateDateTime Me.txtDateTime1
End Sub
Private Sub ValidateDateTime(txt As MSForms.TextBox)
    ' ���� ������� ���� - �������
    If txt.Text = "" Then Exit Sub
    
    Dim isValid As Boolean
    isValid = False
    
    On Error Resume Next
    Dim testDate As Date
    testDate = CDate(txt.Text)
    If Err.Number = 0 Then
        isValid = True
    End If
    On Error GoTo 0
    
    If isValid Then
        txt.ForeColor = RGB(0, 0, 0)
    Else
        txt.ForeColor = RGB(255, 0, 0)
    End If
End Sub
Private Sub txtVisibility_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ����. Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ���. ����� "n/d"
    If (Me.txtVisibility.Text = "" Or Me.txtVisibility.SelLength = Len(Me.txtVisibility.Text)) And _
       (Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N") Then
        Me.txtVisibility.Text = "n/d"
        Me.txtVisibility.SelStart = 3
        KeyAscii = 0
        ShowNoDataDialog "txtVisibility"
        Exit Sub
    End If
    
    ' ���� ��� ���� "n/d", ������. ������ Backspace
    If Me.txtVisibility.Text = "n/d" Then
        If KeyAscii = 8 Then ' ������ Backspace
            Me.txtVisibility.Text = ""
        End If
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ ������ �����
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ ���������
    Dim newText As String
    If Me.txtVisibility.SelLength > 0 Then
        newText = Left(Me.txtVisibility.Text, Me.txtVisibility.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtVisibility.Text, Me.txtVisibility.SelStart + Me.txtVisibility.SelLength + 1)
    Else
        newText = Left(Me.txtVisibility.Text, Me.txtVisibility.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtVisibility.Text, Me.txtVisibility.SelStart + 1)
    End If
    
    ' �������� ��� �� ����� ���������
    If IsNumeric(newText) Then
        If CLng(newText) > 50 Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub

ErrorHandler:
    ' � ����. ����� ������ ���� ����
    KeyAscii = 0
End Sub
Private Sub txtSeaSwell_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ���� ���� ���������� - �����
    If Not Me.txtSeaSwell.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ��� ����� "n/d"
    If (Me.txtSeaSwell.Text = "" Or Me.txtSeaSwell.SelLength = Len(Me.txtSeaSwell.Text)) Then
        If Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
            Me.txtSeaSwell.Text = "n/d"
            Me.txtSeaSwell.SelStart = 3
            KeyAscii = 0
            ShowNoDataDialog "txtSeaSwell"
            Exit Sub
        End If
    End If
    
    ' ���� ��� "n/d", ������ ������ ���������
    If Me.txtSeaSwell.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ ����� ����� � �����
    Select Case KeyAscii
        Case 48 To 57  ' ����� 0-9
            ' ������ �������� � �������
            Dim newText As String
            If Me.txtSeaSwell.SelLength > 0 Then
                newText = Left(Me.txtSeaSwell.Text, Me.txtSeaSwell.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtSeaSwell.Text, Me.txtSeaSwell.SelStart + Me.txtSeaSwell.SelLength + 1)
            Else
                newText = Left(Me.txtSeaSwell.Text, Me.txtSeaSwell.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtSeaSwell.Text, Me.txtSeaSwell.SelStart + 1)
            End If
            
            ' ������ ��� ����� �� �����
            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 20 Then
                    KeyAscii = 0
                End If
            End If
            
        Case 44, 46  ' ����� ��� ����
            ' ������ ������ 1 �����
            If InStr(Me.txtSeaSwell.Text, ",") > 0 Or InStr(Me.txtSeaSwell.Text, ".") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 44  ' ������ �����
            End If
            
            ' �� ������ ����� � ������
            If Me.txtSeaSwell.SelStart = 0 Then
                KeyAscii = 0
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtWindWaveHeight_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ���� ���� ���������� - �����
    If Not Me.txtWindWaveHeight.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ��� ����� "n/d"
    If (Me.txtWindWaveHeight.Text = "" Or Me.txtWindWaveHeight.SelLength = Len(Me.txtWindWaveHeight.Text)) Then
        If Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
            Me.txtWindWaveHeight.Text = "n/d"
            Me.txtWindWaveHeight.SelStart = 3
            KeyAscii = 0
            ShowNoDataDialog "txtWindWaveHeight"
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ������ ������ Backspace
    If Me.txtWindWaveHeight.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ������ ����� ������� � �����
    Select Case KeyAscii
        Case 48 To 57  ' ����� 0-9
            ' ������ �����������
            Dim newText As String
            If Me.txtWindWaveHeight.SelLength > 0 Then
                newText = Left(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.SelStart + Me.txtWindWaveHeight.SelLength + 1)
            Else
                newText = Left(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindWaveHeight.Text, Me.txtWindWaveHeight.SelStart + 1)
            End If
            
            ' ������ ��� ����� �� ����� 20
            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 20 Then
                    KeyAscii = 0
                End If
            End If
            
        Case 44, 46  ' ����� ��� �����
            ' ������ ������ 1 �����
            If InStr(Me.txtWindWaveHeight.Text, ",") > 0 Or InStr(Me.txtWindWaveHeight.Text, ".") > 0 Then
                KeyAscii = 0
            Else
                KeyAscii = 44  ' ������ �������
            End If
            
            ' �� ������ ������� � ������
            If Me.txtWindWaveHeight.SelStart = 0 Then
                KeyAscii = 0
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
' ���� ��������� ���������
Private Sub InitializeCoordinateControls()
    ' ����� ����� ��� ������
    With Me.fraLatitude
        Set LatitudeInput.degrees = .Controls("txtLatDegrees")
        Set LatitudeInput.minutes = .Controls("txtLatMinutes")
        Set LatitudeInput.direction = .Controls("cboLatDirection")
        
        ' ����� ComboBox ��� ������ ������
        With LatitudeInput.direction
            .Clear
            .AddItem "N"
            .AddItem "S"
            .Text = "N"
            .Style = fmStyleDropDownList ' ���� ������ ����
        End With
        
        ' ����� ����� �����
        With LatitudeInput.degrees
            .MaxLength = 2 ' ����. 2 �����
            .Text = ""
        End With
        
        With LatitudeInput.minutes
            .MaxLength = 4 ' XX.X ������
            .Text = ""
        End With
    End With
    
    ' ����� ����� ��� �������
    With Me.fraLongitude
        Set LongitudeInput.degrees = .Controls("txtLonDegrees")
        Set LongitudeInput.minutes = .Controls("txtLonMinutes")
        Set LongitudeInput.direction = .Controls("cboLonDirection")
        
        ' �����. ComboBox ��� ���� �������
        With LongitudeInput.direction
            .Clear
            .AddItem "E"
            .AddItem "W"
            .Text = "E"
            .Style = fmStyleDropDownList ' ���� ������ ����
        End With
        
        ' ����� ����� �����
        With LongitudeInput.degrees
            .MaxLength = 3 ' ���� 3 ����
            .Text = ""
        End With
        
        With LongitudeInput.minutes
            .MaxLength = 4 ' XX.X format
            .Text = ""
        End With
    End With
End Sub
Private Sub optDecimalCoords_Click()
    mCoordFormat = COORD_FORMAT_DECIMAL
    UpdateCoordinateControls
    ConvertAndUpdateCoordinates
End Sub
Private Sub optDegreeCoords_Click()
    mCoordFormat = COORD_FORMAT_DEGREES
    UpdateCoordinateControls
    ConvertAndUpdateCoordinates
End Sub
Private Sub txtLongitude_Click()
    If Not Me.txtLongitude.Enabled Then
        optDecimalCoords.value = True
    End If
End Sub
Private Sub txtLatDegrees_Click()
    If Not LatitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub txtLonDegrees_Click()
    If Not LongitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub txtLatMinutes_Click()
    If Not LatitudeInput.minutes.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub txtLonMinutes_Click()
    If Not LongitudeInput.minutes.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub cboLatDirection_Click()
    If Not LatitudeInput.direction.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub cboLonDirection_Click()
    If Not LongitudeInput.direction.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
' ��� ����� �� ����� �����
Private Sub txtLatitude_Click()
    If Not Me.txtLatitude.Enabled Then
        optDecimalCoords.value = True
    End If
End Sub
Private Sub LatDegrees_Click()
    If Not LatitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub LonDegrees_Click()
    If Not LongitudeInput.degrees.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub fraLatitude_Click()
    If Not Me.fraLatitude.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub fraLongitude_Click()
    If Not Me.fraLongitude.Enabled Then
        optDegreeCoords.value = True
    End If
End Sub
Private Sub ConvertAndUpdateCoordinates()
    On Error GoTo ErrorHandler
    
    With Me
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            ' ����� �� ���� ���� � ����������
            If LatitudeInput.degrees.Text <> "" And LatitudeInput.minutes.Text <> "" Then
                Dim latVal As Double
                latVal = ConvertToDecimal(LatitudeInput.degrees.Text, _
                                        LatitudeInput.minutes.Text, _
                                        LatitudeInput.direction.Text)
                
                .txtLatitude.Text = FormatCoordinate(latVal)
            End If
            
            If LongitudeInput.degrees.Text <> "" And LongitudeInput.minutes.Text <> "" Then
                Dim lonVal As Double
                lonVal = ConvertToDecimal(LongitudeInput.degrees.Text, _
                                        LongitudeInput.minutes.Text, _
                                        LongitudeInput.direction.Text)
                
                .txtLongitude.Text = FormatCoordinate(lonVal)
            End If
        Else
            ' ����� �� �������
            If .txtLatitude.Text <> "" Then
                ConvertToDegreesMinutes CDbl(Replace(.txtLatitude.Text, ".", ",")), _
                                      LatitudeInput.degrees, _
                                      LatitudeInput.minutes, _
                                      LatitudeInput.direction, _
                                      True
            End If
            
            If .txtLongitude.Text <> "" Then
                ConvertToDegreesMinutes CDbl(Replace(.txtLongitude.Text, ".", ",")), _
                                      LongitudeInput.degrees, _
                                      LongitudeInput.minutes, _
                                      LongitudeInput.direction, _
                                      False
            End If
        End If
        
        
    End With
    Exit Sub

ErrorHandler:
    Debug.Print "Error in ConvertAndUpdateCoordinates: " & Err.Description
End Sub
Private Function FormatCoordinate(ByVal value As Double) As String
    ' ��� ������ 0
    Dim result As String
    result = Trim(Str(Round(Abs(value), 4)))
    
    If InStr(result, ".") > 0 Then
        result = Replace(result, ".", ",")
        ' ������ 0 ���� �����
        While Len(result) - InStr(result, ",") < 4
            result = result & "0"
        Wend
    Else
        result = result & ",0000"
    End If
    
    '������ ����
    If value < 0 Then
        result = "-" & result
    End If
    
    FormatCoordinate = result
End Function
Private Function FormatDecimalCoordinate(ByVal value As Double) As String
    ' ������ ������ � 4 ������� �������
    Dim strValue As String
    strValue = Format(Abs(value), "0,0000")
    
    ' ���� ����� 0 ����� �������
    Do While Left(strValue, 1) = "0" And Mid(strValue, 2, 1) <> ","
        strValue = Mid(strValue, 2)
    Loop
    
    ' ������ ���� - ���� �����
    If value < 0 Then strValue = "-" & strValue
    
    FormatDecimalCoordinate = strValue
End Function
Private Sub ConvertToDegreesMinutes(ByVal decimalValue As Double, _
                                  degreesBox As MSForms.TextBox, _
                                  minutesBox As MSForms.TextBox, _
                                  directionBox As MSForms.ComboBox, _
                                  ByVal isLatitude As Boolean)
                                  
    Debug.Print "Converting decimal value: " & decimalValue
    
    Dim isNegative As Boolean
    isNegative = (decimalValue < 0)
    decimalValue = Abs(decimalValue)
    
    ' ������ ������� � ������
    Dim degrees As Long
    Dim minutes As Double
    
    degrees = Int(decimalValue)
    minutes = (decimalValue - degrees) * 60
    minutes = Round(minutes, 1) '������ �� 1 ����������� �����
    
    Debug.Print "Calculated degrees: " & degrees
    Debug.Print "Calculated minutes: " & minutes
    
    ' ���� ����� ����� �������� ����� 60 �����
    If minutes >= 60 Then
        degrees = degrees + 1
        minutes = 0
    End If
    
    ' ��������� ��������
    degreesBox.Text = CStr(degrees)
    
    ' ���� ������ � ���� �������� ������
    Dim minutesStr As String
    minutesStr = Format(minutes, "0.0") '������� �����
    minutesStr = Replace(minutesStr, ".", ",") ' ����� �� �����
    minutesBox.Text = minutesStr
    
    Debug.Print "Final minutes string: " & minutesStr
    
    ' ��� �������
    If isLatitude Then
        directionBox.Text = IIf(isNegative, "S", "N")
    Else
        directionBox.Text = IIf(isNegative, "W", "E")
    End If
End Sub
Private Function ConvertToDecimal(ByVal degrees As String, ByVal minutes As String, ByVal direction As String) As Double
    On Error GoTo ErrorHandler
    ' ������� ���� ������
    degrees = Trim(degrees)
    minutes = Trim(minutes)
    direction = Trim(direction)
    
    ' ���������� � �����
    Dim deg As Double, min As Double
    deg = CDbl(degrees)
    
    ' ��� �����
    If InStr(minutes, ",") > 0 Then
        min = CDbl(minutes)
    Else
        min = CDbl(minutes & ",0")
    End If
    
    ' ������ ��������� ����
    ConvertToDecimal = deg + (min / 60)
    
    ' ���������� �����
    If direction = "S" Or direction = "W" Then
        ConvertToDecimal = -ConvertToDecimal
    End If
    
    Exit Function
    
ErrorHandler:
    Debug.Print "Error in ConvertToDecimal: " & Err.Description
    ConvertToDecimal = 0
End Function
Private Sub UpdateCoordinateControls()
    Dim activeBackColor As Long, inactiveBackColor As Long
    Dim activeTextColor As Long, inactiveTextColor As Long
    
    activeBackColor = vbWhite
    inactiveBackColor = RGB(240, 240, 240)
    activeTextColor = vbBlack
    inactiveTextColor = RGB(192, 192, 192)
    
    With Me.fraMain.fraCoordinates
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            ' Decimal Degrees �������
            .txtLatitude.BackColor = activeBackColor
            .txtLongitude.BackColor = activeBackColor
            .txtLatitude.ForeColor = activeTextColor
            .txtLongitude.ForeColor = activeTextColor
            .lblLatitude.ForeColor = activeTextColor
            .lblLongitude.ForeColor = activeTextColor
            .txtLatitude.Locked = False
            .txtLongitude.Locked = False
            .txtLatitude.Enabled = True
            .txtLongitude.Enabled = True
            
            ' Degrees/Minutes ���������
            .txtLatDegrees.BackColor = inactiveBackColor
            .txtLatMinutes.BackColor = inactiveBackColor
            .cboLatDirection.BackColor = inactiveBackColor
            .txtLonDegrees.BackColor = inactiveBackColor
            .txtLonMinutes.BackColor = inactiveBackColor
            .cboLonDirection.BackColor = inactiveBackColor
            
            .txtLatDegrees.ForeColor = inactiveTextColor
            .txtLatMinutes.ForeColor = inactiveTextColor
            .cboLatDirection.ForeColor = inactiveTextColor
            .txtLonDegrees.ForeColor = inactiveTextColor
            .txtLonMinutes.ForeColor = inactiveTextColor
            .cboLonDirection.ForeColor = inactiveTextColor
            
            ' ���� �����������
            .txtLatDegrees.Locked = True
            .txtLatMinutes.Locked = True
            .txtLonDegrees.Locked = True
            .txtLonMinutes.Locked = True
            .cboLatDirection.Locked = True
            .cboLonDirection.Locked = True
            
            .txtLatDegrees.Enabled = False
            .txtLatMinutes.Enabled = False
            .txtLonDegrees.Enabled = False
            .txtLonMinutes.Enabled = False
            .cboLatDirection.Enabled = False
            .cboLonDirection.Enabled = False
            
        Else
            ' Degrees/Minutes �������
            .txtLatDegrees.BackColor = activeBackColor
            .txtLatMinutes.BackColor = activeBackColor
            .cboLatDirection.BackColor = activeBackColor
            .txtLonDegrees.BackColor = activeBackColor
            .txtLonMinutes.BackColor = activeBackColor
            .cboLonDirection.BackColor = activeBackColor
            
            .txtLatDegrees.ForeColor = activeTextColor
            .txtLatMinutes.ForeColor = activeTextColor
            .cboLatDirection.ForeColor = activeTextColor
            .txtLonDegrees.ForeColor = activeTextColor
            .txtLonMinutes.ForeColor = activeTextColor
            .cboLonDirection.ForeColor = activeTextColor
            
            .txtLatDegrees.Locked = False
            .txtLatMinutes.Locked = False
            .txtLonDegrees.Locked = False
            .txtLonMinutes.Locked = False
            .cboLatDirection.Locked = False
            .cboLonDirection.Locked = False
            
            .txtLatDegrees.Enabled = True
            .txtLatMinutes.Enabled = True
            .txtLonDegrees.Enabled = True
            .txtLonMinutes.Enabled = True
            .cboLatDirection.Enabled = True
            .cboLonDirection.Enabled = True
            
            ' Decimal Degrees ���������
            .txtLatitude.BackColor = inactiveBackColor
            .txtLongitude.BackColor = inactiveBackColor
            .txtLatitude.ForeColor = inactiveTextColor
            .txtLongitude.ForeColor = inactiveTextColor
            .lblLatitude.ForeColor = inactiveTextColor
            .lblLongitude.ForeColor = inactiveTextColor
            .txtLatitude.Locked = True
            .txtLongitude.Locked = True
            .txtLatitude.Enabled = False
            .txtLongitude.Enabled = False
        End If
    End With
End Sub
Private Sub UpdateSeaIceControls()
    Dim activeColor As Long
    Dim inactiveColor As Long
    activeColor = RGB(0, 0, 0)
    inactiveColor = RGB(192, 192, 192)
    
    With Me
        If .chkSeaSwell.value Then
            ' Sea/Wave �������
            .txtSeaSwell.Enabled = True
            .txtSeaSwellDirection.Enabled = True
            .txtWindWaveDirection.Enabled = True
            .txtWindWaveHeight.Enabled = True
            .txtSeaSwell.BackColor = vbWhite
            .txtSeaSwellDirection.BackColor = vbWhite
            .txtWindWaveDirection.BackColor = vbWhite
            .txtWindWaveHeight.BackColor = vbWhite
            
            ' Ice ���������
            .cboIceType.Enabled = False
            .cboIceScore.Enabled = False
            .cboIceType.BackColor = RGB(240, 240, 240)
            .cboIceScore.BackColor = RGB(240, 240, 240)
            .lblIceType.ForeColor = inactiveColor
            .lblIceScore.ForeColor = inactiveColor
        Else
            ' Ice �������
            .cboIceType.Enabled = True
            .cboIceScore.Enabled = True
            .cboIceType.BackColor = vbWhite
            .cboIceScore.BackColor = vbWhite
            .lblIceType.ForeColor = activeColor
            .lblIceScore.ForeColor = activeColor
            
            ' Sea/Wave ���������
            .txtSeaSwell.Enabled = False
            .txtSeaSwellDirection.Enabled = False
            .txtWindWaveDirection.Enabled = False
            .txtWindWaveHeight.Enabled = False
            .txtSeaSwell.BackColor = RGB(240, 240, 240)
            .txtSeaSwellDirection.BackColor = RGB(240, 240, 240)
            .txtWindWaveDirection.BackColor = RGB(240, 240, 240)
            .txtWindWaveHeight.BackColor = RGB(240, 240, 240)
        End If
        
        ' ���� port ������
        If .chkPort.value Then
            .txtSeaSwell.Enabled = False
            .txtSeaSwellDirection.Enabled = False
            .txtWindWaveDirection.Enabled = False
            .txtWindWaveHeight.Enabled = False
            .cboIceType.Enabled = False
            .cboIceScore.Enabled = False
            .txtSeaSwell.BackColor = RGB(240, 240, 240)
            .txtSeaSwellDirection.BackColor = RGB(240, 240, 240)
            .txtWindWaveDirection.BackColor = RGB(240, 240, 240)
            .txtWindWaveHeight.BackColor = RGB(240, 240, 240)
            .cboIceType.BackColor = RGB(240, 240, 240)
            .cboIceScore.BackColor = RGB(240, 240, 240)
        End If
    End With
End Sub
' ���� ��� ��� �����
Private Sub HandleMinutesKeyPress(ByVal KeyAscii As MSForms.ReturnInteger, txt As MSForms.TextBox)
    Debug.Print "Current text: [" & txt.Text & "], KeyAscii: " & KeyAscii
    
    Select Case KeyAscii
        Case 8  ' Backspace
            ' ������ ������
            
        Case 44, 46 ' ����� ��� �����
            ' ������ ����� ���� ��� ���
            If InStr(txt.Text, ",") = 0 Then
                KeyAscii = 44 ' ������ �����
            Else
                KeyAscii = 0
            End If
            
        Case 48 To 57 ' �����
            ' ����� ������� ������� � �������
            Dim cursorPos As Long
            Dim commaPos As Long
            cursorPos = txt.SelStart
            commaPos = InStr(txt.Text, ",")
            
            ' ������ ����� �� ��� �����
            If commaPos = 0 Then
                ' �� �������
                If Len(txt.Text) - txt.SelLength >= 2 Then
                    ' �������� �� ������ 59
                    Dim newValue As String
                    newValue = Left(txt.Text, txt.SelStart) & Chr(KeyAscii) & _
                              Mid(txt.Text, txt.SelStart + txt.SelLength + 1)
                    If IsNumeric(newValue) Then
                        If CDbl(newValue) >= 60 Then
                            KeyAscii = 0
                        End If
                    End If
                End If
            Else
                ' ����� �������
                If cursorPos > commaPos Then
                    ' ������ ������ 1 ����� ����� �������
                    If Len(txt.Text) - commaPos >= 2 Then
                        KeyAscii = 0
                    End If
                End If
            End If
            
        Case Else
            KeyAscii = 0
    End Select
    
    ' ���������� ����������
    Debug.Print "After processing: KeyAscii = " & KeyAscii
End Sub
' ������ � �������� �����
Private Sub txtLatitude_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleNumericKeyPress KeyAscii, Me.txtLatitude, True, 90
End Sub
Private Sub txtLongitude_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleNumericKeyPress KeyAscii, Me.txtLongitude, True, 180
End Sub
Private Sub txtLatDegrees_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleNumericKeyPress KeyAscii, LatitudeInput.degrees, False, 90
End Sub
Private Sub txtLonDegrees_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleNumericKeyPress KeyAscii, LongitudeInput.degrees, False, 180
End Sub
Private Sub txtLatMinutes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleMinutesKeyPress KeyAscii, LatitudeInput.minutes
End Sub
Private Sub txtLonMinutes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    HandleMinutesKeyPress KeyAscii, LongitudeInput.minutes
End Sub
' ��� ������� ��� Decimal Degrees
Private Sub txtLatitude_Change()
    If Not Me.fraMain.fraCoordinates.txtLatitude.Locked Then
        ConvertDecimalToMinutes
    End If
End Sub
Private Sub txtLongitude_Change()
    If Not Me.fraMain.fraCoordinates.txtLongitude.Locked Then
        ConvertDecimalToMinutes
    End If
End Sub
Private Sub txtLatDegrees_Change()
    If Not Me.fraMain.fraCoordinates.txtLatDegrees.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
Private Sub txtLatMinutes_Change()
    If Not Me.fraMain.fraCoordinates.txtLatMinutes.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
Private Sub txtLonMinutes_Change()
    If Not Me.fraMain.fraCoordinates.txtLonMinutes.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
Private Sub cboLonDirection_Change()
    If Not Me.fraMain.fraCoordinates.cboLonDirection.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
Private Sub cboLatDirection_Change()
    If Not Me.fraMain.fraCoordinates.cboLatDirection.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
' ����� ����������
Private Sub ConvertDecimalToMinutes()
    On Error GoTo ErrorHandler
    
    With Me.fraMain.fraCoordinates
        ' ������� ������
        If .txtLatitude.Text <> "" And IsNumeric(Replace(.txtLatitude.Text, ".", ",")) Then
            Dim latValue As Double
            latValue = CDbl(Replace(.txtLatitude.Text, ".", ","))
            
            If Abs(latValue) <= 90 Then
                ConvertToDegreesMinutes latValue, _
                                      LatitudeInput.degrees, _
                                      LatitudeInput.minutes, _
                                      LatitudeInput.direction, _
                                      True
            End If
        End If
        
        ' ������� �������
        If .txtLongitude.Text <> "" And IsNumeric(Replace(.txtLongitude.Text, ".", ",")) Then
            Dim lonValue As Double
            lonValue = CDbl(Replace(.txtLongitude.Text, ".", ","))
            
            If Abs(lonValue) <= 180 Then
                ConvertToDegreesMinutes lonValue, _
                                      LongitudeInput.degrees, _
                                      LongitudeInput.minutes, _
                                      LongitudeInput.direction, _
                                      False
            End If
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ConvertDecimalToMinutes: " & Err.Description
End Sub
Private Sub ConvertMinutesToDecimal()
    On Error GoTo ErrorHandler
    
    With Me.fraMain.fraCoordinates
        ' ������� ������
        If LatitudeInput.degrees.Text <> "" And LatitudeInput.minutes.Text <> "" Then
            Dim latDec As Double
            latDec = ConvertToDecimal(LatitudeInput.degrees.Text, _
                                    LatitudeInput.minutes.Text, _
                                    LatitudeInput.direction.Text)
            .txtLatitude.Text = FormatCoordinate(latDec)
        End If
        
        ' �������
        If LongitudeInput.degrees.Text <> "" And LongitudeInput.minutes.Text <> "" Then
            Dim lonDec As Double
            lonDec = ConvertToDecimal(LongitudeInput.degrees.Text, _
                                    LongitudeInput.minutes.Text, _
                                    LongitudeInput.direction.Text)
            .txtLongitude.Text = FormatCoordinate(lonDec)
        End If
    End With
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in ConvertMinutesToDecimal: " & Err.Description
End Sub
Private Sub txtLonDegrees_Change()
    If Not Me.fraMain.fraCoordinates.txtLonDegrees.Locked Then
        ConvertMinutesToDecimal
    End If
End Sub
' ��������������� �������
Private Function IsValidDecimalCoordinate(ByVal value As String, ByVal isLatitude As Boolean) As Boolean
    If value = "" Or value = "-" Then Exit Function
    If Not IsNumeric(Replace(value, ".", ",")) Then Exit Function
    
    Dim numValue As Double
    numValue = CDbl(Replace(value, ".", ","))
    
    If isLatitude Then
        IsValidDecimalCoordinate = (Abs(numValue) <= 90)
    Else
        IsValidDecimalCoordinate = (Abs(numValue) <= 180)
    End If
End Function
Private Sub UpdateDecimalFromDegrees(ByVal isLatitude As Boolean)
    On Error GoTo ErrorHandler
    
    Dim degrees As TextBox
    Dim minutes As TextBox
    Dim direction As ComboBox
    Dim decimalOutput As TextBox
    
    ' �������� ������ ������� �����������
    If isLatitude Then
        Set degrees = LatitudeInput.degrees
        Set minutes = LatitudeInput.minutes
        Set direction = LatitudeInput.direction
        Set decimalOutput = Me.fraMain.fraCoordinates.txtLatitude
    Else
        Set degrees = LongitudeInput.degrees
        Set minutes = LongitudeInput.minutes
        Set direction = LongitudeInput.direction
        Set decimalOutput = Me.fraMain.fraCoordinates.txtLongitude
    End If
    
    ' ��������� ��� ���� �������� ���������
    If degrees.Text = "" Or minutes.Text = "" Or direction.Text = "" Then Exit Sub
    If Not IsNumeric(degrees.Text) Or Not IsNumeric(Replace(minutes.Text, ",", ".")) Then Exit Sub
    
    ' ������ ��������
    Dim deg As Double, min As Double
    deg = CDbl(degrees.Text)
    min = CDbl(Replace(minutes.Text, ",", "."))
    
    ' ������ �������� ��������
    If min >= 60 Then Exit Sub
    If isLatitude And deg > 90 Then Exit Sub
    If Not isLatitude And deg > 180 Then Exit Sub
    
    ' ������ ��������� �������
    Dim decimalValue As Double
    decimalValue = deg + (min / 60)
    
    ' ������ ����
    If direction.Text = "S" Or direction.Text = "W" Then
        decimalValue = -decimalValue
    End If
    
    ' ����� ����������
    decimalOutput.Text = FormatCoordinate(decimalValue)
    
    Exit Sub
    
ErrorHandler:
    Debug.Print "Error in UpdateDecimalFromDegrees: " & Err.Description
End Sub
' ��������� ��������� �������� ��� �����
Private Sub ValidateMinutes(txt As MSForms.TextBox)
    If Len(txt.Text) = 0 Then Exit Sub
    If txt.Text = "," Then
        txt.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    ' ����� ����� �� ������� ���� ����'
    Dim value As String
    value = Replace(txt.Text, ".", ",")
    
    If Not IsNumeric(value) Then
        txt.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    Dim numValue As Double
    numValue = CDbl(value)
    
    ' ������ ��������
    If numValue >= 60 Or numValue < 0 Then
        txt.ForeColor = RGB(255, 0, 0)
    Else
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub
' ������� ��� �������������� ���������
Private Function GetDecimalCoordinates(degrees As String, minutes As String, direction As String) As Double
    If Len(degrees) = 0 Or Len(minutes) = 0 Then Exit Function
    
    Dim deg As Double
    Dim min As Double
    
    deg = Val(degrees)
    min = Val(Replace(minutes, ".", ","))
    
    GetDecimalCoordinates = deg + (min / 60)
    
    If direction = "S" Or direction = "W" Then
        GetDecimalCoordinates = -GetDecimalCoordinates
    End If
End Function
Private Sub SetDegreeMinuteCoordinates(ByVal decimalValue As Double, degreesBox As MSForms.TextBox, minutesBox As MSForms.TextBox, directionBox As MSForms.ComboBox)
    Dim isNegative As Boolean
    isNegative = (decimalValue < 0)
    decimalValue = Abs(decimalValue)
    
    Dim degrees As Long
    Dim minutes As Double
    
    degrees = Int(decimalValue)
    minutes = (decimalValue - degrees) * 60
    
    degreesBox.Text = Format(degrees, "00")
    minutesBox.Text = Format(minutes, "00.0")
    
    If TypeOf directionBox.Parent Is MSForms.Frame Then
        If directionBox.Parent.Name = "fraLatitude" Then
            directionBox.Text = IIf(isNegative, "S", "N")
        Else
            directionBox.Text = IIf(isNegative, "W", "E")
        End If
    End If
End Sub
' ��������� �������� ������
Private Sub chkCalm_Click()
    mIsCalm = Me.chkCalm.value
    UpdateWindControls
End Sub
Private Sub UpdateWindControls()
    With Me
        .txtWindDirection.Enabled = Not mIsCalm
        .txtWindSpeed.Enabled = Not mIsCalm
        
        If mIsCalm Then
            .txtWindDirection.Text = "calm"
            .txtWindSpeed.Text = "calm"
        Else
            .txtWindDirection.Text = ""
            .txtWindSpeed.Text = ""
        End If
    End With
End Sub
Private Sub chkPort_Click()
    mIsPort = Me.chkPort.value
    UpdatePortControls
End Sub
Private Sub UpdateSeaControls()
    Dim activeBackColor As Long, inactiveBackColor As Long
    Dim activeTextColor As Long, inactiveTextColor As Long
    
    activeBackColor = vbWhite
    inactiveBackColor = RGB(240, 240, 240)
    activeTextColor = vbBlack
    inactiveTextColor = RGB(192, 192, 192)
    
    With Me
        ' ��������� ����� ��������
        If .chkSeaSwell.value Then
            ' ���������� ���� ��������
            .txtSeaSwell.BackColor = activeBackColor
            .txtSeaSwellDirection.BackColor = activeBackColor
            .txtWindWaveDirection.BackColor = activeBackColor
            .txtWindWaveHeight.BackColor = activeBackColor
            
            .txtSeaSwell.ForeColor = activeTextColor
            .txtSeaSwellDirection.ForeColor = activeTextColor
            .txtWindWaveDirection.ForeColor = activeTextColor
            .txtWindWaveHeight.ForeColor = activeTextColor
            
            .lblSeaSwell.ForeColor = activeTextColor
            .lblSeaSwellDirection.ForeColor = activeTextColor
            .lblWindWaveDirection.ForeColor = activeTextColor
            .lblWindWaveHeight.ForeColor = activeTextColor
            
            .txtSeaSwell.Enabled = True
            .txtSeaSwellDirection.Enabled = True
            .txtWindWaveDirection.Enabled = True
            .txtWindWaveHeight.Enabled = True
            
            .txtSeaSwell.Locked = False
            .txtSeaSwellDirection.Locked = False
            .txtWindWaveDirection.Locked = False
            .txtWindWaveHeight.Locked = False
            
            ' ������� ����� ���� ��� ���� "0"
            If .txtSeaSwell.Text = "0" Then .txtSeaSwell.Text = ""
            If .txtSeaSwellDirection.Text = "0" Then .txtSeaSwellDirection.Text = ""
            If .txtWindWaveDirection.Text = "0" Then .txtWindWaveDirection.Text = ""
            If .txtWindWaveHeight.Text = "0" Then .txtWindWaveHeight.Text = ""
        Else
            '������������ ���� �������� ��� ��������� ������
            .txtSeaSwell.BackColor = inactiveBackColor
            .txtSeaSwellDirection.BackColor = inactiveBackColor
            .txtWindWaveDirection.BackColor = inactiveBackColor
            .txtWindWaveHeight.BackColor = inactiveBackColor
            
            .txtSeaSwell.ForeColor = inactiveTextColor
            .txtSeaSwellDirection.ForeColor = inactiveTextColor
            .txtWindWaveDirection.ForeColor = inactiveTextColor
            .txtWindWaveHeight.ForeColor = inactiveTextColor
            
            .lblSeaSwell.ForeColor = inactiveTextColor
            .lblSeaSwellDirection.ForeColor = inactiveTextColor
            .lblWindWaveDirection.ForeColor = inactiveTextColor
            .lblWindWaveHeight.ForeColor = inactiveTextColor
            
            .txtSeaSwell.Enabled = False
            .txtSeaSwellDirection.Enabled = False
            .txtWindWaveDirection.Enabled = False
            .txtWindWaveHeight.Enabled = False
            
            .txtSeaSwell.Text = "0"
            .txtSeaSwellDirection.Text = "0"
            .txtWindWaveDirection.Text = "0"
            .txtWindWaveHeight.Text = "0"
        End If
        
        ' ��������� ����� ����
        If .chkIceNotated.value Then
            ' ���������� ���� ����
            .cboIceType.BackColor = activeBackColor
            .cboIceScore.BackColor = activeBackColor
            .cboIceShape.BackColor = activeBackColor
            
            .cboIceType.ForeColor = activeTextColor
            .cboIceScore.ForeColor = activeTextColor
            .cboIceShape.ForeColor = activeTextColor
            
            .lblIceType.ForeColor = activeTextColor
            .lblIceScore.ForeColor = activeTextColor
            .lblIceShape.ForeColor = activeTextColor
            
            .cboIceType.Enabled = True
            .cboIceScore.Enabled = True
            .cboIceShape.Enabled = True
            
            ' ���� ���� ������ ��� �������� "������ ����" - ������� ��
            If .cboIceType.Text = "������ ����" Then .cboIceType.ListIndex = -1
            If .cboIceScore.Text = "������ ����" Then .cboIceScore.ListIndex = -1
            If .cboIceShape.Text = "������ ����" Then .cboIceShape.ListIndex = -1
        Else
            ' ������������ ���� ����
            .cboIceType.BackColor = inactiveBackColor
            .cboIceScore.BackColor = inactiveBackColor
            .cboIceShape.BackColor = inactiveBackColor
            
            .cboIceType.ForeColor = inactiveTextColor
            .cboIceScore.ForeColor = inactiveTextColor
            .cboIceShape.ForeColor = inactiveTextColor
            
            .lblIceType.ForeColor = inactiveTextColor
            .lblIceScore.ForeColor = inactiveTextColor
            .lblIceShape.ForeColor = inactiveTextColor
            
            .cboIceType.Enabled = False
            .cboIceScore.Enabled = False
            .cboIceShape.Enabled = False
            
            ' ������������� �������� ��� ���������
            .cboIceType.Text = "������ ����"
            .cboIceScore.Text = "������ ����"
            .cboIceShape.Text = "������ ����"
        End If
    End With
End Sub
Private Sub SetAllWaveFieldsToND()
    With Me
        .txtSeaSwell.Text = "n/d"
        .txtSeaSwellDirection.Text = "n/d"
        .txtWindWaveDirection.Text = "n/d"
        .txtWindWaveHeight.Text = "n/d"
    End With
End Sub

Private Sub ClearSeaSwellFields()
    With Me
        .txtSeaSwell.Text = ""
        .txtSeaSwellDirection.Text = ""
        .txtWindWaveDirection.Text = ""
        .txtWindWaveHeight.Text = ""
    End With
End Sub
Private Sub chkIceNotated_Click()
    mIsIceNotated = Me.chkIceNotated.value
    UpdateSeaControls
End Sub
Private Sub chkSeaSwell_Click()
    UpdateSeaControls
End Sub
Private Sub ShowSeaSwellControls()
    With Me
        .txtSeaSwell.Visible = True
        .txtSeaSwellDirection.Visible = True
        .txtWindWaveDirection.Visible = True
        .txtWindWaveHeight.Visible = True
        .lblSeaSwell.Visible = True
        .lblSeaSwellDirection.Visible = True
        .lblWindWaveDirection.Visible = True
        .lblWindWaveHeight.Visible = True
        
        .cboIceType.Visible = False
        .cboIceScore.Visible = False
        .lblIceType.Visible = False
        .lblIceScore.Visible = False
    End With
End Sub
Private Sub HideSeaSwellControls()
    With Me
        .txtSeaSwell.Visible = False
        .txtSeaSwellDirection.Visible = False
        .txtWindWaveDirection.Visible = False
        .txtWindWaveHeight.Visible = False
        .lblSeaSwell.Visible = False
        .lblSeaSwellDirection.Visible = False
        .lblWindWaveDirection.Visible = False
        .lblWindWaveHeight.Visible = False
        
        .cboIceType.Visible = True
        .cboIceScore.Visible = True
        .lblIceType.Visible = True
        .lblIceScore.Visible = True
    End With
End Sub
' ��������� ��������� ��������
Private Sub ValidateNumeric(txt As MSForms.TextBox, Optional AllowMinus As Boolean = True)
    If Len(txt.Text) = 0 Then Exit Sub
    If txt.Text = "-" And AllowMinus Then Exit Sub
    
    If Not IsNumeric(Replace(Replace(txt.Text, ".", ","), "-", "")) Then
        txt.ForeColor = RGB(255, 0, 0)
    Else
        Dim value As Double
        value = CDbl(Replace(txt.Text, ".", ","))
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub
Private Sub ValidateNumericRange(txt As MSForms.TextBox, MinValue As Double, MaxValue As Double)
    If Len(txt.Text) = 0 Then Exit Sub
    If txt.Text = "-" Then Exit Sub
    
    If Not IsNumeric(Replace(txt.Text, ",", ".")) Then
        txt.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    Dim value As Double
    value = CDbl(Replace(txt.Text, ",", "."))
    
    If value < MinValue Or value > MaxValue Then
        txt.ForeColor = RGB(255, 0, 0)
    Else
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub
Private Sub ValidatePositiveNumeric(txt As MSForms.TextBox)
    If Len(txt.Text) = 0 Then Exit Sub
    
    If Not IsNumeric(Replace(txt.Text, ".", ",")) Then
        txt.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    Dim value As Double
    value = CDbl(Replace(txt.Text, ".", ","))
    
    If value < 0 Then
        txt.ForeColor = RGB(255, 0, 0)
    Else
        txt.ForeColor = RGB(0, 0, 0)
    End If
End Sub
Private Sub txtBarometer_Change()
    ValidateNumeric Me.txtBarometer, True
End Sub
Private Sub txtWindDirection_Change()
    ' ���� ���� ������
    If Me.txtWindDirection.Text = "" Then
        Me.txtWindDirection.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    
    ' ���� ������� "n/d"
    If Me.txtWindDirection.Text = "n/d" Then
        Me.txtWindDirection.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    
    ' �������� �� �������� ��������
    If Not IsNumeric(Me.txtWindDirection.Text) Then
        Me.txtWindDirection.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    ' �������� ���������
    Dim value As Double
    value = CDbl(Me.txtWindDirection.Text)
    
    If value >= 0 And value <= 360 Then
        Me.txtWindDirection.ForeColor = RGB(0, 0, 0)
    Else
        Me.txtWindDirection.ForeColor = RGB(255, 0, 0)
    End If
End Sub
Private Sub txtWindSpeed_Change()
    ' ���� ���� ������
    If Me.txtWindSpeed.Text = "" Then
        Me.txtWindSpeed.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    
    ' ���� ������� "n/d"
    If Me.txtWindSpeed.Text = "n/d" Then
        Me.txtWindSpeed.ForeColor = RGB(0, 0, 0)
        Exit Sub
    End If
    
    ' �������� �� �������� ��������
    If Not IsNumeric(Replace(Me.txtWindSpeed.Text, ",", ".")) Then
        Me.txtWindSpeed.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    ' �������� ���������
    Dim value As Double
    value = CDbl(Replace(Me.txtWindSpeed.Text, ",", "."))
    
    If value >= 0 And value <= 100 Then
        Me.txtWindSpeed.ForeColor = RGB(0, 0, 0)
    Else
        Me.txtWindSpeed.ForeColor = RGB(255, 0, 0)
    End If
End Sub
Private Sub txtSeaSwell_Change()
    If Me.txtSeaSwell.Text = "n/d" And Me.chkSeaSwell.value Then
        SetAllWaveFieldsToND
    End If
End Sub
Private Sub txtSeaSwellDirection_Change()
    If Me.txtSeaSwellDirection.Text = "n/d" And Me.chkSeaSwell.value Then
        SetAllWaveFieldsToND
    End If
End Sub
Private Sub txtWindWaveDirection_Change()
    If Me.txtWindWaveDirection.Text = "n/d" And Me.chkSeaSwell.value Then
        SetAllWaveFieldsToND
    End If
End Sub
Private Sub txtWindWaveHeight_Change()
    If Me.txtWindWaveHeight.Text = "n/d" And Me.chkSeaSwell.value Then
        SetAllWaveFieldsToND
    End If
End Sub
Private Sub txtVisibility_Change()
    ' �������� ��������
    If Me.txtVisibility.Text = "n/d" Then Exit Sub
    If Me.txtVisibility.Text = "" Then Exit Sub
    
    If Not IsNumeric(Me.txtVisibility.Text) Then
        Me.txtVisibility.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    Dim value As Long
    value = CLng(Me.txtVisibility.Text)
    
    If value >= 0 And value <= 1000 Then
        Me.txtVisibility.ForeColor = RGB(0, 0, 0)
    Else
        Me.txtVisibility.ForeColor = RGB(255, 0, 0)
    End If
End Sub
Private Sub txtTemp_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    '��������� Backspace ������
    If KeyAscii = 8 Then Exit Sub
    ' �������� �� ���� "n/d"
    If Me.txtTemp.Text = "" Or Me.txtTemp.SelLength = Len(Me.txtTemp.Text) Then
        If KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
            Me.txtTemp.Text = "n/d"
            Me.txtTemp.SelStart = 3
            ShowNoDataDialog "txtTemp"
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ��������� ������ Backspace
    If Me.txtTemp.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ����� � ������
    If KeyAscii = 45 And (Me.txtTemp.Text = "" Or Me.txtTemp.SelLength = Len(Me.txtTemp.Text)) Then
        Exit Sub
    End If
    
    ' ��������� ����� � �������
    Select Case KeyAscii
        Case 48 To 57 ' �����
            ' ���������� ��� ��������� � ����������
            Dim newText As String
            If Me.txtTemp.SelLength > 0 Then
                newText = Left(Me.txtTemp.Text, Me.txtTemp.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtTemp.Text, Me.txtTemp.SelStart + Me.txtTemp.SelLength + 1)
            Else
                newText = Left(Me.txtTemp.Text, Me.txtTemp.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtTemp.Text, Me.txtTemp.SelStart + 1)
            End If
            
            If IsNumeric(Replace(newText, ",", ".")) Then
                If Abs(CDbl(Replace(newText, ",", "."))) > 100 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' ������� ��� �����
            If InStr(Me.txtTemp.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = 44 ' ������ �������
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtTemp_Change()
    '�������� �� �������� �����
    If Me.txtTemp.Text = "-" Then
        Me.txtTemp.ForeColor = RGB(255, 0, 0)
        Exit Sub
    End If
    
    Me.txtTemp.ForeColor = RGB(0, 0, 0)
End Sub
Private Sub txtSeaSwellDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    ' ���� ���� �� �������� - �����
    If Not Me.txtSeaSwellDirection.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ��������� ����� "n/d"
    If (Me.txtSeaSwellDirection.Text = "" Or Me.txtSeaSwellDirection.SelLength = Len(Me.txtSeaSwellDirection.Text)) Then
        If Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
            Me.txtSeaSwellDirection.Text = "n/d"
            Me.txtSeaSwellDirection.SelStart = 3
            KeyAscii = 0
            ShowNoDataDialog "txtSeaSwellDirection"
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ��������� ������ Backspace
    If Me.txtSeaSwellDirection.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ������ �����
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ��� ��������� � ����������
    Dim newText As String
    If Me.txtSeaSwellDirection.SelLength > 0 Then
        newText = Left(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.SelStart + Me.txtSeaSwellDirection.SelLength + 1)
    Else
        newText = Left(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtSeaSwellDirection.Text, Me.txtSeaSwellDirection.SelStart + 1)
    End If
    
    ' ��������� ��� ���� �� ������ 360
    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtWindWaveDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    ' ���� ���� �� �������� - �����
    If Not Me.txtWindWaveDirection.Enabled Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ��������� ����� "n/d"
    If (Me.txtWindWaveDirection.Text = "" Or Me.txtWindWaveDirection.SelLength = Len(Me.txtWindWaveDirection.Text)) Then
        If Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
            Me.txtWindWaveDirection.Text = "n/d"
            Me.txtWindWaveDirection.SelStart = 3
            KeyAscii = 0
            ShowNoDataDialog "txtWindWaveDirection"
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ��������� ������ Backspace
    If Me.txtWindWaveDirection.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ������ �����
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    ' ��������� ��� ��������� � ����������
    Dim newText As String
    If Me.txtWindWaveDirection.SelLength > 0 Then
        newText = Left(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.SelStart + Me.txtWindWaveDirection.SelLength + 1)
    Else
        newText = Left(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindWaveDirection.Text, Me.txtWindWaveDirection.SelStart + 1)
    End If
    ' ��������� ��� ����� �� ������ 360
    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtBarometer_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ��������� Backspace ������
    If KeyAscii = 8 Then Exit Sub
    
    ' �������� �� ���� "n/d"
    If Me.txtBarometer.Text = "" Or Me.txtBarometer.SelLength = Len(Me.txtBarometer.Text) Then
        If KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
            Me.txtBarometer.Text = "n/d"
            Me.txtBarometer.SelStart = 3
            ShowNoDataDialog "txtBarometer"
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ��������� ������ Backspace
    If Me.txtBarometer.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ����� � �������
    Select Case KeyAscii
        Case 48 To 57 ' �����
            ' �������� ��� ���������� � ����������
            Dim newText As String
            If Me.txtBarometer.SelLength > 0 Then
                newText = Left(Me.txtBarometer.Text, Me.txtBarometer.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtBarometer.Text, Me.txtBarometer.SelStart + Me.txtBarometer.SelLength + 1)
            Else
                newText = Left(Me.txtBarometer.Text, Me.txtBarometer.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtBarometer.Text, Me.txtBarometer.SelStart + 1)
            End If
            
            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 9000 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' ������� ��� �����
            If InStr(Me.txtBarometer.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            KeyAscii = 44 ' ������ �������
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtWindDirection_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ��������� Backspace ������
    If KeyAscii = 8 Then Exit Sub
    
    ' �������� �� ���� "n/d"
    If Me.txtWindDirection.Text = "" Or Me.txtWindDirection.SelLength = Len(Me.txtWindDirection.Text) Then
        If KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
            Me.txtWindDirection.Text = "n/d"
            Me.txtWindDirection.SelStart = 3
            ShowNoDataDialog "txtWindDirection"
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    ' ���� ��� ����' "n/d", ��������� ������ Backspace
    If Me.txtWindDirection.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ������ �����
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ���� �������� 0 � ������ ���� - ��������� ��� ���� ������
    If KeyAscii = 48 And (Me.txtWindDirection.Text = "" Or Me.txtWindDirection.SelLength = Len(Me.txtWindDirection.Text)) Then
        Me.txtWindDirection.Text = "0"
        Me.txtWindSpeed.Text = "0"
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ��� ��������� � ����������
    Dim newText As String
    If Me.txtWindDirection.SelLength > 0 Then
        newText = Left(Me.txtWindDirection.Text, Me.txtWindDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindDirection.Text, Me.txtWindDirection.SelStart + Me.txtWindDirection.SelLength + 1)
    Else
        newText = Left(Me.txtWindDirection.Text, Me.txtWindDirection.SelStart) & Chr(KeyAscii) & _
                 Mid(Me.txtWindDirection.Text, Me.txtWindDirection.SelStart + 1)
    End If
    
    ' �� ���� ������� ����� ������ 360
    If IsNumeric(newText) Then
        If CLng(newText) > 360 Then
            KeyAscii = 0
        End If
    End If
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub txtWindSpeed_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    On Error GoTo ErrorHandler
    
    ' ������. Backspace ������
    If KeyAscii = 8 Then Exit Sub
    
    ' �������� �� ���� "n/d"
    If Me.txtWindSpeed.Text = "" Or Me.txtWindSpeed.SelLength = Len(Me.txtWindSpeed.Text) Then
        If KeyAscii = Asc("n") Or KeyAscii = Asc("N") Then
            Me.txtWindSpeed.Text = "n/d"
            Me.txtWindSpeed.SelStart = 3
            ShowNoDataDialog "txtWindSpeed"
            KeyAscii = 0
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� "n/d", ���������� ������ Backspace
    If Me.txtWindSpeed.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    '���� �������� 0 � ������ ���� - ��������� ��� ����� ������
    If KeyAscii = 48 And (Me.txtWindSpeed.Text = "" Or Me.txtWindSpeed.SelLength = Len(Me.txtWindSpeed.Text)) Then
        Me.txtWindSpeed.Text = "0"
        Me.txtWindDirection.Text = "0"
        KeyAscii = 0
        Exit Sub
    End If
    ' ��������� ����� � �������
    Select Case KeyAscii
        Case 48 To 57 ' �����
            ' ��������� ��� ��������� � ����������
            Dim newText As String
            If Me.txtWindSpeed.SelLength > 0 Then
                newText = Left(Me.txtWindSpeed.Text, Me.txtWindSpeed.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindSpeed.Text, Me.txtWindSpeed.SelStart + Me.txtWindSpeed.SelLength + 1)
            Else
                newText = Left(Me.txtWindSpeed.Text, Me.txtWindSpeed.SelStart) & Chr(KeyAscii) & _
                         Mid(Me.txtWindSpeed.Text, Me.txtWindSpeed.SelStart + 1)
            End If
            ' �� ���� ������� ����� ������ 100
            If IsNumeric(Replace(newText, ",", ".")) Then
                If CDbl(Replace(newText, ",", ".")) > 100 Then
                    KeyAscii = 0
                    Exit Sub
                End If
            End If
            
        Case 44, 46 ' ������� ��� �����
            ' ��������� ���� ��� ���� �������
            If InStr(Me.txtWindSpeed.Text, ",") > 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            ' ��������� � ������ ������
            If Me.txtWindSpeed.SelStart = 0 Then
                KeyAscii = 0
                Exit Sub
            End If
            ' ������ ������ �������
            KeyAscii = 44
            
        Case Else
            KeyAscii = 0
    End Select
    
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub HandleNumericKeyPress(ByVal KeyAscii As MSForms.ReturnInteger, txt As MSForms.TextBox, _
                                Optional MinValue As Double = -9999999, _
                                Optional MaxValue As Double = 9999999)
    On Error GoTo ErrorHandler
    
    ' ������. Backspace ������
    If KeyAscii = 8 Then
        Exit Sub
    End If
    
    ' ��������� ����� n/d
    If (txt.Text = "" Or txt.SelLength = Len(txt.Text)) Then
        If Chr(KeyAscii) = "n" Or Chr(KeyAscii) = "N" Then
            txt.Text = "n/d"
            txt.SelStart = 3
            KeyAscii = 0
            ShowNoDataDialog txt.Name
            Exit Sub
        End If
    End If
    
    ' ���� ��� ���� n/d, ���� ���� ���� ����� Backspace
    If txt.Text = "n/d" Then
        KeyAscii = 0
        Exit Sub
    End If
    
    ' ��������� ������
    If KeyAscii = 45 Then ' �����
        If MinValue >= 0 Or txt.SelStart > 0 Or InStr(txt.Text, "-") > 0 Then
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    ' ��������� �������
    If KeyAscii = 44 Or KeyAscii = 46 Then ' ������� ��� �����
        If InStr(txt.Text, ",") > 0 Then
            KeyAscii = 0
        Else
            KeyAscii = 44 ' ������ �������
        End If
        Exit Sub
    End If
    
    ' ��������� ����
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Dim newText As String
        
        ' ��������� ����� �����
        If txt.SelLength > 0 Then
            newText = Left(txt.Text, txt.SelStart) & Chr(KeyAscii) & _
                     Mid(txt.Text, txt.SelStart + txt.SelLength + 1)
        Else
            newText = Left(txt.Text, txt.SelStart) & Chr(KeyAscii) & _
                     Mid(txt.Text, txt.SelStart + 1)
        End If
        
        ' ��������� �������� ��������
        If newText = "-" Then Exit Sub ' ��������� ���� �����
        
        If IsNumeric(Replace(newText, ",", ".")) Then
            Dim numValue As Double
            numValue = CDbl(Replace(newText, ",", "."))
            If numValue < MinValue Or numValue > MaxValue Then
                KeyAscii = 0
            End If
        Else
            KeyAscii = 0
        End If
        Exit Sub
    End If
    
    ' ��� ��������� ������� ���������
    KeyAscii = 0
    Exit Sub

ErrorHandler:
    KeyAscii = 0
End Sub
Private Sub ShowNoDataDialog(FieldName As String)
    On Error GoTo ErrorHandler
    
    Dim Reason As String
    Dim fieldLabel As String
    
    ' �������� �������� �������� ����
    Select Case FieldName
        Case "txtTemp": fieldLabel = "�����������"
        Case "txtBarometer": fieldLabel = "��������"
        Case "txtVisibility": fieldLabel = "���������"
        Case "txtWindDirection": fieldLabel = "����������� �����"
        Case "txtWindSpeed": fieldLabel = "�������� �����"
        Case "txtSeaSwellDirection": fieldLabel = "����������� ��������"
        Case "txtSeaSwell": fieldLabel = "������ ��������"
        Case "txtWindWaveDirection": fieldLabel = "����������� �������� ����"
        Case "txtWindWaveHeight": fieldLabel = "������ �������� ����"
        Case Else: fieldLabel = FieldName
    End Select
    
    Reason = InputBox("������� ������� ���������� ������ ��� ���� '" & fieldLabel & "'" & _
                     vbNewLine & vbNewLine & "��������: ������� ������ ���� ���������� �� �������������!", _
                     "������� ���������� ������")
    
    If Reason = "" Then
        Me.Controls(FieldName).Text = ""
    Else
        SaveReason FieldName, Reason
    End If
    Exit Sub

ErrorHandler:
    MsgBox "��������� ������ ��� ����� �������. ���������� ��� ���.", vbExclamation
End Sub
Private Sub SaveReason(FieldName As String, Reason As String)
    On Error GoTo ErrorHandler
    
    Dim ws As Worksheet
    
    ' ��������/�������� ����� Reasons
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Reasons")
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Reasons"
        
        With ws
            .Cells(1, 1) = "����/�����"
            .Cells(1, 2) = "����"
            .Cells(1, 3) = "�������"
            
            With .Range("A1:C1")
                .Font.Bold = True
                .Interior.Color = RGB(220, 220, 220)
            End With
            
            .Columns("A").ColumnWidth = 20
            .Columns("B").ColumnWidth = 25
            .Columns("C").ColumnWidth = 50
        End With
    End If
    
    ' �����. ����� ������
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    
    With ws
        .Cells(lastRow, 1) = Now
        .Cells(lastRow, 2) = FieldName
        .Cells(lastRow, 3) = Reason
        .Cells(lastRow, 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    End With
    
    Exit Sub
ErrorHandler:
    MsgBox "Error while loading data: / ������ ��� ����������: " & Err.Description, vbExclamation
End Sub
Private Sub SetDefaultValues()
    If Me.Tag = "New" Then
        Dim currentTime As Date
        currentTime = Now
        If Minute(currentTime) > 30 Then
            currentTime = DateAdd("h", 1, currentTime)
        End If
        
        Me.txtDateTime1.value = Format(DateSerial(Year(currentTime), Month(currentTime), day(currentTime)) + _
                           Hour(currentTime) / 24, "dd.mm.yyyy hh:00")
    End If
End Sub



Private Sub FindAndSelectComboValue(cmb As MSForms.ComboBox, value As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i, 0) = value Then
            cmb.ListIndex = i
            Exit For
        End If
    Next i
End Sub
Private Sub LoadSeaSwellData(ByVal rowNum As Long, ByVal dataSheet As Worksheet)
    With Me
        .txtSeaSwell.value = Format(dataSheet.Cells(rowNum, 8).value, "0.0")
        .txtSeaSwellDirection.value = dataSheet.Cells(rowNum, 9).value
        .txtWindWaveDirection.value = dataSheet.Cells(rowNum, 10).value
        .txtWindWaveHeight.value = Format(dataSheet.Cells(rowNum, 11).value, "0.0")
        .cboIceScore.value = dataSheet.Cells(rowNum, 12).value  ' Ice score
        .cboIceType.value = dataSheet.Cells(rowNum, 13).value   ' Ice type
        .cboIceShape.value = dataSheet.Cells(rowNum, 14).value  ' Ice shape

    End With
End Sub
Private Sub LoadIceData(ByVal rowNum As Long, ByVal dataSheet As Worksheet)
    With Me
        .cboIceScore.value = dataSheet.Cells(rowNum, "L").value  ' Ice score
        
        ' ��� Ice Type � Shape ���� ��������������� ��������
        Dim iceTypeValue As String, iceShapeValue As String
        iceTypeValue = dataSheet.Cells(rowNum, "M").value
        iceShapeValue = dataSheet.Cells(rowNum, "N").value
        
        ' ���� � ������������� �������� � �������
        FindAndSelectComboValue .cboIceType, iceTypeValue
        FindAndSelectComboValue .cboIceShape, iceShapeValue
    End With
End Sub
' ��������������� ������� ��� ��������� �������� � ComboBox �� ������� B
Private Sub SetComboBoxValueByColumn2(cmb As MSForms.ComboBox, value As String)
    Dim i As Long
    For i = 0 To cmb.ListCount - 1
        If cmb.List(i, 1) = value Then
            cmb.ListIndex = i
            Exit For
        End If
    Next i
End Sub
Private Sub cmdSave_Click()
    On Error GoTo ErrorHandler
    
    If Not ValidateData Then Exit Sub
    
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Data")
    
    ' �������� ������� ������ ����� ����� �����������
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    ' �������� ������� ������
    Dim targetRow As Long
    If Me.Tag = "New" Then
        targetRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
    Else
        targetRow = CLng(Me.Tag)
    End If
    
    ' ��������� ������
    SaveDataToSheet ws, targetRow
    
    ' ��������������� ������ �����
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    
    ' ��������� ����� ��� ��������������� ���������
    Unload Me
    Exit Sub

ErrorHandler:
    MsgBox "Error while saving data: " & vbNewLine & Err.Description, vbCritical
    
    ' ��������, ��� ������ ������������� ���� � ������ ������
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
Private Sub SaveDataToSheet(ByRef ws As Worksheet, ByVal targetRow As Long)
    On Error GoTo ErrorHandler
    
    ' �������� ������� ������ (�� ������ ������)
    On Error Resume Next
    ws.Unprotect
    On Error GoTo 0
    
    With ws
      ' Date/Time
        .Cells(targetRow, 1) = CDate(Me.txtDateTime1.value)
        
        ' Latitude
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            .Cells(targetRow, 2) = CDbl(Replace(Me.txtLatitude.value, ".", ","))
        Else
            .Cells(targetRow, 2) = GetDecimalCoordinates(LatitudeInput.degrees.Text, _
                                                       LatitudeInput.minutes.Text, _
                                                       LatitudeInput.direction.Text)
        End If
        
        ' Longitude
        If mCoordFormat = COORD_FORMAT_DECIMAL Then
            .Cells(targetRow, 3) = CDbl(Replace(Me.txtLongitude.value, ".", ","))
        Else
            .Cells(targetRow, 3) = GetDecimalCoordinates(LongitudeInput.degrees.Text, _
                                                       LongitudeInput.minutes.Text, _
                                                       LongitudeInput.direction.Text)
        End If
        
        ' Temperature
        If Me.txtTemp.Text = "n/d" Then
            .Cells(targetRow, 4) = Me.txtTemp.Text
        Else
            .Cells(targetRow, 4) = CDbl(Replace(Me.txtTemp.value, ".", ","))
        End If
        
        ' Barometer, mm
        If Me.txtBarometer.Text = "n/d" Then
            .Cells(targetRow, 5) = Me.txtBarometer.Text
        Else
            .Cells(targetRow, 5) = CDbl(Replace(Me.txtBarometer.value, ".", ","))
        End If
        
        ' Visibility, m
        If Me.txtVisibility.Text = "n/d" Then
            .Cells(targetRow, 6) = Me.txtVisibility.Text
        Else
            .Cells(targetRow, 6) = CDbl(Replace(Me.txtVisibility.value, ".", ","))
        End If
        
        ' Wind Direction, degree
        If Me.txtWindDirection.Text = "0" And Me.txtWindSpeed.Text = "0" Then
            .Cells(targetRow, 7) = "0"
            .Cells(targetRow, 8) = "0"
        Else
            If Me.txtWindDirection.Text = "n/d" Then
                .Cells(targetRow, 7) = Me.txtWindDirection.Text
            Else
                .Cells(targetRow, 7) = CDbl(Me.txtWindDirection.value)
            End If
            
            ' Wind SpeedAVG, m/s
            If Me.txtWindSpeed.Text = "n/d" Then
                .Cells(targetRow, 8) = Me.txtWindSpeed.Text
            Else
                .Cells(targetRow, 8) = CDbl(Replace(Me.txtWindSpeed.value, ".", ","))
            End If
        End If
        
        ' Sea conditions
        If Me.chkSeaSwell.value Then
            ' Sea Swell Direction, degree
            If Me.txtSeaSwellDirection.Text = "n/d" Then
                .Cells(targetRow, 9) = Me.txtSeaSwellDirection.Text
            Else
                .Cells(targetRow, 9) = CDbl(Me.txtSeaSwellDirection.value)
            End If
            
            ' Sea Swell, m
            If Me.txtSeaSwell.Text = "n/d" Then
                .Cells(targetRow, 10) = Me.txtSeaSwell.Text
            Else
                .Cells(targetRow, 10) = CDbl(Replace(Me.txtSeaSwell.value, ".", ","))
            End If
            
            ' Wind wave direction, degree
            If Me.txtWindWaveDirection.Text = "n/d" Then
                .Cells(targetRow, 11) = Me.txtWindWaveDirection.Text
            Else
                .Cells(targetRow, 11) = CDbl(Me.txtWindWaveDirection.value)
            End If
            
            ' Wind wave height, m
            If Me.txtWindWaveHeight.Text = "n/d" Then
                .Cells(targetRow, 12) = Me.txtWindWaveHeight.Text
            Else
                .Cells(targetRow, 12) = CDbl(Replace(Me.txtWindWaveHeight.value, ".", ","))
            End If
        Else
            .Cells(targetRow, 9) = "0"   ' Sea Swell Direction
            .Cells(targetRow, 10) = "0"  ' Sea Swell
            .Cells(targetRow, 11) = "0"  ' Wind wave direction
            .Cells(targetRow, 12) = "0"  ' Wind wave height
        End If
        
        ' Ice Conditions
        If Me.chkIceNotated.value Then
            ' Ice score
            If Me.cboIceScore.ListIndex <> -1 Then
                .Cells(targetRow, 13) = Me.cboIceScore.List(Me.cboIceScore.ListIndex, 0)
            Else
                .Cells(targetRow, 13) = "�� ������������ ��� ����������"
            End If
            
            ' Ice type
            If Me.cboIceType.ListIndex <> -1 Then
                .Cells(targetRow, 14) = Me.cboIceType.List(Me.cboIceType.ListIndex, 0)
            Else
                .Cells(targetRow, 14) = "�� ������������ ��� ����������"
            End If
            
            ' Ice shape
            If Me.cboIceShape.ListIndex <> -1 Then
                .Cells(targetRow, 15) = Me.cboIceShape.List(Me.cboIceShape.ListIndex, 0)
            Else
                .Cells(targetRow, 15) = "�� ������������ ��� ����������"
            End If
        Else
            .Cells(targetRow, 13) = "������ ����"      ' Ice score
            .Cells(targetRow, 14) = "������ ����"      ' Ice type
            .Cells(targetRow, 15) = "�� ������������ ��� ����������" ' Ice shape
        End If
        
        ' ������� �������������� (��� .Select)
        On Error Resume Next
        With .Range(.Cells(targetRow, 1), .Cells(targetRow, 15))
            .Borders.LineStyle = xlContinuous
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
       End With
    
    MsgBox "Data saved successfully! / ������ ������� ���������!", vbInformation
    
    ' ��������������� ������
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
    
    Exit Sub

ErrorHandler:
    MsgBox "Error while saving data / ������ ��� ���������� ������" & vbNewLine & _
           "Error Description: " & Err.Description & vbNewLine & _
           "Error Number: " & Err.Number & vbNewLine & _
           "Error Source: " & Err.Source, vbCritical
    
    ' ��������������� ������ ���� � ������ ������
    On Error Resume Next
    ws.Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub
Private Sub UserForm_Terminate()
    ' ��������, ��� ���� ������� ��� �������� �����
    On Error Resume Next
    ThisWorkbook.Sheets("Data").Protect UserInterfaceOnly:=True
    On Error GoTo 0
End Sub

Private Sub cboIceScore_DropDown()
    EnableMouseWheel Me.cboIceScore
End Sub
Private Sub cboIceType_DropDown()
    EnableMouseWheel Me.cboIceType
End Sub
Private Sub cboIceShape_DropDown()
    EnableMouseWheel Me.cboIceShape
End Sub
Private Sub EnableMouseWheel(cmb As MSForms.ComboBox)
    Dim hwndList As LongPtr
    hwndList = FindWindowEx(cmb.hwnd, 0, "ComboBox", vbNullString)
    If hwndList <> 0 Then
        SendMessage hwndList, WM_MOUSEWHEEL, 0, 0
    End If
End Sub
Private Sub cboIceScore_GotFocus()
    SendKeys "{F4}", True
    SendKeys "{F4}", True
End Sub
Private Sub cboIceType_GotFocus()
    SendKeys "{F4}", True
    SendKeys "{F4}", True
End Sub
Private Sub cboIceShape_GotFocus()
    SendKeys "{F4}", True
    SendKeys "{F4}", True
End Sub
Private Sub cboIceShape_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.cboIceShape.SetFocus
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub



