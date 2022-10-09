VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uMouseRecorder 
   Caption         =   "Mouse Macro"
   ClientHeight    =   4896
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8280.001
   OleObjectBlob   =   "uMouseRecorder.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uMouseRecorder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uMouseRecorder
'* Created    : 06-10-2022 10:38
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Completed As Boolean

Private Sub UserForm_Initialize()
    Me.Height = 125
    Me.Width = 230
    LoadPosition Me
    UserformOnTop Me
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    MouseFolder = Environ("USERprofile") & "\Documents\vbArc\MouseMacro\"
    FoldersCreate MouseFolder
    checkFile
    ClicksOnly.Value = ws.Range("h2")
    LoadMRcaption
    LoadListbox

    CreateListboxHeader lBoxData, lBoxHeader, Array("X", "Y", "L", "R", "NOTE")
End Sub

Function CursorPosition() As Variant
    Dim lngCurPos As POINTAPI, activeX As Long, activeY As Long
    GetCursorPos lngCurPos
    activeX = lngCurPos.X
    activeY = lngCurPos.Y
    Dim out(1) As Variant
    out(0) = activeX
    out(1) = activeY
    CursorPosition = out
End Function

Sub ShowCoordinates(X As Long, Y As Long)
    uCoordinates.Load
    uCoordinates.left = X
    uCoordinates.top = Y
    uCoordinates.Show
End Sub

Private Sub iLogLink_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogLink
    LoadListbox
End Sub

Sub LogLink()
    '#INCLUDE InputboxString
    '#INCLUDE IsFileFolderURL
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim msg As String
    msg = InputboxString()
    If Len(msg) = 0 Then Exit Sub
    If msg = "False" Then Exit Sub
    If IsFileFolderURL(msg) = "I" Then Exit Sub
    Dim cell As Range
    Set cell = ws.Range("A" & rows.count).End(xlUp).OFFSET(1)
    cell = "go"
    cell.OFFSET(0, 1) = msg
End Sub

Private Sub iLogLink_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLogLink.ControlTipText
End Sub

Private Sub iLogRight_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogClick "right"
    LoadListbox
End Sub

Private Sub iCoordinates_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    IndexMouseLocation
End Sub

Private Sub iCoordinates_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iCoordinates.ControlTipText
End Sub

Private Sub iSize_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Me.Height > 125 Then
        Me.Height = 125
        iSize.SpecialEffect = fmSpecialEffectRaised
    Else
        Me.Height = 275
        iSize.SpecialEffect = fmSpecialEffectSunken
        CreateListboxHeader lBoxData, lBoxHeader, Array("X", "Y", "L", "R", "NOTE")
    End If
End Sub

Private Sub iSize_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iSize.ControlTipText
End Sub

Private Sub lBoxData_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = lBoxData.ControlTipText
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = "Hold ESC to STOP recording or playback"
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SavePosition Me
End Sub

Private Sub ClicksOnly_Click()
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    ws.Range("H2") = ClicksOnly
End Sub

Sub DeleteRows()
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    rng.Delete Shift:=xlUp
End Sub

Sub DoubleClick()
    'Double click as a quick series of two clicks
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Sub DuplicateRows()
    '#INCLUDE LoadListbox
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    Dim var
    var = rng.Value
    rng.OFFSET(rng.rows.count).Insert
    rng.OFFSET(rng.rows.count).RESIZE(rng.rows.count) = var
    Application.CutCopyMode = False
    LoadListbox
End Sub

Sub EditMemo()
    '#INCLUDE LoadListbox
    '#INCLUDE InputboxString
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    Dim s As String
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ans As String
    ans = InputboxString
    If ans = "False" Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Cells(2 + r, 5).RESIZE(c)
    rng.Value = ans
    LoadListbox
End Sub

Sub EditRow()
    '#INCLUDE LoadListbox
    '#INCLUDE InputboxString
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    Dim s As String
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    s = lBoxData.list(lBoxData.ListIndex, 0)
    s = s & "|" & lBoxData.list(lBoxData.ListIndex, 1)
    s = s & "|" & lBoxData.list(lBoxData.ListIndex, 2)
    s = s & "|" & lBoxData.list(lBoxData.ListIndex, 3)
    s = s & "|" & lBoxData.list(lBoxData.ListIndex, 4)
    Dim ans As String
    ans = InputboxString(, , s)
    If ans = "False" Then Exit Sub
    If UBound(Split(s, "|")) <> 4 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    rng.Value = (Split(ans, "|"))
    LoadListbox
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Sub LeftClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

'Private Sub iMove_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'If Button = 1 Then
'    m_sngDownX = X
'    m_sngDownY = Y
'End If
'End Sub
'Private Sub iMove_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
'If Button And 1 Then
'    Me.left = Me.left + (X - m_sngDownX)
'    Me.top = Me.top + (Y - m_sngDownY)
'End If
'End Sub
Sub LoadListbox()
    '#INCLUDE RecordRange
    Dim rng As Range
    Set rng = RecordRange
    lBoxData.clear
    If rng Is Nothing Then Exit Sub
    lBoxData.columnCount = rng.Columns.count
    lBoxData.list = rng.Value
End Sub

Sub LoadMRcaption()
    '#INCLUDE RecordFileFullName
    '#INCLUDE FileExists
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim FileFullPath As String
    FileFullPath = RecordFileFullName
    If ws.Range("H1") = "" Then
        If ws.Range("A2") = "" Then
            Me.Caption = "New Recording"
        ElseIf ws.Range("A2") <> "" Then
            Me.Caption = "Existing Recording - NOT SAVED"
        End If
    ElseIf ws.Range("H1") <> "" Then
        Me.Caption = IIf(FileExists(FileFullPath), "Loaded - " & ws.Range("H1"), "New Recording")
    End If
End Sub

Sub LoadRecord()
    '#INCLUDE TXTtoArray
    '#INCLUDE PickRecord
    '#INCLUDE RecordFileFullName
    '#INCLUDE newRecord
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim FName As String
    FName = PickRecord(MouseFolder)
    If FName = "" Or Right(FName, 7) <> "_mr.txt" Then
        infoLab.Caption = "No valid file selected"
        Exit Sub
    End If
    newRecord
    FName = Mid(FName, InStrRev(FName, "\") + 1)
    FName = left(FName, InStr(1, FName, "_") - 1)
    uMouseRecorder.LoadedRecording.Caption = FName
    ws.Range("H1") = FName
    Dim recFile As String
    recFile = RecordFileFullName
    Dim arr
    arr = TXTtoArray(recFile)
    If IsEmpty(arr) Then Exit Sub
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion.OFFSET(1)
    rng.ClearContents
    rng.RESIZE(UBound(arr, 1), 4) = arr
End Sub

'VBA function to open a CSV file in memory and parse it to a 2D
'array without ever touching a worksheet:

Function TXTtoArray(sFile$)
    '#INCLUDE OpenTextFile
    Dim c&, i&, j&, p&, d$, s$, rows&, cols&, a, r, v
    Const Q = """", QQ = Q & Q
    Const ENQ = ""        'Chr(5)
    Const ESC = ""        'Chr(27)
    Const COM = ","
    
    d = OpenTextFile$(sFile)
    If LenB(d) Then
        r = Split(Trim(d), vbCrLf)
        rows = UBound(r) + 1
        cols = UBound(Split(r(0), ",")) + 1
        ReDim v(1 To rows, 1 To cols)
        For i = 1 To rows
            s = r(i - 1)
            If LenB(s) Then
                If InStrB(s, QQ) Then s = Replace(s, QQ, ENQ)
                For p = 1 To Len(s)
                    Select Case Mid(s, p, 1)
                        Case Q:   c = c + 1
                        Case COM: If c Mod 2 Then Mid(s, p, 1) = ESC
                    End Select
                Next
                If InStrB(s, Q) Then s = Replace(s, Q, "")
                a = Split(s, COM)
                For j = 1 To cols
                    s = a(j - 1)
                    If InStrB(s, ESC) Then s = Replace(s, ESC, COM)
                    If InStrB(s, ENQ) Then s = Replace(s, ENQ, Q)
                    v(i, j) = s
                Next
            End If
        Next
        TXTtoArray = v
    End If
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
        .Close
    End With
End Function

Private Sub LocMouse_Click()
    PreviewMousePosition
End Sub

Sub LogAsk()
    '#INCLUDE InputboxString
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim cell As Range
    Set cell = ws.Range("A" & rows.count).End(xlUp).OFFSET(1)
    Dim msg As String
    msg = InputboxString()
    If Len(msg) = 0 Then Exit Sub
    If msg = "False" Then Exit Sub
    cell = "ask"
    cell.OFFSET(0, 1) = msg
End Sub

Sub IndexMouseLocation()
    Dim lngCurPos As POINTAPI
    Dim activeX As Long, activeY As Long
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        GetCursorPos lngCurPos
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        LabX.TEXT = activeX
        LabY.TEXT = activeY
        Sleep 20
        DoEvents
    Loop
LoopEnd:
    Application.EnableCancelKey = xlInterrupt
End Sub

Sub LogClick(ClickType As String)
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Erase MouseArray
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim activeX As Long, activeY As Long
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        GetCursorPos lngCurPos
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        LabX.TEXT = activeX
        LabY.TEXT = activeY
        Sleep 20
        DoEvents
    Loop
LoopEnd:
    'If err = 18 Then
    Application.EnableCancelKey = xlInterrupt
    Set rng = ws.Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
    Set rng = rng.RESIZE(, 5)
    rng.Value = Array(ClickType, activeX, activeY, "", "")
        
    infoLab.Caption = "Macro recorded."
    'End If
End Sub

Sub LogClickImmediate()
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Erase MouseArray
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim previousX As Long, previousY As Long, activeX As Long, activeY As Long
    Dim previousL As Long, previousR As Long, activeL As Long, activeR As Long
    Dim arrayCounter As Long: arrayCounter = 1
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Dim counter As Long
    Do
        ReDim Preserve MouseArray(1 To arrayCounter)
        GetCursorPos lngCurPos
        activeL = IIf(GetAsyncKeyState(1) = 0, 0, 1)
        activeR = IIf(GetAsyncKeyState(2) = 0, 0, 1)
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        If previousL <> activeL Or previousR <> activeR Then
            previousX = activeX
            previousY = activeY
            previousL = activeL
            previousR = activeR
            MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
            arrayCounter = arrayCounter + 1
            DoEvents
            counter = counter + 1
            If counter = 4 Then GoTo LoopEnd
        End If
    Loop
LoopEnd:
    If err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Set rng = ws.Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
        Set rng = rng.RESIZE(UBound(MouseArray), 1)
        rng = WorksheetFunction.Transpose(MouseArray)
        rng.TextToColumns rng, comma:=True
        Range(rng.Cells(1, 1), rng.Cells(2, 4)).Delete Shift:=xlUp
        'ws.Range("A3:D3").Delete Shift:=xlUp
        infoLab.Caption = "Macro recorded."
        '        infoLab.Caption = "Macro recorded at rows: " & rng.Row & " to " & rng.Row + rng.Rows.Count
        Exit Sub
    End If
End Sub

Sub LogDoulbe()
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Erase MouseArray
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim activeX As Long, activeY As Long
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        GetCursorPos lngCurPos
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        LabX.TEXT = activeX
        LabY.TEXT = activeY
        Sleep 20
        DoEvents
    Loop
LoopEnd:
    If err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Set rng = ws.Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
        Set rng = rng.RESIZE(5)
        rng.Value = Array("double", activeX, activeY, "", "")
        infoLab.Caption = "Macro recorded."
        '        infoLab.Caption = "Macro recorded at rows: " & rng.Row & " to " & rng.Row + rng.Rows.Count
        Exit Sub
    End If
End Sub

Sub LogText()
    '#INCLUDE InputboxString
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim cell As Range
    Set cell = ws.Range("A" & rows.count).End(xlUp).OFFSET(1)
    Dim msg As String
    msg = InputboxString()
    If Len(msg) = 0 Then Exit Sub
    If msg = "False" Then Exit Sub
    cell = "sendkeys"
    cell.OFFSET(0, 1) = msg
End Sub

Sub MouseReplay(Optional rng As Range)
    '#INCLUDE DoubleClick
    '#INCLUDE LeftClick
    '#INCLUDE RightClick
    '#INCLUDE dragMouse
    '#INCLUDE FollowLink
    '#INCLUDE InputboxString
    '#INCLUDE CLIP
    '#INCLUDE IsFileFolderURL
    Completed = False
    'ActiveWindow.WindowState = xlMaximized
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim cell As Range
    If rng Is Nothing Then
        Set rng = ws.Range("A2").CurrentRegion
        Set rng = rng.OFFSET(1).RESIZE(rng.rows.count - 1, 1)
    End If
    If WorksheetFunction.CountA(rng) = 0 Then Exit Sub
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Dim DefaultSleep As Long
    DefaultSleep = 300
    Dim msg As String

    For Each cell In rng
        
        Rem if automatic record of clicks and motion
        If IsNumeric(cell) Then

            SetCursorPos cell, cell.OFFSET(, 1)
            
            If cell.OFFSET(0, 2) > 1 Then
                dragMouse cell.Value, cell.OFFSET(0, 1), cell.OFFSET(0, 2), cell.OFFSET(0, 3)
            ElseIf cell.OFFSET(0, 2) = 1 Then
                If cell.OFFSET(1, 2) = 0 Then
                    '                    If cell.Offset(2, 2) = 0 And cell.Offset(-1, 2) = 0 Then
                    LeftClick
                    Set cell = cell.OFFSET(2, 0)
                    Rem This way doesn't work if logging clicks only and not motion because two left clicks will be interpreted as double click
                    '                    ElseIf cell.Offset(2, 2) = 1 Then
                    '                        DoubleClick
                    '                        Set cell = cell.Offset(2, 0)
                    '                    End If
                End If
            ElseIf cell.OFFSET(0, 3) = 1 Then
                RightClick
            End If
        Else
            Rem if manual entry
            If cell = "wait" Then
                Sleep IIf(cell.OFFSET(0, 1) <> "", cell.OFFSET(0, 1), DefaultSleep)
            ElseIf cell = "go" Then
                msg = Replace(cell.OFFSET(0, 1), """", "")
                If IsFileFolderURL(msg) <> "I" Then
                    FollowLink msg
                    Sleep 500
                End If
            ElseIf cell = "move" Then
                SetCursorPos cell.OFFSET(0, 1), cell.OFFSET(0, 2)
            ElseIf cell = "left" Then
                SetCursorPos cell.OFFSET(0, 1), cell.OFFSET(0, 2)
                LeftClick
            ElseIf cell = "right" Then
                SetCursorPos cell.OFFSET(0, 1), cell.OFFSET(0, 2)
                RightClick
            ElseIf cell = "double" Then
                SetCursorPos cell.OFFSET(0, 1), cell.OFFSET(0, 2)
                DoubleClick
            ElseIf cell = "drag" Then
                dragMouse cell.OFFSET(0, 1), cell.OFFSET(0, 2), cell.OFFSET(0, 3), cell.OFFSET(0, 4)
            ElseIf cell = "ask" Then
                msg = InputboxString(0, cell.OFFSET(0, 1))
                If Len(msg) > 0 Then
                    CLIP msg
                    SendKeys CLIP, True
                End If
            ElseIf cell = "sendkeys" Then
                Dim ClipText As String
                ClipText = IIf(cell.OFFSET(0, 2) = "", cell.OFFSET(0, 1), String(cell.OFFSET(0, 2), cell.OFFSET(0, 1)))
                CLIP ClipText
                SendKeys CLIP, True

            End If
        End If
        
        DoEvents
        Sleep 20        'DefaultSleep
        If Completed Then Exit Sub
    Next
LoopEnd:
    '    If err = 18 Then
    Application.EnableCancelKey = xlInterrupt
    Do While GetAsyncKeyState(1) = 1
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        DoEvents
    Loop
    Completed = True
    '    End If
End Sub

Sub MoveRows(offsetRows As Long)
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    On Error Resume Next        ' in case user makes unreasonable action like only 1 row exists and try to move it
    rng.Cut
    If 2 + r + offsetRows < 2 Then
        ws.Range("A2:E2").Insert
    ElseIf 2 + r + offsetRows > ws.Range("A1").CurrentRegion.rows.count Then
        Dim lRow As Long
        lRow = ws.Range("A1").CurrentRegion.rows.count
        ws.Range("A" & lRow).RESIZE(, 5).Insert
    Else
        rng.OFFSET(offsetRows).Insert
    End If
    Application.CutCopyMode = False
End Sub

Sub MoveToBottom()
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If lBoxData.ListIndex = -1 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    On Error Resume Next        ' in case user makes unreasonable action like only 1 row exists and try to move it
    rng.Cut
    Dim lRow As Long
    lRow = ws.Range("A1").CurrentRegion.rows.count + 1
    ws.Range("A" & lRow).RESIZE(, 5).Insert
    Application.CutCopyMode = False
End Sub

Sub MoveToTop()
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If lBoxData.ListIndex = -1 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 5)).RESIZE(c)
    On Error Resume Next        ' in case user makes unreasonable action like only 1 row exists and try to move it
    rng.Cut
    ws.Range("A2").RESIZE(, 5).Insert
    Application.CutCopyMode = False
End Sub

Function PickRecord(Optional initFolder As String) As String
    If initFolder = "" Then initFolder = MouseFolder
    Dim strFile As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.clear
        .Filters.Add "MouseRecord", "*.txt"
        .title = "Choose Mouse Record"
        .AllowMultiSelect = False
        .initialFileName = initFolder
        If .Show = True Then
            strFile = .SelectedItems(1)
            PickRecord = strFile
        End If
    End With
End Function

Sub PlayBackSelectedRows()
    '#INCLUDE MouseReplay
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Cells(2 + r, 1).RESIZE(c)
    MouseReplay rng
End Sub

Sub PlayFromHere()
    '#INCLUDE MouseReplay
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = lBoxData.ListCount - r
    Set rng = ws.Cells(2 + r, 1).RESIZE(c)
    MouseReplay rng
End Sub

Sub PlayUntilHere()
    '#INCLUDE MouseReplay
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If ListboxSelectedCount(uMouseRecorder.lBoxData) = 0 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Set rng = ws.Cells(2, 1).RESIZE(r)
    MouseReplay rng
End Sub

Sub PreviewMousePosition()
    '#INCLUDE ListboxSelectedCount
    '#INCLUDE ListboxSelectedIndexes
    If lBoxData.ListIndex = -1 Then Exit Sub
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim r As Long
    r = ListboxSelectedIndexes(lBoxData)(1)
    Dim c As Long
    c = ListboxSelectedCount(lBoxData)
    Set rng = ws.Range(ws.Cells(2 + r, 1), ws.Cells(2 + r, 2))
    SetCursorPos rng.Cells(1, 1), rng.Cells(1, 2)
End Sub

Function RecordFileFullName() As String
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    RecordFileFullName = MouseFolder & ws.Range("H1") & "_mr.txt"
End Function

Function RecordRange() As Range
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    If ws.Range("A2") = "" Then Exit Function
    Dim rng As Range
    Set rng = ws.Range("A1").CurrentRegion
    Set rng = rng.OFFSET(1).RESIZE(rng.rows.count - 1, 5)
    Set RecordRange = rng
End Function

Sub RecordStart(Optional recordWholeMotion As Boolean)
    'ActiveWindow.WindowState = xlMaximized
    '#INCLUDE LoadMRcaption
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim previousX As Long, previousY As Long, activeX As Long, activeY As Long
    Dim previousL As Long, previousR As Long, activeL As Long, activeR As Long
    Erase MouseArray
    Dim arrayCounter As Long: arrayCounter = 1
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        ReDim Preserve MouseArray(1 To arrayCounter)
        GetCursorPos lngCurPos
        activeL = IIf(GetAsyncKeyState(1) = 0, 0, 1)
        activeR = IIf(GetAsyncKeyState(2) = 0, 0, 1)
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        If recordWholeMotion Then
            If previousX <> lngCurPos.X Or previousY <> lngCurPos.Y Or previousL <> activeL Or previousR <> activeR Then
                previousX = activeX
                previousY = activeY
                previousL = activeL
                previousR = activeR
                MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
                arrayCounter = arrayCounter + 1
                DoEvents
            End If
        Else
            If previousL <> activeL Or previousR <> activeR Then
                previousX = activeX
                previousY = activeY
                previousL = activeL
                previousR = activeR
                MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
                arrayCounter = arrayCounter + 1
                DoEvents
            End If
            LabX.TEXT = activeX
            LabY.TEXT = activeY
        End If
    Loop
LoopEnd:
    If err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Set rng = ws.Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
        Set rng = rng.RESIZE(UBound(MouseArray), 1)
        rng = WorksheetFunction.Transpose(MouseArray)
        rng.TextToColumns rng, comma:=True
        rng.Columns(1).Cells.Font.Bold = False
        rng.Cells.Font.Bold = True
        infoLab.Caption = "Macro recorded at rows: " & rng.row & " to " & rng.row + rng.rows.count - 3
        Range(rng.Cells(1, 1), rng.Cells(2, 4)).Delete Shift:=xlUp
        LoadMRcaption
    End If
End Sub

Function RecordedMacro() As String
    '#INCLUDE ArrayToString
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion.OFFSET(1)
    Set rng = rng.RESIZE(rng.rows.count - 1)
    Dim arr
    arr = rng.Value
    RecordedMacro = ArrayToString(arr)
End Function

'RETURNS A STRING FROM A 2 DIM ARRAY, SPERATED BY OPTIONAL DELIMITER AND VBNEWLINE FOR EACH ROW
'
'@AUTHOR ROBERT TODAR
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    '#INCLUDE ArrayDimensionLength
    
    Dim temp As String
    
    Select Case ArrayDimensionLength(SourceArray)
            'SINGLE DIMENTIONAL ARRAY
        Case 1
            temp = Join(SourceArray, Delimiter)
        
            '2 DIMENSIONAL ARRAY
        Case 2
            Dim RowIndex As Long
            Dim ColIndex As Long
            
            'LOOP EACH ROW IN MULTI ARRAY
            For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                
                'LOOP EACH COLUMN ADDING VALUE TO STRING
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    temp = temp & SourceArray(RowIndex, ColIndex)
                    If ColIndex <> UBound(SourceArray, 2) Then temp = temp & Delimiter
                Next ColIndex
                
                'ADD NEWLINE FOR THE NEXT ROW (MINUS LAST ROW)
                If RowIndex <> UBound(SourceArray, 1) Then temp = temp & vbNewLine
        
            Next RowIndex
    End Select
    
    ArrayToString = temp
    
End Function

'RETURNS THE LENGHT OF THE DIMENSION OF AN ARRAY
Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    
    Dim i As Integer
    Dim test As Long

    On Error GoTo catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
    
catch:
    ArrayDimensionLength = i - 1

End Function

Sub RightClick()
    'Right click
    mouse_event MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0
End Sub

Sub SaveRecord()
    '#INCLUDE RecordFileFullName
    '#INCLUDE RecordedMacro
    '#INCLUDE txtoverwrite
    '#INCLUDE TxtOverwrite
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Set rng = ws.Range("A2").CurrentRegion
    Set rng = rng.OFFSET(1).RESIZE(rng.rows.count - 1)
    TxtOverwrite RecordFileFullName, RecordedMacro
End Sub

Sub checkFile()
    '#INCLUDE RecordFileFullName
    '#INCLUDE FileExists
    Dim recFile As String
    recFile = RecordFileFullName
    Dim recFileName As String
    recFileName = IIf(FileExists(recFile) = True, recFile, "NONE")
    LoadedRecording.Caption = recFileName
    Me.LoadedRecording.ControlTipText = Mid(recFileName, InStrRev(recFileName, "\") + 1)
End Sub

Sub dragMouse(x0 As Long, y0 As Long, X1 As Long, Y1 As Long)
    SetCursorPos x0, y0
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    Sleep 20
    SetCursorPos X1, Y1
    Sleep 20
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Private Sub iBottom_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveToBottom
    LoadListbox
End Sub

Private Sub iBottom_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iBottom.ControlTipText
End Sub

Private Sub iDelete_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    DeleteRows
    LoadListbox
End Sub

Private Sub iDelete_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iDelete.ControlTipText
End Sub

Private Sub iDown_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveRows 2
    LoadListbox
End Sub

Private Sub iDown_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iDown.ControlTipText
End Sub

Private Sub iDuplicate_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    DuplicateRows
End Sub

Private Sub iDuplicate_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iDuplicate.ControlTipText
End Sub

Private Sub iFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink MouseFolder
End Sub

Private Sub iFolder_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iFolder.ControlTipText
End Sub

Private Sub iLoadRecord_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LoadRecord
    LoadedRecording.ControlTipText = LoadedRecording.Caption
    Dim s As String
    s = LoadedRecording.Caption
    s = Mid(s, InStrRev(s, "\") + 1)
    Me.Caption = s
End Sub

Private Sub iLoadRecord_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLoadRecord.ControlTipText
End Sub

Private Sub iLogAsk_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogAsk
    LoadListbox
End Sub

Private Sub iLogAsk_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLogAsk.ControlTipText
End Sub

Private Sub iLogClick_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogClick "left"
    LoadListbox
End Sub

Private Sub iLogClick_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLogClick.ControlTipText
End Sub

Private Sub iLogDouble_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogClick "double"
    LoadListbox
End Sub

Private Sub iLogDouble_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLogDouble.ControlTipText
End Sub

Private Sub iLogInput_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    LogText
    LoadListbox
End Sub

Private Sub iLogInput_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iLogInput.ControlTipText
End Sub

Private Sub iMemo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    EditMemo

End Sub

Private Sub iMemo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iMemo.ControlTipText
End Sub

Private Sub iNewFile_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    newRecord
    LoadListbox
End Sub

Private Sub iNewFile_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iNewFile.ControlTipText
End Sub

Private Sub iPlayAll_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
    MouseReplay
    Me.Show
End Sub

Private Sub iPlayAll_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iPlayAll.ControlTipText
End Sub

Private Sub iPlayFromHere_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
    PlayFromHere
    Me.Show
End Sub

Private Sub iPlayFromHere_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iPlayFromHere.ControlTipText
End Sub

Private Sub iPlaySelection_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    PlayBackSelectedRows
End Sub

Private Sub iPlaySelection_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iPlaySelection.ControlTipText
End Sub

Private Sub iPlayUntilHere_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Me.Hide
    PlayUntilHere
    Me.Show
End Sub

Private Sub iPlayUntilHere_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iPlayUntilHere.ControlTipText
End Sub

Private Sub iRecClick_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    RecordStart ClicksOnly.Value
    LoadListbox
End Sub

Private Sub iRecClick_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iRecClick.ControlTipText
End Sub

Private Sub iRecDrag_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    recordDrag
    LoadListbox
End Sub

Private Sub iRecDrag_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iRecDrag.ControlTipText
End Sub

Private Sub iSave_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    If ws.Range("A2") = "" Then
        infoLab.Caption = "Record something first"
        Exit Sub
    End If
    Dim FName As String
    FName = InputboxString(, , ws.Range("H1"))
    If Len(FName) <> 0 And FName <> "False" Then
        ws.Range("H1") = FName
        SaveRecord
    End If
    LoadMRcaption
End Sub

Private Sub iSave_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iSave.ControlTipText
End Sub

Private Sub iTop_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveToTop
    LoadListbox
End Sub

Private Sub iTop_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iTop.ControlTipText
End Sub

Private Sub iUp_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveRows -1
    LoadListbox
End Sub

Private Sub iUp_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = iUp.ControlTipText
End Sub

Private Sub info_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    infoLab.Caption = info.ControlTipText
End Sub

Private Sub lBoxData_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    EditRow
End Sub

Private Sub labHome_Click()
    labHome.SpecialEffect = fmSpecialEffectSunken
    labRec.SpecialEffect = fmSpecialEffectRaised
    fFile.left = 30
    fPlay.left = 30
    fRecord.left = 255
    fLog.left = 255
End Sub

Private Sub labRec_Click()
    labHome.SpecialEffect = fmSpecialEffectRaised
    labRec.SpecialEffect = fmSpecialEffectSunken
    fFile.left = 255
    fPlay.left = 255
    fRecord.left = 30
    fLog.left = 83
End Sub

Rem @NOT WORKING - forces motion top to bottom?
Sub moveFromAtoB(x0 As Long, y0 As Long, X1 As Long, Y1 As Long)
    Dim steep As Boolean: steep = Abs(Y1 - y0) > Abs(X1 - x0)
    Dim t As Integer
    If steep Then
        '// swap(x0, y0);
        t = x0
        x0 = y0
        y0 = t
        ' // swap(x1, y1);
        t = X1
        X1 = Y1
        Y1 = t
    End If
    If x0 > X1 Then
        '// swap(x0, x1);
        t = x0
        x0 = X1
        X1 = t
        '// swap(y0, y1);
        t = y0
        y0 = Y1
        Y1 = t
    End If
    Dim deltax As Integer: deltax = X1 - x0
    Dim deltay As Integer: deltay = Abs(Y1 - y0)
    Dim deviation As Integer: deviation = deltax / 2
    Dim ystep As Integer
    Dim Y  As Integer: Y = y0
    If y0 < Y1 Then
        ystep = 1
    Else
        ystep = -1
    End If
    Dim X As Integer
    For X = x0 To X1 Step ystep
        If steep Then
            SetCursorPos Y, X
        Else
            SetCursorPos X, Y
        End If
        deviation = deviation - deltay
        If deviation < 0 Then
            Y = Y + ystep
            deviation = deviation + deltax
        End If
        DoEvents
        Sleep 1
    Next
End Sub

Sub newRecord()
    '#INCLUDE LoadMRcaption
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    ws.Range("A2").CurrentRegion.OFFSET(1).ClearContents
    ws.Range("R7").CurrentRegion.OFFSET(1).ClearContents
    ws.Range("H1").ClearContents
    LoadMRcaption
End Sub

Sub recordDrag()
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("MouseDB")
    Dim rng As Range
    Dim lngCurPos As POINTAPI
    Dim previousX As Long, previousY As Long, activeX As Long, activeY As Long
    Dim previousL As Long, previousR As Long, activeL As Long, activeR As Long
    Erase MouseArray
    Dim arrayCounter As Long: arrayCounter = 1
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Do
        ReDim Preserve MouseArray(1 To arrayCounter)
        GetCursorPos lngCurPos
        activeL = IIf(GetAsyncKeyState(1) = 0, 0, 1)
        activeR = IIf(GetAsyncKeyState(2) = 0, 0, 1)
        activeX = lngCurPos.X
        activeY = lngCurPos.Y
        If previousX <> lngCurPos.X Or previousY <> lngCurPos.Y Or previousL <> activeL Or previousR <> activeR Then
            previousX = activeX
            previousY = activeY
            previousL = activeL
            previousR = activeR
            MouseArray(arrayCounter) = Join(Array(previousX, previousY, activeL, activeR), ",")
            arrayCounter = arrayCounter + 1
        End If
    Loop
LoopEnd:
    If err = 18 Then
        Application.EnableCancelKey = xlInterrupt
        Dim arr
        arr = MouseArray
        arr = Filter(arr, ",1,", , vbTextCompare)
        Set rng = ws.Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
        'Set rng = rng.Resize(UBound(arr), 1)
        rng = WorksheetFunction.Transpose(arr)
        rng.TextToColumns rng, comma:=True
        rng.OFFSET(0, 4) = "DRAG"
        rng.OFFSET(2).RESIZE(rng.rows.count - 3, 5).Delete xlUp
        rng.RESIZE(1, 5).Delete xlUp
        rng.OFFSET(0, 2).RESIZE(1, 2).Value = rng.OFFSET(1).RESIZE(1, 2).Value
        rng.OFFSET(1).RESIZE(1, 5).Delete xlUp
        infoLab.Caption = "Drag recorded."
        '        infoLab.Caption = "Drag recorded at rows: " & rng.Row & " to " & rng.Row + rng.Rows.Count
    End If
End Sub

' Enum MouseButtonConstants
' vbLeftButton
' vbMiddleButton
' vbRightButton
' End Enum
'
''simulate the MouseDown event
' Sub ButtonDown(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    Dim lFlag As Long
'    If Button = vbLeftButton Then
'        lFlag = MOUSEEVENTF_LEFTDOWN
'    ElseIf Button = vbMiddleButton Then
'        lFlag = MOUSEEVENTF_MIDDLEDOWN
'    ElseIf Button = vbRightButton Then
'        lFlag = MOUSEEVENTF_RIGHTDOWN
'    End If
'    mouse_event lFlag, 0, 0, 0, 0
'End Sub
'
''simulate the MouseUp event
'
' Sub ButtonUp(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    Dim lFlag As Long
'    If Button = vbLeftButton Then
'        lFlag = MOUSEEVENTF_LEFTUP
'    ElseIf Button = vbMiddleButton Then
'        lFlag = MOUSEEVENTF_MIDDLEUP
'    ElseIf Button = vbRightButton Then
'        lFlag = MOUSEEVENTF_RIGHTUP
'    End If
'    mouse_event lFlag, 0, 0, 0, 0
'End Sub
'
''simulate the MouseClick event
'
' Sub ButtonClick(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    ButtonDown Button
'    ButtonUp Button
'End Sub
'
''simulate the MouseDblClick event
'
' Sub ButtonDblClick(Optional ByVal Button As MouseButtonConstants = _
'    vbLeftButton)
'    ButtonClick Button
'    ButtonClick Button
'End Sub


'Sub AlternativeLogPlayback()
'Rem from different logging style
'Dim DefaultSleep As Long
'DefaultSleep = 1000
'Dim cell As Range, rng As Range
'Set rng = ActiveSheet.Range("A1").CurrentRegion
'Set rng = rng.Resize(, 1).offset(1).Resize(rng.rows.count - 1)
'    Dim lngCurPos As POINTAPI, activeX As Long, activeY As Long
'    GetCursorPos lngCurPos
'    activeX = lngCurPos.x
'    activeY = lngCurPos.y
'For Each cell In rng
'    If cell <> "drag" Then
'        If cell.offset(0, 1) <> "" And cell.offset(0, 2) <> "" Then
'            'moveFromAtoB activeX, activeY, CLng(cell.offset(0, 1)), CLng(cell.offset(0, 2).Value)
'            SetCursorPos cell.offset(0, 1), cell.offset(0, 2)
'        End If
'   End If
'    If cell = "move" Then
'        'moveFromAtoB activeX, activeY, CLng(cell.offset(0, 1)), CLng(cell.offset(0, 2).Value)
'       SetCursorPos cell.offset(0, 1), cell.offset(0, 2)
'    ElseIf cell = "left" Then LeftClick
'    ElseIf cell = "double" Then DoubleClick
'    ElseIf cell = "right" Then RightClick
'    ElseIf cell = "drag" Then
'        GetCursorPos lngCurPos
'        activeX = lngCurPos.x
'        activeY = lngCurPos.y
'        dragMouse activeX, activeY, cell.offset(0, 1), cell.offset(0, 2)
'    ElseIf cell = "type" Then
'        CLIP cell.offset(0, 1)
'        SendKeys CLIP, True
'    ElseIf cell = "ask" Then
'        Dim msg As String
'        msg = InputboxString()
'        CLIP msg
'        SendKeys CLIP, True
'    End If
'    If cell = "wait" Then
'        Sleep IIf(cell.offset(0, 1) <> "", cell.offset(0, 1), DefaultSleep)
'    Else
'        Sleep DefaultSleep
'    End If
'    DoEvents
'Next
'End Sub


Sub CreateListboxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
    '#INCLUDE ArrayDimensions

    header.Width = body.Width
    Dim i As Long
    
    'must have a listbox to use as headers
    
    'header.columnCount = body.columnCount
    If header.ColumnWidths <> body.ColumnWidths And body.ColumnWidths <> "" Then
        header.ColumnWidths = body.ColumnWidths
    End If
    'add headerelements
    header.clear
    header.AddItem
    
    If ArrayDimensions(arrHeaders) = 1 Then
        For i = 0 To UBound(arrHeaders)
            'make it prety
            header.list(0, i) = arrHeaders(i)
        Next i
    Else
        For i = 1 To UBound(arrHeaders, 2)
            header.list(0, i - 1) = arrHeaders(1, i)
        Next i
    End If
    body.ZOrder (1)
    header.ZOrder (0)
    header.SpecialEffect = fmSpecialEffectFlat
    header.BackColor = RGB(200, 200, 200)
    'align header to body
    header.Height = 15
    header.Width = body.Width
    header.left = body.left
    header.top = body.top - header.Height - 1
    header.Font.Bold = True
    header.Font.Name = "Comic Sans MS"
    header.Font.Size = 9
    '    header.ForeColor =
    '    header.BackColor =
End Sub


