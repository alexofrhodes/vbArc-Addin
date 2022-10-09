Attribute VB_Name = "M_BarMan"

Rem @Folder Barman
Public Const rBUILD_ON_OPEN = "I2"
Public Const rC_TAG = "I4"
Public Const rMENU_TYPE = "I5"
Public Const rBAR_LOCATION = "I6"
Public Const rTARGET_CONTROL = "I7"
Public C_TAG As String
Public MenuEvent As CVBECommandHandler
Public EventHandlers As New Collection
Public CmdBarItem As CommandBarControl
Public TargetCommandbar
Public TargetControl As CommandBarControl
Public MainMenu As CommandBarControl
Public MenuItem As CommandBarControl
Public ctrl As Office.CommandBarControl
Public MenuLevel, NextLevel, Caption, Divider, FaceId
Public action As String
Public MenuSheet As Worksheet
Public row As Integer
Public MenuType As Long
Public Const WorksheetMenu = 1
Public Const VbeMenu = 2
Public Const RightClickMenu = 3
Public BarLocation As String

Public Sub CreateAllBars()
    '#INCLUDE CommandBarBuilder
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(left(ws.Name, 4)) = "BAR_" Then
            If ws.Range(rBUILD_ON_OPEN) = True Then CommandBarBuilder ws
        End If
    Next
End Sub

Public Sub DeleteAllBars()
    '#INCLUDE DeleteControlsAndHandlers
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(left(ws.Name, 4)) = "BAR_" Then DeleteControlsAndHandlers ws
    Next
End Sub

Public Sub RestoreBars()
    '#INCLUDE CreateAllBars
    Application.OnTime Now, "CreateAllBars"
End Sub

Public Sub ListBars()
    '#INCLUDE ListWorksheetBars
    '#INCLUDE ListVBEBars
    ListWorksheetBars
    ListVBEBars
End Sub

Public Sub NewBar()
    '#INCLUDE lastBar
    Dim wsMain As Worksheet
    Set wsMain = ThisWorkbook.Worksheets("BAR_Main")
    Dim wsCopy As Worksheet
    wsMain.Copy After:=ThisWorkbook.SHEETS(ThisWorkbook.SHEETS.count)
    Set wsCopy = ThisWorkbook.SHEETS(ThisWorkbook.SHEETS.count)
    wsCopy.Name = "BAR_" & lastBar + 1
    wsCopy.Range("A1").CurrentRegion.OFFSET(1).ClearContents
    wsCopy.Range("I4:I7").ClearContents
    wsCopy.Range("I2") = False
End Sub

Private Function lastBar() As Long
    Dim counter As Long
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If UCase(left(ws.Name, 4)) = "BAR_" Then counter = counter + 1
    Next
    lastBar = counter
End Function

Private Function SetCMDbar(ws As Worksheet) As Boolean
    C_TAG = ws.Range(rC_TAG)
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case Is = LCase("WorksheetMenu")
            MenuType = WorksheetMenu
        Case Is = LCase("vbeMenu")
            MenuType = VbeMenu
        Case Is = LCase("RightClickMenu")
            MenuType = RightClickMenu
        Case Else
    End Select
    If ws.Range(rBAR_LOCATION) <> "" Then
        BarLocation = ws.Range(rBAR_LOCATION)
    Else
        BarLocation = 0
    End If
    If MenuType = VbeMenu Then
        Select Case BarLocation
            Case Is = "Menu Bar", "Code Window", "Project Window", "Edit", "Debug", "Userform"
                Set TargetCommandbar = Application.VBE.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
                Set TargetCommandbar = Application.VBE.CommandBars.Add(C_TAG, Position:=msoBarTop, Temporary:=True)
                TargetCommandbar.visible = True
        End Select
    ElseIf MenuType = WorksheetMenu Then
        Select Case ws.Range(rBAR_LOCATION)
            Case Is = "Worksheet Menu Bar", "Cell", "Column", "Row"
                Set TargetCommandbar = Application.CommandBars(BarLocation)
                SetCMDbar = True
            Case Else
        End Select
    Else
    End If
End Function

Public Function BarExists(findBarName As String) As Boolean
    Dim bar As CommandBar
    For Each bar In Application.CommandBars
        If UCase(bar.Name) = UCase(findBarName) Then
            BarExists = True
            Exit Function
        End If
    Next bar
    For Each bar In Application.VBE.CommandBars
        If UCase(bar.Name) = UCase(findBarName) Then
            BarExists = True
            Exit Function
        End If
    Next bar
End Function

Public Sub BuildBarFromShape()
    '#INCLUDE CommandBarBuilder
    CommandBarBuilder ActiveSheet
End Sub

Public Sub DeleteBarFromShape()
    '#INCLUDE DeleteControlsAndHandlers
    DeleteControlsAndHandlers ActiveSheet
End Sub

Public Sub CommandBarBuilder(ws As Worksheet)
    '#INCLUDE SetCMDbar
    '#INCLUDE SetVariables
    '#INCLUDE ReSetVariables
    '#INCLUDE CreateMainMenu
    '#INCLUDE CreatePopup
    '#INCLUDE CreateButton
    '#INCLUDE DirectButton
    '#INCLUDE DeleteControlsAndHandlers
    If ws.Range("I4") = "" Or ws.Range("I5") = "" Or ws.Range("I6") = "" Then
        MsgBox "Ranges I4, I5 and I6 cannot be empty"
        Exit Sub
    End If
    DeleteControlsAndHandlers ws
    SetCMDbar ws
    Set MenuSheet = ws
    row = 2
    If MenuType = VbeMenu Then
        If BarLocation = "Menu Bar" Then
            Select Case LCase(ws.Range(rTARGET_CONTROL))
                Case LCase(ws.Range(rC_TAG))
                    Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
                Case Else
                    Dim vbControl As String
                    vbControl = ws.Range(rTARGET_CONTROL)
                    Set TargetControl = TargetCommandbar.Controls(vbControl).Controls.Add(Type:=msoControlPopup, Temporary:=True)
            End Select
        Else
            If LCase(ws.Range(rTARGET_CONTROL)) <> LCase(ws.Range(rC_TAG)) _
        And ws.Range(rTARGET_CONTROL) <> "" Then _
                                      Set TargetControl = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
        End If
        If Not TargetControl Is Nothing Then
            TargetControl.Caption = C_TAG
            TargetControl.Tag = C_TAG
        End If
    End If
    Do Until IsEmpty(MenuSheet.Cells(row, 1))
        With MenuSheet
            SetVariables
        End With
        Select Case MenuLevel
            Case 1
                If NextLevel > MenuLevel Then
                    CreateMainMenu
                Else
                    DirectButton
                End If
            Case 2
                If NextLevel > MenuLevel Then
                    CreatePopup
                Else
                    DirectButton
                End If
            Case 3
                CreateButton
        End Select
        row = row + 1
        ReSetVariables
    Loop
End Sub

Private Sub markControlType(ws As Worksheet)
    ws.Columns("F").ClearContents
    Dim idx As Long: idx = 0
    Dim Description() As Variant
    Dim cell As Range
    Set cell = ws.Cells(2, 1)
    Do Until IsEmpty(cell)
        idx = idx + 1
        ReDim Preserve Description(1 To idx)
        Description(idx) = IIf(cell.OFFSET(1) > cell, "PopUp", "Button")
        Set cell = cell.OFFSET(1)
    Loop
    ws.Range("F2").RESIZE(UBound(Description)) = WorksheetFunction.Transpose(Description)
End Sub

Private Sub SetVariables()
    With MenuSheet
        MenuLevel = .Cells(row, 1)
        Caption = .Cells(row, 2)
        action = .Cells(row, 3)
        Divider = .Cells(row, 4)
        FaceId = .Cells(row, 5)
        NextLevel = .Cells(row + 1, 1)
    End With
End Sub

Private Sub ReSetVariables()
    MenuLevel = ""
    Caption = ""
    action = ""
    Divider = ""
    FaceId = ""
    NextLevel = ""
End Sub

Private Sub CreateMainMenu()
    If MenuType = VbeMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup)
    ElseIf MenuType = WorksheetMenu Then
        Set MainMenu = TargetCommandbar.Controls.Add(Type:=msoControlPopup, Temporary:=True)
    ElseIf MenuType = RightClickMenu Then
        On Error Resume Next
        CommandBars.Add C_TAG, msoBarPopup, , True
        On Error GoTo 0
        Set MainMenu = CommandBars(C_TAG).Controls.Add(Type:=msoControlPopup)
    End If
    With MainMenu
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreatePopup()
    If MenuType = RightClickMenu Then
        Set MenuItem = MainMenu.Controls.Add(Type:=msoControlPopup)
    Else
        Set MenuItem = MainMenu.Controls.Add(Type:=msoControlPopup)
    End If
    With MenuItem
        .Caption = Caption
        .BeginGroup = Divider
        If FaceId <> "" And action <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
End Sub

Private Sub CreateButton()
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Set CmdBarItem = MenuItem.Controls.Add
    With CmdBarItem
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub

Private Sub DirectButton()
    Dim CmdBarItem As CommandBarControl
    If MenuType = VbeMenu Then
        Set MenuEvent = New CVBECommandHandler
    End If
    Select Case MenuLevel
        Case Is = 1
            Set CmdBarItem = TargetCommandbar.Controls.Add
        Case Is = 2
            Set CmdBarItem = MainMenu.Controls.Add
    End Select
    With CmdBarItem
        .Style = msoButtonIconAndCaption
        .Caption = Caption
        .BeginGroup = Divider
        .OnAction = action
        If FaceId <> "" Then .FaceId = FaceId
        .Tag = C_TAG
    End With
    If MenuType = VbeMenu Then
        Set MenuEvent.EvtHandler = Application.VBE.Events.CommandBarEvents(CmdBarItem)
        EventHandlers.Add MenuEvent
    End If
End Sub

Private Sub DeleteControlsAndHandlers(ws As Worksheet)
    '#INCLUDE BarExists
    '#INCLUDE DeleteHandlersFor
    If ws.Range(rC_TAG).TEXT = vbNullString Then Exit Sub
    Select Case LCase(ws.Range(rMENU_TYPE))
        Case "vbemenu"
            MenuType = VbeMenu
        Case "worksheetmenu"
            MenuType = WorksheetMenu
        Case "rightclickmenu"
            MenuType = RightClickMenu
    End Select
    If MenuType = VbeMenu Then
        If BarExists(ws.Range(rC_TAG)) Then
            Application.VBE.CommandBars(ws.Range(rC_TAG).TEXT).Delete
            Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).TEXT)
        End If
        Rem
    ElseIf MenuType = WorksheetMenu Then
        Set ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).TEXT)
    ElseIf MenuType = RightClickMenu Then
        If BarExists(ws.Range(rC_TAG).TEXT) Then
            CommandBars(ws.Range(rC_TAG).TEXT).Delete
        End If
        Exit Sub
    End If
    On Error Resume Next
    Do
        ctrl.Delete
        If MenuType = VbeMenu Then
            Rem
            Set ctrl = Application.VBE.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).TEXT)
        Else
            Set ctrl = Application.CommandBars.FindControl(Tag:=ws.Range(rC_TAG).TEXT)
        End If
    Loop While Not ctrl Is Nothing
    On Error GoTo 0
    DeleteHandlersFor ws
End Sub

Private Sub DeleteHandlersFor(ws As Worksheet)
    On Error Resume Next
    Dim cell As Range
    Set cell = ws.Cells(2, 6)
    Do Until IsEmpty(cell)
        If cell.TEXT = "Button" Then
            EventHandlers.Remove cell.OFFSET(0, -3).TEXT
        End If
        Set cell = cell.OFFSET(1)
    Loop
End Sub

Private Sub ListWorksheetBars()
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.SHEETS("ListSheetBars")
    oWK.Cells.clear
    Dim arr As Variant
    arr = Array("Type", "Index", "Name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").RESIZE(1, iCol) = arr
    oWK.Range("a1").RESIZE(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 4) As Variant
    For Each oCB In Excel.Application.CommandBars
        cbVar(i - 2, 0) = Choose(oCB.Type + 1, "Toolbar", "Menu", "PopUp")
        cbVar(i - 2, 1) = oCB.index
        cbVar(i - 2, 2) = oCB.Name
        cbVar(i - 2, 3) = oCB.BuiltIn
        cbVar(i - 2, 4) = oCB.visible
        i = i + 1
    Next
    oWK.Cells(2, 1).RESIZE(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub ListVBEBars()
    Dim oCB As CommandBar
    Dim oWK As Worksheet
    Set oWK = ThisWorkbook.SHEETS("ListVBEBars")
    oWK.Cells.clear
    Dim arr As Variant
    arr = Array("Type", "Index", "Name", "Built-in", "Visible")
    Dim iCol As Integer
    iCol = UBound(arr) + 1
    oWK.Range("a1").RESIZE(1, iCol) = arr
    oWK.Range("a1").RESIZE(1, iCol).Cells.Font.Bold = True
    Dim i As Long
    i = 2
    Dim cbVar(300, 4) As Variant
    For Each oCB In Application.VBE.CommandBars
        cbVar(i - 2, 0) = Choose(oCB.Type + 1, "Toolbar", "Menu", "PopUp")
        cbVar(i - 2, 1) = oCB.index
        cbVar(i - 2, 2) = oCB.Name
        cbVar(i - 2, 3) = oCB.BuiltIn
        cbVar(i - 2, 4) = oCB.visible
        i = i + 1
    Next
    oWK.Cells(2, 1).RESIZE(UBound(cbVar, 1) + 1, UBound(cbVar, 2) + 1) = cbVar
    oWK.Columns.AutoFit
End Sub

Private Sub exampleOfControls()
    Dim cbc As CommandBarControl
    Dim cbb As CommandBarButton
    Dim cbcm As CommandBarComboBox
    Dim cbp As CommandBarPopup
    With Application.VBE.CommandBars("CodeArchive")
        Set cbc = .Controls.Add(ID:=3, Temporary:=True)
        Set cbb = .Controls.Add(Temporary:=True)
        cbb.Caption = "A new command"
        cbb.Style = msoButtonCaption
        cbb.OnAction = "NewCommand_OnAction"
        Set cbcm = .Controls.Add(Type:=msoControlComboBox, Temporary:=True)
        cbcm.Caption = "Combo:"
        cbcm.AddItem "list entry 1"
        cbcm.AddItem "list entry 2"
        cbcm.OnAction = "NewCommand_OnAction"
        cbcm.Style = msoComboLabel
        Set cbc = .Controls.Add(Type:=msoControlDropdown, Temporary:=True)
        cbc.Caption = "Dropdown:"
        cbc.AddItem "list entry 1"
        cbc.AddItem "list entry 2"
        cbc.OnAction = "MenuDropdown_OnAction"
        Set cbp = .Controls.Add(Type:=msoControlPopup, Temporary:=True)
        cbp.Caption = "new sub menu"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 1"
        Set cbb = cbp.Controls.Add
        cbb.Caption = "sub entry 2"
    End With
End Sub

Private Sub ImageFromEmbedded()
    Dim p As Excel.Picture
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    Set p = Worksheets("Sheet1").Pictures("ThePict")
    p.CopyPicture xlScreen, xlBitmap
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .PasteFace
    End With
End Sub

Private Sub ImageFromExternalFile()
    Dim Btn As Office.CommandBarButton
    Set Btn = Application.CommandBars.FindControl(ID:=30007) _
        .Controls.Add(Type:=msoControlButton, Temporary:=True)
    With Btn
        .Caption = "Click Me"
        .Style = msoButtonIconAndCaption
        .Picture = LoadPicture("C:\TestPic.bmp")
    End With
End Sub

Private Sub ResetCBAR()
    Excel.Application.CommandBars("Cell").Reset
End Sub

Public Function IsLoaded(FormName As String) As Boolean
    Dim frm As Object
    For Each frm In VBA.UserForms
        If frm.Name = FormName Then
            IsLoaded = True
            Exit Function
        End If
    Next frm
    IsLoaded = False
End Function

Sub openUValiationDropdown()
    lngValType = ActiveCell.Validation.Type
    Select Case lngValType
        Case Is = 3
            uValidationDropdown.Show
        Case Else
            Unload uValidationDropdown
    End Select
End Sub

'''''NOTES'''''''
'''''''''''''''''
'----------------------
'WORKSHEET COMMAND BARS
'----------------------
'Application.CommandBars("Worksheet Menu Bar").Controls.Add
'Application.CommandBars("Cell").Controls.Add
'Application.CommandBars("Column").Controls.Add
'Application.CommandBars("Row").Controls.Add
'----------------------
''VBE COMMAND BARS
'----------------------
'----------------------
''add your own command bar
'----------------------
'With Application.VBE.CommandBars.Add("CodeArchive", Position:=msoBarFloating, Temporary:=True)
'    .Visible = True
'End With
'Application.VBE.CommandBars("CodeArchive").Delete
'----------------------
''use existing command bars
'----------------------
'Set TargetControl = Application.VBE.CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Code Window").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Project Window").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Edit").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Debug").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'Set TargetControl = Application.VBE.CommandBars("Userform").Controls.Add(Type:=msoControlPopup, Temporary:=True)
'----------------------
''use existing controls
'----------------------
'Set TargetControl = Application.VBE.CommandBars("Menu Bar").Controls.("Tools")
'-----------
'Use combobox
'-----------
''call a sub through class events handler
''the sub to contain the following
'With Application.VBE.ActiveCodePane
'  Text = Application.VBE.CommandBars(mcToolBar).Controls(mcInsertList).Text
'  .GetSelection StartLine, StartColumn, EndLine, EndColumn
'  .CodeModule.InsertLines StartLine, Text
'  .SetSelection StartLine, 1, StartLine, 1
'End With


