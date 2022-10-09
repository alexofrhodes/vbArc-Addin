VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetVisibility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Const VISIBILITY_MANAGER As String = "o"

Sub AddGoToSheetVisibilityManagerShape()
    '#INCLUDE GoToSheetVisibilityManager
    If TypeName(Selection) <> "Range" Then Exit Sub
    Application.ScreenUpdating = False
    Dim s As Shape
    Dim ws As Worksheet
    Set s = ThisWorkbook.SHEETS(VISIBILITY_MANAGER).Shapes("GoToSheetVisibilityManager")
    s.Copy
    Dim ts As Shape
    Dim cell As Range
    For Each cell In Selection
        If cell <> "" Then
            Set ws = Workbooks(s.parent.parent.Name).SHEETS(cell.TEXT)
            On Error Resume Next
            Set ts = ws.Shapes("GoToSheetVisibilityManager")
            On Error GoTo 0
            If ts Is Nothing Then
                ws.Paste
                Set ts = ws.Shapes(ws.Shapes.count)
                ts.left = 1
                ts.top = 1
                ws.Hyperlinks.Add Anchor:=ts, Address:="", _
                                  SubAddress:="o!A1", ScreenTip:=""
                Application.GoTo ws.Cells(1, 1), True
            End If
        End If
    Next
    GoToSheetVisibilityManager
    Application.ScreenUpdating = True
End Sub

Sub GoToSheetVisibilityManager()
    ThisWorkbook.SHEETS(VISIBILITY_MANAGER).Activate
End Sub

Sub AddVisibilityControls()
    Dim FromCell As Range
    Dim ToCell As Range
    Dim s As Shape
    With ThisWorkbook.SHEETS(VISIBILITY_MANAGER)
        Set FromCell = .Range("A2")
        Set ToCell = .Cells(2, 3)
        Do While CellHasShape(ToCell) = True
            Set ToCell = ToCell.OFFSET(0, 1)
        Loop
        .Columns(ToCell.Column).ColumnWidth = 10
        FromCell.Copy ToCell
        For Each s In .Shapes
            If Not Intersect(ToCell, s.TopLeftCell) Is Nothing Then
                s.visible = True
                s.Name = "Shape_" & Replace(s.TopLeftCell.Address, "$", "")
            End If
        Next
        .Columns(ToCell.Column).AutoFit
    End With
End Sub

Function CellHasShape(cell As Range)
    Dim s As Shape
    With ThisWorkbook.SHEETS(VISIBILITY_MANAGER)
        For Each s In .Shapes
            If s.TopLeftCell.Address = cell.Address Then
                CellHasShape = True
                Exit Function
            End If
        Next
    End With
End Function

Sub DeleteShapesFromSelectedCells()
    Dim s As Shape
    For Each s In ThisWorkbook.SHEETS(VISIBILITY_MANAGER).Shapes
        If s.TopLeftCell.Column > 2 Then
            If Not Intersect(s.TopLeftCell, Selection) Is Nothing Then s.Delete
        End If
    Next
End Sub

Sub ListSheets()
    '#INCLUDE LoopSheetsSignalVisibility
    Application.ScreenUpdating = False
    Dim cell As Range
    Set cell = ThisWorkbook.SHEETS(VISIBILITY_MANAGER).Range("A3")
    Range(cell, cell.End(xlDown)).ClearContents
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name <> VISIBILITY_MANAGER Then
            cell = ws.Name
            Set cell = cell.OFFSET(1, 0)
        End If
    Next
    Range(cell, cell.End(xlDown)).HorizontalAlignment = xlLeft
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub ActivateSheet()
    SHEETS(ActiveSheet.Shapes(Application.Caller).TextFrame2.TextRange.TEXT).Activate
End Sub

Sub AllSheetsVisible()
    '#INCLUDE LoopSheetsSignalVisibility
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        ws.visible = xlSheetVisible
    Next
    SheetVisibility.Activate
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub AllSheetsHidden()
    '#INCLUDE LoopSheetsSignalVisibility
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    For Each ws In Worksheets
        If ws.Name <> VISIBILITY_MANAGER Then ws.visible = xlSheetHidden
    Next
    SheetVisibility.Activate
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub LoopSheetsSignalVisibility()
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ThisWorkbook.SHEETS(VISIBILITY_MANAGER).UsedRange.SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    Dim cell As Range
    For Each cell In rng
        On Error Resume Next
        Set ws = SHEETS(cell.TEXT)
        On Error GoTo 0
        If Not ws Is Nothing Then
            If ws.visible = xlSheetVisible Then
                cell.Font.color = RGB(0, 128, 0)
            Else
                cell.Font.color = RGB(128, 0, 0)
            End If
        End If
        Set ws = Nothing
    Next
    ThisWorkbook.SHEETS(VISIBILITY_MANAGER).Activate
    Application.ScreenUpdating = True
End Sub

Sub LoopSheetsMakeVisible()
    '#INCLUDE LoopSheetsSignalVisibility
    Application.ScreenUpdating = False
    Dim actSheet As Worksheet: Set actSheet = ActiveSheet
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ActiveSheet.Columns(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    Dim cell As Range
    For Each cell In rng
        On Error Resume Next
        Set ws = SHEETS(cell.TEXT)
        On Error GoTo 0
        If Not ws Is Nothing Then ws.visible = xlSheetVisible
    Next
    actSheet.Activate
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub LoopSheetsMakeHidden()
    '#INCLUDE LoopSheetsSignalVisibility
    Application.ScreenUpdating = False
    Dim actSheet As Worksheet: Set actSheet = ActiveSheet
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ActiveSheet.Columns(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    Dim cell As Range
    For Each cell In rng
        On Error Resume Next
        Set ws = SHEETS(cell.TEXT)
        On Error GoTo 0
        If Not ws Is Nothing Then ws.visible = xlSheetHidden
    Next
    actSheet.Activate
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub LoopSheetsToggleVisibility()
    '#INCLUDE LoopSheetsSignalVisibility
    '#INCLUDE ToggleSheetView
    Application.ScreenUpdating = False
    Dim actSheet As Worksheet: Set actSheet = ActiveSheet
    Dim ws As Worksheet
    Dim rng As Range
    On Error Resume Next
    Set rng = ActiveSheet.Columns(ActiveSheet.Shapes(Application.Caller).TopLeftCell.Column).SpecialCells(xlCellTypeConstants)
    On Error GoTo 0
    If rng Is Nothing Then Exit Sub
    Dim cell As Range
    For Each cell In rng
        On Error Resume Next
        Set ws = SHEETS(cell.TEXT)
        On Error GoTo 0
        If Not ws Is Nothing Then ToggleSheetView ws
    Next
    actSheet.Activate
    LoopSheetsSignalVisibility
    Application.ScreenUpdating = True
End Sub

Sub ToggleSheetView(ws As Worksheet)
    If ws.visible = xlSheetHidden Then
        ws.visible = xlSheetVisible
    Else
        If ws.Name <> ">" Then ws.visible = xlSheetHidden
    End If
End Sub

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim rng As Range
    Set rng = ActiveSheet.Cells.SpecialCells(xlCellTypeConstants)
    If Intersect(Target, rng) Is Nothing Then Exit Sub
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS(Target.TEXT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    
    ToggleSheetView ThisWorkbook.SHEETS(Target.TEXT)
   
    LoopSheetsSignalVisibility
    Cancel = True
End Sub

Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.SHEETS(CStr(Target))
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.visible = xlSheetVisible
    LoopSheetsSignalVisibility
    ws.Activate
    Cancel = True
End Sub


