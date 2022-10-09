Attribute VB_Name = "M_DataEntry"
Rem @Folder DataEntry
Sub ShakeTableWand()
    If SetControlRange("TableControl" & "_" & ActiveSheet.Name, _
                       "Range selection", _
                       "Select table controller RANGE" & vbNewLine & vbNewLine & _
                       "Include cells with text (label) and empty cells to the right (input field)") = False Then Exit Sub
    '#INCLUDE CopyModule
    '#INCLUDE addVBEreference
    '#INCLUDE SetControlRange
    '#INCLUDE setControlInputFields
    '#INCLUDE SetTableHeadersRange
    '#INCLUDE CreateTable
    '#INCLUDE setLoadedRow
    '#INCLUDE SetMappingRange
    '#INCLUDE TableWorksheetCode
    '#INCLUDE CreateControlShapes
    Application.ScreenUpdating = False
    Call setControlInputFields
    Call SetTableHeadersRange("TableHeaderRange" & "_" & ActiveSheet.Name, 2)
    Call CreateTable(ActiveSheet.Range("TableHeaderRange" & "_" & ActiveSheet.Name).Worksheet, _
                     ActiveSheet.Range("TableHeaderRange" & "_" & ActiveSheet.Name), _
                     "DynamicTable" & "_" & ActiveSheet.Name)
    Call setLoadedRow(0)
    Call SetMappingRange("MappingRange" & "_" & ActiveSheet.Name)
    Call CreateControlShapes
    CopyModule "mTableWand", ThisWorkbook, ActiveWorkbook, False
    addVBEreference
    Call TableWorksheetCode
    ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Cells(1, 2).Select
    Application.ScreenUpdating = True
End Sub

Sub addVBEreference()
    On Error Resume Next
    ActiveWorkbook.VBProject.REFERENCES.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=5, Minor:=3
    On Error GoTo 0
End Sub

Function SetControlRange(RangeName As String, _
                         Optional sTitle As String, _
                         Optional sPrompt As String) As Boolean
    '#INCLUDE getRangeFromSelection
    '#INCLUDE DuplicateInRange
    '#INCLUDE createNamedRange
    SetControlRange = True
    Dim SelectedRange As Range
    Set SelectedRange = getRangeFromSelection(sTitle, sPrompt)
    If SelectedRange Is Nothing Then
        MsgBox "The range was cancelled"
        SetControlRange = False
        Exit Function
    Else
        If DuplicateInRange(SelectedRange) = True Then
            MsgBox ("Duplicates found in chosen range.")
            SetControlRange = False
            Exit Function
        ElseIf Application.WorksheetFunction.CountA(SelectedRange) = 0 Then
            MsgBox ("No Control labels found")
            SetControlRange = False
            Exit Function
        Else
            If SelectedRange.Column = 1 Then
                SelectedRange.Cells(1, 1).EntireColumn.Insert
            End If
            Call createNamedRange(RangeName, _
                                  SelectedRange.Worksheet, _
                                  SelectedRange)
        End If
    End If
    With SelectedRange.SpecialCells(xlCellTypeConstants).BORDERS
        .LineStyle = xlContinuous
        .color = vbBlack
        .Weight = xlThin
    End With
    With SelectedRange.SpecialCells(xlCellTypeConstants)
        .Interior.ColorIndex = 23
        .Characters.Font.color = vbWhite
        .Characters.Font.Bold = True
    End With
End Function

Function getRangeFromSelection(Optional sTitle As String, _
                               Optional sPrompt As String) As Range
    On Error Resume Next
    Set getRangeFromSelection = _
                              Application.InputBox(title:=sTitle, _
                                                   Prompt:=sPrompt, _
                                                   Type:=8, _
                                                   Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""))
    On Error GoTo 0
End Function

Function DuplicateInRange(r As Range) As Boolean
    Dim cell As Range
    Dim scr As Object
    Set scr = CreateObject("scripting.dictionary")
    With scr
        For Each cell In Selection.SpecialCells(xlCellTypeConstants)
            If cell.TEXT <> "" Then
                Debug.Print cell.Address & vbTab & cell.TEXT
                If Not .Exists(cell.Value) Then
                    .Add cell.Value, Nothing
                Else
                    DuplicateInRange = True
                    Set scr = Nothing
                    Exit Function
                End If
            End If
        Next cell
    End With
    Set scr = Nothing
End Function

Sub createNamedRange(RangeName As String, _
                     Worksheet As Worksheet, _
                     NamedRange As Variant)
    ActiveWorkbook.Names.Add Name:=RangeName, RefersTo:=NamedRange
End Sub

Sub setControlInputFields()
    '#INCLUDE createNamedRange
    Dim SelectedRange As Range
    Dim cell As Range
    For Each cell In ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).SpecialCells(xlCellTypeConstants)
        If SelectedRange Is Nothing Then
            Set SelectedRange = cell.MergeArea.OFFSET(0, 1).MergeArea
        Else
            Set SelectedRange = Union(SelectedRange, cell.MergeArea.OFFSET(0, 1).MergeArea)
        End If
    Next
    Call createNamedRange("ControlInputFields" & "_" & ActiveSheet.Name, _
                          SelectedRange.parent, _
                          SelectedRange)
    For Each cell In SelectedRange
        cell.MergeArea.BorderAround xlContinuous, xlThin
    Next
End Sub

Sub SetTableHeadersRange(RangeName As String, _
                         Optional offsetRows As Long = 3)
    '#INCLUDE createNamedRange
    '#INCLUDE getTableLastRow
    If offsetRows < 3 Then offsetRows = 3
    Dim cell As Range
    Dim coll As New Collection
    Dim txt As String
    On Error Resume Next
    For Each cell In ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).SpecialCells(xlCellTypeConstants)
        txt = WorksheetFunction.Concat(cell.MergeArea)
        coll.Add txt, txt
    Next cell
    On Error GoTo 0
    Dim SelectedRange As Range
    Set SelectedRange = ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Worksheet.Cells( _
        getTableLastRow(ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name)), _
        ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Column) _
        .OFFSET(offsetRows, 0). _
        RESIZE(1, coll.count)
    Call createNamedRange(RangeName, _
                          SelectedRange.Worksheet, _
                          SelectedRange)
    Dim i As Long
    For i = 1 To coll.count
        SelectedRange.Cells(1, i).Value = coll(i)
    Next i
End Sub

Function getTableLastRow(TableRange As Range) As Long
    getTableLastRow = TableRange.row + TableRange.rows.count - 1
End Function

Sub CreateTable(myWorksheet As Worksheet, _
                myRange As Range, _
                TableName As String)
    Dim MyTable As ListObject
    myWorksheet.ListObjects.Add _
        (xlSrcRange, myRange, , xlYes) _
        .Name = TableName
End Sub

Function setLoadedRow(LoadedTableRow As Long)
    Call createNamedRange("LoadedRow" & "_" & ActiveSheet.Name, _
                          ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent, _
                          LoadedTableRow)
    '#INCLUDE createNamedRange
End Function

Sub SetMappingRange(RangeName As String)
    '#INCLUDE createNamedRange
    '#INCLUDE RangeOfValueTableWand
    Dim cell As Range
    For Each cell In ActiveSheet.Range("TableHeaderRange" & "_" & ActiveSheet.Name)
        cell.OFFSET(-1, 0).Value = RangeOfValueTableWand(cell.Value, ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name), 0, 1).Address
    Next cell
    Call createNamedRange(RangeName, _
                          ActiveSheet.Range("TableHeaderRange" & "_" & ActiveSheet.Name).parent, _
                          ActiveSheet.Range("TableHeaderRange" & "_" & ActiveSheet.Name).OFFSET(-1, 0))
End Sub

Function RangeOfValueTableWand(findWhat As String, _
                               findWhere As Range, _
                               Optional offsetRow As Long = 0, _
                               Optional offsetCol As Long = 0) As Range
    For Each cell In findWhere
        If cell.Value = findWhat Then
            Set RangeOfValueTableWand = cell.MergeArea.OFFSET(offsetRow, offsetCol)
            Exit Function
        End If
    Next
End Function

Sub TableWorksheetCode()
    '#INCLUDE setLoadedRow
    '#INCLUDE PopulateTableControl
    '#INCLUDE WorksheetOfDynamicTable
    Dim CodeText As String
    CodeText = CodeText & "Private Sub Worksheet_SelectionChange(ByVal Target As Range)" & vbNewLine
    CodeText = CodeText & "    If Selection.ListObject Is Nothing Then Exit Sub" & vbNewLine
    CodeText = CodeText & "    If Selection.Cells.Count <> 1 Then Exit Sub" & vbNewLine
    CodeText = CodeText & "    If Selection.Row - Selection.ListObject.Range.Row = ActiveWorkbook.Names(""LoadedRow"" & ""_"" & ActiveSheet.Name) Then Exit Sub" & vbNewLine
    CodeText = CodeText & "    If Not Intersect(Target, ActiveSheet.Range(""DynamicTable"" & ""_"" & ActiveSheet.Name)) Is Nothing Then" & vbNewLine
    CodeText = CodeText & "        PopulateTableControl ActiveSheet.Range(""DynamicTable"" & ""_"" & ActiveSheet.Name).ListObject" & vbNewLine
    CodeText = CodeText & "        setLoadedRow Selection.Row - Selection.ListObject.Range.Row" & vbNewLine
    CodeText = CodeText & "    End If" & vbNewLine
    CodeText = CodeText & "End Sub"
    With WorksheetOfDynamicTable.CodeModule
        .AddFromString CodeText
    End With
End Sub

Function TableActiveRow() As Long
    On Error Resume Next
    Dim ActiveTableRow As Long
    TableActiveRow = Selection.row - Selection.ListObject.Range.row
End Function

Sub PopulateTableControl(MyTable As ListObject)
    '#INCLUDE TableActiveRow
    Application.ScreenUpdating = False
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = MyTable.parent
    Dim mapRow As Long
    Dim map As Range
    On Error GoTo eh
    For Each cell In MyTable.ListRows(TableActiveRow).Range
        Set map = Cells(ActiveSheet.Range("MappingRange" & "_" & ActiveSheet.Name).row, cell.Column)
        ws.Range(map.Value).Value = cell.Value
    Next cell
eh:
    Application.ScreenUpdating = True
End Sub

Function WorksheetOfDynamicTable() As VBComponent
    Dim vbProj As VBProject
    Set vbProj = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent.parent.VBProject
    Set WorksheetOfDynamicTable = vbProj.VBComponents(ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent.Name)
End Function

Sub CreateControlShapes()
    '#INCLUDE CreateShape
    '#INCLUDE ShapesFill
    '#INCLUDE MoveShapes
    '#INCLUDE TableSave
    '#INCLUDE TableDeleteLoaded
    '#INCLUDE TableDeleteMultiple
    '#INCLUDE InputStartNew
    '#INCLUDE TableReset
    Call CreateShape("InputStartNew", "NEW ENTRY")
    Call CreateShape("TableSave", "SAVE / UPDATE")
    Call CreateShape("TableDeleteLoaded", "DELETE LOADED")
    Call CreateShape("TableDeleteMultiple", "DELETE MULTIPLE")
    Call CreateShape("TableReset", "RESET EVERYTHING")
    Call ShapesFill
    Call MoveShapes
End Sub

Sub CreateShape(OnActionText As String, ShapeText As String)
    '#INCLUDE ShapeFactory
    '#INCLUDE ShapeLeftAfterLastColumn
    Dim LoopShape As Shape
    Dim LastShape As Shape
    Dim ShapesCount As Long
    Dim myShapeControl As Shape
    Select Case ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).parent.Shapes.count
        Case Is = 0
            Call ShapeFactory(OnActionText, _
                              ShapeText, _
                              ShapeLeftAfterLastColumn(ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name)), _
                              ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).top)
        Case Else
            For Each LoopShape In ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).parent.Shapes
                If LoopShape.AutoShapeType = msoShapeRoundedRectangle Then
                    Set LastShape = LoopShape
                    ShapesCount = ShapesCount + 1
                End If
            Next LoopShape
            If ShapesCount = 0 Then
                Call ShapeFactory(OnActionText, _
                                  ShapeText, _
                                  ShapeLeftAfterLastColumn(ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name)), _
                                  ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).top)
            Else
                Call ShapeFactory(OnActionText, _
                                  ShapeText, _
                                  LastShape.left, _
                                  LastShape.top + LastShape.Height + 10)
            End If
    End Select
End Sub

Function ShapeFactory(ShapeOnAction As String, _
                      ShapeText As String, _
                      ShapeLeft As Long, _
                      ShapeTop As Long) _
        As Shape
    Set ShapeFactory = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ShapeLeft, _
        ShapeTop, 50, 50)
    '#INCLUDE AddShape
    With ShapeFactory
        .OnAction = ShapeOnAction
        .Fill.BackColor.RGB = vbBlue
        With .ThreeD
            .BevelTopType = msoBevelCircle
            .BevelTopInset = 6
            .BevelTopDepth = 6
        End With
        With .TextFrame
            .Characters.TEXT = ShapeText
            .AutoSize = True
            .HorizontalAlignment = xlHAlignCenter
            .VerticalAlignment = xlVAlignCenter
            .Characters.Font.color = vbWhite
            .Characters.Font.Bold = True
            .Characters.Font.Size = 12
        End With
        .left = ShapeLeft
        .top = ShapeTop
        .Placement = xlFreeFloating
    End With
End Function

Function ShapeLeftAfterLastColumn(rng As Range) As Long
    ShapeLeftAfterLastColumn = rng.OFFSET(0, rng.Columns.count).left + 50
End Function

Sub ShapesFill()
    On Error GoTo 0
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If shp.AutoShapeType = msoShapeRoundedRectangle Then
            With shp.Fill
                .visible = msoTrue
                .Transparency = 0
                .Solid
                If shp.TextFrame.Characters.TEXT Like "*DELETE*" Then
                    .ForeColor.RGB = RGB(128, 0, 0)
                ElseIf shp.TextFrame.Characters.TEXT Like "*SAVE*" Then
                    .ForeColor.RGB = RGB(0, 128, 0)
                ElseIf shp.TextFrame.Characters.TEXT Like "*RESET*" Then
                    .ForeColor.RGB = vbBlack
                ElseIf shp.TextFrame.Characters.TEXT Like "*NEW*" Then
                    .ForeColor.RGB = RGB(0, 0, 128)
                Else
                    .ForeColor.RGB = RGB(0, 0, 128)
                End If
            End With
        End If
    Next
End Sub

Sub MoveShapes()
    '#INCLUDE LargestShapeWidth
    LargestShapeWidth
    Dim GroupedShapes
    ActiveSheet.Shapes.SelectAll
    Set GroupedShapes = Selection.Group
    GroupedShapes.Name = "ControlGroup" & "_" & ActiveSheet.Name
    ActiveSheet.Shapes("ControlGroup" & "_" & ActiveSheet.Name).Placement = xlFreeFloating
    ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Cells(1, 1).OFFSET(0, -1).ColumnWidth = _
                                                                                                     ActiveSheet.Shapes("ControlGroup" & "_" & ActiveSheet.Name).Width * 0.2
    ActiveSheet.Shapes("ControlGroup" & "_" & ActiveSheet.Name).left = _
                                                                     ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Cells(1, 1).OFFSET(0, -1).left + 5
End Sub

Function LargestShapeWidth() As Long
    Dim shp As Shape
    For i = 1 To ActiveSheet.Shapes.count
        If i = 1 Then
            LargestShapeWidth = ActiveSheet.Shapes(i).Width
        ElseIf ActiveSheet.Shapes(i).Width > ActiveSheet.Shapes(i - 1).Width Then
            LargestShapeWidth = ActiveSheet.Shapes(i).Width
        End If
    Next i
    For Each shp In ActiveSheet.Shapes
        shp.Width = LargestShapeWidth
    Next shp
End Function

Sub TableSave()
    '#INCLUDE InputStartNew
    Application.ScreenUpdating = False
    Dim tbl As ListObject
    Set tbl = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject
    Dim newRow As ListRow
    Dim RowLoaded As Variant
    RowLoaded = ActiveWorkbook.Names("LoadedRow" & "_" & ActiveSheet.Name)
    If RowLoaded = "=0" Then
        Set newRow = tbl.ListRows.Add
    Else
        Dim numOfNameRow As Long
        numOfNameRow = Mid(RowLoaded, 2)
        Set newRow = tbl.ListRows(numOfNameRow)
    End If
    Dim cell As Range
    For Each cell In newRow.Range
        cell.Value = Range(Cells(ActiveSheet.Range("MappingRange" & "_" & ActiveSheet.Name).row, cell.Column).Value).Value
    Next cell
    Call InputStartNew
    Application.ScreenUpdating = True
End Sub

Sub TableDeleteLoaded()
    '#INCLUDE InputStartNew
    Dim ws As Worksheet
    Set ws = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent
    If ActiveWorkbook.Names("LoadedRow" & "_" & ActiveSheet.Name) = "=0" Then Exit Sub
    Dim tbl As ListObject
    Set tbl = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject
    tbl.Range.AutoFilter
    ws.rows.Hidden = False
    tbl.ListRows(Mid(ActiveWorkbook.Names("LoadedRow" & "_" & ActiveSheet.Name), 2)).Delete
    Call InputStartNew
    tbl.Range.AutoFilter
End Sub

Sub TableDeleteMultiple()
    '#INCLUDE TableAutofilterRemove
    '#INCLUDE QuickSort
    '#INCLUDE TableSelectedRows
    '#INCLUDE InputStartNew
    Call TableAutofilterRemove
    Dim arr
    arr = TableSelectedRows
    If IsEmpty(arr) Then Exit Sub
    Call QuickSort(arr)
    Dim i As Long
    For i = UBound(arr) To LBound(arr) Step -1
        ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject.ListRows(arr(i)).Delete
    Next
    Call InputStartNew
End Sub

Sub TableAutofilterRemove()
    ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject.Range.AutoFilter
End Sub

Public Sub QuickSort(ByRef SortArray As Variant, Optional lngMin As Long = -1, Optional lngMax As Long = -1)
    '#INCLUDE SortArray
    On Error Resume Next
    Dim i As Long
    Dim j As Long
    Dim varMid As Variant
    Dim varX As Variant
    If IsEmpty(SortArray) Then
        Exit Sub
    End If
    If InStr(TypeName(SortArray), "()") < 1 Then
        Exit Sub
    End If
    If lngMin = -1 Then
        lngMin = LBound(SortArray)
    End If
    If lngMax = -1 Then
        lngMax = UBound(SortArray)
    End If
    If lngMin >= lngMax Then
        Exit Sub
    End If
    i = lngMin
    j = lngMax
    varMid = Empty
    varMid = SortArray((lngMin + lngMax) \ 2)
    If IsObject(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsEmpty(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf IsNull(varMid) Then
        i = lngMax
        j = lngMin
    ElseIf varMid = "" Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) = vbError Then
        i = lngMax
        j = lngMin
    ElseIf VarType(varMid) > 17 Then
        i = lngMax
        j = lngMin
    End If
    While i <= j
        While SortArray(i) < varMid And i < lngMax
            i = i + 1
        Wend
        While varMid < SortArray(j) And j > lngMin
            j = j - 1
        Wend
        If i <= j Then
            varX = SortArray(i)
            SortArray(i) = SortArray(j)
            SortArray(j) = varX
            i = i + 1
            j = j - 1
        End If
    Wend
    If (lngMin < j) Then Call QuickSort(SortArray, lngMin, j)
    If (i < lngMax) Then Call QuickSort(SortArray, i, lngMax)
End Sub

Sub TableShowAll()
    ActiveSheet.rows.Hidden = False
End Sub

Function TableSelectedRows() As Variant
    If TypeName(Selection) <> "Range" Then Exit Function
    Dim MyTable As ListObject
    Set MyTable = ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject
    If Intersect(Selection, MyTable.DataBodyRange) Is Nothing Then Exit Function
    On Error GoTo eh
    Dim arr()
    If Selection.Cells.count = 1 Then
        ReDim Preserve arr(0)
        arr(0) = ActiveCell.row - ActiveCell.ListObject.Range.row
        TableSelectedRows = arr
        Exit Function
    End If
    Dim i As Long
    Dim cell As Range
    Dim SelectedRange As Range
    Set SelectedRange = Selection.SpecialCells(xlCellTypeVisible)
    Dim coll As New Collection
    i = 0
    On Error Resume Next
    For Each cell In SelectedRange
        If Not Intersect(cell, MyTable.DataBodyRange) Is Nothing Then
            coll.Add cell.row - cell.ListObject.Range.row, CStr(cell.row)
            ReDim Preserve arr(i)
            arr(i) = coll(i + 1)
            i = i + 1
        End If
    Next cell
    On Error GoTo 0
    TableSelectedRows = arr
eh:
End Function

Sub InputStartNew()
    '#INCLUDE createNamedRange
    Application.ScreenUpdating = False
    Call createNamedRange("LoadedRow" & "_" & ActiveSheet.Name, _
                          ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).parent, 0)
    ActiveSheet.Range("ControlInputFields" & "_" & ActiveSheet.Name).ClearContents
    ActiveSheet.Range("TableControl" & "_" & ActiveSheet.Name).Cells(1, 2).Select
    Application.ScreenUpdating = True
End Sub

Sub TableReset()
    '#INCLUDE DeleteNames
    '#INCLUDE DeleteShapes
    '#INCLUDE DeleteActiveSheetCodemodule
    Call DeleteActiveSheetCodemodule
    ActiveSheet.Range("DynamicTable" & "_" & ActiveSheet.Name).ListObject.Unlist
    ActiveSheet.Cells.ClearContents
    Call DeleteShapes
    Call DeleteNames
End Sub

Sub DeleteNames()
    Dim xName As Variant
    For Each xName In Array("ControlInputFields" & "_" & ActiveSheet.Name, _
                            "LoadedRow" & "_" & ActiveSheet.Name, _
                            "MappingRange" & "_" & ActiveSheet.Name, _
                            "TableControl" & "_" & ActiveSheet.Name, _
                            "TableHeaderRange" & "_" & ActiveSheet.Name)
        ActiveWorkbook.Names(xName).Delete
    Next xName
End Sub

Sub DeleteShapes()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        shp.Delete
    Next
End Sub

Sub DeleteActiveSheetCodemodule()
    '#INCLUDE WorksheetOfDynamicTable
    With WorksheetOfDynamicTable.CodeModule
        .DeleteLines 1, .CountOfLines
    End With
End Sub

Sub TableSort(MyTable As ListObject, _
              myKey1 As Range, _
              Optional myKey2 As Range, _
              Optional myKey3 As Range)
    With MyTable.Sort
        .header = xlYes
        .SortFields.clear
        .SortFields.Add key:=myKey1, SortOn:=xlSortOnValues
        If Not myKey2 Is Nothing Then
            .SortFields.Add key:=myKey2, SortOn:=xlSortOnValues
        End If
        If Not myKey3 Is Nothing Then
            .SortFields.Add key:=myKey3, SortOn:=xlSortOnValues
        End If
        .Apply
    End With
End Sub


