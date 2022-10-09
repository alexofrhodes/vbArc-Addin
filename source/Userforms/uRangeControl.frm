VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uRangeControl 
   Caption         =   "RangeControl"
   ClientHeight    =   9492.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   13332
   OleObjectBlob   =   "uRangeControl.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uRangeControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uRangeControl
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Private seper As String

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Sub CommandButton20_Click()
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    Dim ans As Long
    ans = MsgBox("By column?", vbYesNoCancel)
    If ans = vbCancel Then Exit Sub
    
    Dim columnDataArray, rowDataArray
    Dim mergedData As String
    Dim firstCell As Range, columnRange As Range, rowRange As Range
    Dim columnCounter As Long, rowCounter
    Dim columnCell, rowCell
    Dim vArea
    For Each vArea In Selection.Areas
        If ans = vbYes Then
            For columnCounter = 1 To vArea.Columns.count
                mergedData = ""
                Set columnRange = vArea.Columns(columnCounter)
                columnDataArray = columnRange.Value
                columnRange.Cells.ClearContents
                For Each columnCell In columnDataArray
                    mergedData = IIf(mergedData = "", columnCell, mergedData & vbNewLine & columnCell)
                Next
                Set firstCell = vArea.Columns(columnCounter).Cells(1)
                firstCell.Value = mergedData
                firstCell.EntireRow.AutoFit
            Next
        Else
            For rowCounter = 1 To vArea.rows.count
                mergedData = ""
                Set rowRange = vArea.rows(rowCounter)
                rowDataArray = rowRange.Value
                rowRange.Cells.ClearContents
                For Each rowCell In rowDataArray
                    mergedData = IIf(mergedData = "", rowCell, mergedData & vbNewLine & rowCell)
                Next
                Set firstCell = vArea.rows(rowCounter).Cells(1)
                firstCell.Value = mergedData
                firstCell.EntireRow.AutoFit
            Next
        End If
    Next
End Sub

Private Sub CommandButton21_Click()
    Dim del As String
    Dim cell As Range, rng As Range
    Set rng = Application.InputBox(Prompt:="Select a range:", _
                                   title:="Sort values inside cells", _
                                   Default:=Selection.Address, Type:=8)
    del = InputBox(Prompt:="Delimiting character: (leave empty for newline)", _
                   title:="Sort values in a single cell")
    If del = "" Then del = vbNewLine
    For Each cell In rng
        RotateArray Split(cell.Value, del)
    Next
End Sub

Private Sub CommandButton22_Click()
    ShakeTableWand
End Sub

Private Sub CommandButton23_Click()
    AddReadmeToWorkbook
End Sub

Private Sub UserForm_Initialize()
    seper = vbNewLine
    Me.Width = 300
    Me.Height = 280
    Dim anc As MSForms.control

    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            'c.Caption = ""
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.visible = False
                If InStr(1, c.Tag, "anchor") > 0 Then
                    On Error Resume Next
                    Set anc = Me.Controls("Anchor" & Mid(c.Tag, InStr(1, c.Tag, "Anchor", vbTextCompare) + Len("Anchor"), 2))
                    If anc Is Nothing Then Stop
                    On Error GoTo 0
                    c.top = anc.top        'Anchor01.Top
                    c.left = anc.left        ' Anchor01.Left
                    Set anc = Nothing
                End If
            End If
        End If
    Next
End Sub

Private Sub UserForm_Activate()
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
    
    CenterMouseOver Me, Label5
    LeftClick

    'ResizeUserformToFitControls Me

End Sub

Private Sub Emitter_LabelMouseOut(label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then
        If label.BackColor <> &H80B91E Then label.BackColor = &H534848
    End If
End Sub

Private Sub Emitter_LabelMouseOver(label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then
        If label.BackColor <> &H80B91E Then label.BackColor = &H808080
    End If
End Sub

Sub Emitter_LabelClick(ByRef label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then Reframe Me, label
End Sub

Private Sub cAddCol_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Expand(0, 1)
End Sub

Private Sub cADDrow_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Expand(1, 0)
End Sub

Private Sub cCenterAcross_Click()
    ConvertMergedCellsToCenterAcross
End Sub

Sub ConvertMergedCellsToCenterAcross()
    Dim c As Range
    Dim mergedRange As Range
    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    For Each c In ActiveSheet.UsedRange.SpecialCells(xlCellTypeConstants)
        If c.MergeCells = True And c.MergeArea.rows.count = 1 Then
            Set mergedRange = c.MergeArea
            mergedRange.UnMerge
            mergedRange.HorizontalAlignment = xlCenterAcrossSelection
        End If
    Next
End Sub

Private Sub CommandButton1_Click()
    If oFirstblank.Value = True Then
        Call GoToBlankFirstInColumn
    ElseIf oLastblank.Value = True Then
        Call GoToBlankLastWithinColumn
    ElseIf oNextBlank.Value = True Then
        Call GoToBlankNext
    ElseIf oNewRow.Value = True Then
        Call GoToBlankAfterColumn
    End If
End Sub

Private Sub CommandButton10_Click()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    For Each cell In Selection
        cell.TEXT = GreekToLatin(cell.TEXT)
    Next
End Sub

Private Sub CommandButton11_Click()
    CircleBoxADD
End Sub

Private Sub CommandButton12_Click()
    CircleBoxREMOVE
End Sub

Private Sub CommandButton14_Click()
    RemoveArrows
End Sub

Private Sub CommandButton15_Click()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If Selection.Areas.count <> 2 Then Exit Sub
    If Selection.Areas(1).Cells.count <> 1 Then Exit Sub
    If Selection.Areas(2).Cells.count <> 1 Then Exit Sub
    DrawArrows Selection.Areas(1).Cells, Selection.Areas(2).Cells, , Switch(oArrowDouble.Value = True, "DOUBLE", oArrowLine.Value = True, "LINE", oArrowSingle, "SINGLE")
End Sub

Private Sub CommandButton16_Click()
    HideShapes "myArrow"
End Sub

Private Sub CommandButton17_Click()
    ShowShapes "myArrow"
End Sub

Sub HideShapes(shapeName As String)
    For Each shp In ActiveSheet.Shapes
        If UCase(shp.Name) = UCase(shapeName) Then shp.visible = False
    Next shp
End Sub

Sub ShowShapes(shapeName As String)
    For Each shp In ActiveSheet.Shapes
        If UCase(shp.Name) = UCase(shapeName) Then shp.visible = True
    Next shp
End Sub

Private Sub CommandButton18_Click()
    HideShapes "myCircle"
End Sub

Private Sub CommandButton19_Click()
    ShowShapes "myCircle"
End Sub

Private Sub CommandButton2_Click()
    AddBorders
End Sub

Private Sub CommandButton3_Click()
    ClearBorders
End Sub

Private Sub CommandButton4_Click()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim rng As Range, i As Long
    Set rng = Selection
    Dim areaCount As Long
    areaCount = rng.Areas.count
    For i = 1 To areaCount
        rng.Areas(i).HorizontalAlignment = xlCenterAcrossSelection
    Next
End Sub

Private Sub CommandButton5_Click()
    CopyWorksheetsIncludingTables
End Sub

Private Sub CommandButton6_Click()
    deleteBlankWorksheets
End Sub

Private Sub CommandButton7_Click()
    SortValuesInCell
End Sub

Private Sub CommandButton8_Click()
    getEmailCollection
End Sub

Private Sub CommandButton9_Click()
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim cell As Range
    For Each cell In Selection
        cell.Value = LatinToGreek(cell.TEXT)
    Next
End Sub

Private Sub cREMcol_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Expand(0, -1)
End Sub

Private Sub cREMrow_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Expand(-1, 0)
End Sub

Private Sub cFlipH_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlipHorizontally
End Sub

Private Sub cFlipV_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FlipVertically
End Sub

Private Sub cSwap_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    SwapSelectedRanges
End Sub

'
'Private Sub cIup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'    MoveCell ("up")
'End Sub
'
'Private Sub cIdown_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'    MoveCell ("down")
'End Sub
'
'Private Sub cIleft_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'    MoveCell ("left")
'End Sub
'
'Private Sub cIright_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)
'    MoveCell ("right")
'End Sub

Private Sub cSup_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveSelection ("up")
End Sub

Private Sub cSdown_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveSelection ("down")
End Sub

Private Sub cSleft_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveSelection ("left")
End Sub

Private Sub cSright_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    MoveSelection ("right")
End Sub

Private Sub cCopy_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error GoTo eh
    Selection.Copy
eh:                     MsgBox "copy one area only"
End Sub

Private Sub Image2_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    uDEV.Show
End Sub

Private Sub Image3_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call MergeCustom
End Sub

Private Sub Image4_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call UnmergeCustom
End Sub

Private Sub Image5_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SelectAllMergedCells
End Sub

Private Sub Image6_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select a range"
        Exit Sub
    End If
    If oAppend.Value = False And oPretend.Value = False And oReplace.Value = False Then
        MsgBox "Select an option"
        Exit Sub
    End If
    Application.ScreenUpdating = False
    Dim str As String
    str = tString
    Dim cell As Range
    Dim rng As Range
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    If oReplace.Value = True Then
        If Selection.count = 1 Then
            ActiveCell.Value = str
        ElseIf Selection.count > 1 Then
            rng.Value = str
        End If
    ElseIf oPretend.Value = True Then
        For Each cell In rng
            cell.Value = str & cell.Value
        Next cell
    ElseIf oAppend.Value = True Then
        For Each cell In rng
            cell.Value = cell.Value & str
        Next cell
    End If
    Application.ScreenUpdating = True
End Sub

Private Sub tCol_Change()
    If tCol.TEXT = vbNullString Or Not IsNumeric(tCol.TEXT) Then tCol.TEXT = "1"
    If CInt(tCol.TEXT) <= 0 Then tCol.TEXT = "1"
    ResizeByTextbox 0, CInt(tCol.TEXT)
End Sub

Private Sub tMergeLink_Change()
    seper = IIf(tMergeLink.TEXT <> "", tMergeLink.TEXT, vbNewLine)
End Sub

Private Sub tRow_Change()
    If tRow.TEXT = vbNullString Or Not IsNumeric(tRow.TEXT) Then tRow.TEXT = "1"
    If CInt(tRow.TEXT) <= 0 Then tRow.TEXT = "1"
    ResizeByTextbox CInt(tRow.TEXT), 0
End Sub

Sub AddBorders()
    If chTop.Value = True Then Selection.BORDERS(xlEdgeTop).LineStyle = xlContinuous
    If chBottom.Value = True Then Selection.BORDERS(xlEdgeBottom).LineStyle = xlContinuous
    If chLeft.Value = True Then Selection.BORDERS(xlEdgeLeft).LineStyle = xlContinuous
    If chRight.Value = True Then Selection.BORDERS(xlEdgeRight).LineStyle = xlContinuous
    If chHorizontal.Value = True Then Selection.BORDERS(xlInsideHorizontal).LineStyle = xlContinuous
    If chVertical.Value = True Then Selection.BORDERS(xlInsideVertical).LineStyle = xlContinuous
End Sub

Sub ClearBorders()
    Dim myBorders() As Variant, item As Variant
    myBorders = Array(xlEdgeLeft, xlEdgeTop, xlEdgeBottom, xlEdgeRight, xlInsideVertical, xlInsideHorizontal, xlInsideVertical)
    For Each item In myBorders
        With Selection.BORDERS(item)
            .LineStyle = xlNone
        End With
    Next
End Sub

Sub getEmailCollection()
    Dim output As Collection
    Set output = New Collection
    Dim i As Long
    Dim str As String
    Dim rng As Range
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    Call ExtractEmailCollection(rng, output)
    For i = 1 To output.count
        If str = vbNullString Then
            str = output(i)
        Else
            str = str & vbNewLine & output(i)
        End If
    Next i
    CLIP str
    Set output = Nothing
    MsgBox "Copied to clipboard"
End Sub

Sub ExtractEmailCollection(rng As Range, output As Collection)
    Dim aEmails               As Variant
    Dim i                     As Long
    Dim cell As Range
    For Each cell In rng
        aEmails = ExtractEmailAddresses(cell.Value)
        If IsNull(aEmails) = False Then
            For i = 0 To UBound(aEmails)
                output.Add aEmails(i)
            Next i
        End If
    Next cell
End Sub

Public Function ExtractEmailAddresses(ByVal sInput As Variant) As Variant
    On Error GoTo Error_Handler
    Dim oRegEx                As Object
    Dim oMatches              As Object
    Dim oMatch                As Object
    Dim sEmail                As String
    If Not IsNull(sInput) Then
        Set oRegEx = CreateObject("vbscript.regexp")
        With oRegEx
            .Pattern = "([a-zA-ZF0-9\u00C0-\u017F._-]+@[a-zA-Z0-9\u00C0-\u017F._-]+\.[a-zA-Z0-9\u00C0-\u017F_-]+)"
            .Global = True
            .IgnoreCase = True
            .MultiLine = True
            Set oMatches = .Execute(sInput)
        End With
        For Each oMatch In oMatches
            sEmail = oMatch.Value & "," & sEmail
        Next oMatch
        If Right(sEmail, 1) = "," Then sEmail = left(sEmail, Len(sEmail) - 1)
        ExtractEmailAddresses = Split(sEmail, ",")
    Else
        ExtractEmailAddresses = Null
    End If
Error_Handler_Exit:
    On Error Resume Next
    If Not oMatch Is Nothing Then Set oMatch = Nothing
    If Not oMatches Is Nothing Then Set oMatches = Nothing
    If Not oRegEx Is Nothing Then Set oRegEx = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: ExtractEmailAddresses" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Sub SwapSelectedRanges()
    Dim rng As Range
    Dim tempRng As Variant
    Dim areaCount As Long
    Dim areaRows As Long
    Dim areaCols As Long
    Dim i As Integer
    Dim j As Integer
    Set rng = Selection
    areaCount = rng.Areas.count
    If areaCount < 2 Then
        MsgBox "Please select atleast two ranges."
        Exit Sub
    End If
    areaRows = rng.Areas(1).rows.count
    areaCols = rng.Areas(1).Columns.count
    For i = 2 To areaCount
        If rng.Areas(i).rows.count <> areaRows Or _
                                   rng.Areas(i).Columns.count <> areaCols Then
            MsgBox "All ranges must have the same number of rows and columns."
            Exit Sub
        End If
    Next i
    For j = 1 To areaCount - 1
        For i = 1 + j To areaCount
            If Not Intersect(rng.Areas(i), rng.Areas(j)) Is Nothing Then
                MsgBox "Selected areas must not overlap."
            End If
        Next i
    Next j
    tempRng = rng.Areas(areaCount).Cells.Formula
    For i = areaCount To 2 Step -1
        rng.Areas(i).Cells.Formula = rng.Areas(i - 1).Cells.Formula
    Next i
    rng.Areas(1).Cells.Formula = tempRng
End Sub

Sub FlipHorizontally()
    Dim rng As Range
    Dim WorkRng As Range
    Dim arr As Variant
    Dim i As Integer, j As Integer, k As Integer
    On Error Resume Next
    xTitleId = "Flip Data Horizontally"
    Set WorkRng = Application.Selection
    arr = WorkRng.Formula
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For i = 1 To UBound(arr, 1)
        k = UBound(arr, 2)
        For j = 1 To UBound(arr, 2) / 2
            xTemp = arr(i, j)
            arr(i, j) = arr(i, k)
            arr(i, k) = xTemp
            k = k - 1
        Next
    Next
    WorkRng.Formula = arr
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub FlipVertically()
    Dim rng As Range
    Dim WorkRng As Range
    Dim arr As Variant
    Dim i As Integer, j As Integer, k As Integer
    On Error Resume Next
    xTitleId = "Flip columns vertically"
    Set WorkRng = Application.Selection
    arr = WorkRng.Formula
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    For j = 1 To UBound(arr, 2)
        k = UBound(arr, 1)
        For i = 1 To UBound(arr, 1) / 2
            xTemp = arr(i, j)
            arr(i, j) = arr(k, j)
            arr(k, j) = xTemp
            k = k - 1
        Next
    Next
    WorkRng.Formula = arr
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Sub MergeCustom()
    If TypeName(Selection) <> "Range" Then Exit Sub
    If seper = "" Then seper = Chr(10)
    Dim sRange As Range
    Dim cell As Range
    Dim TmpStr As String
    Set sRange = Selection
    Dim element As Variant
    For Each element In sRange.Areas
        If element.MergeCells = True Then GoTo Skip
        For Each cell In element.Cells
            TmpStr = TmpStr & seper & CStr(cell.Value)
        Next cell
        TmpStr = Mid(TmpStr, Len(seper) + 1)
        Application.DisplayAlerts = False
        With element
            .ClearContents
            .MERGE
            .Value = TmpStr
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
        End With
        TmpStr = ""
Skip:
    Next
    Application.DisplayAlerts = True
End Sub

Sub UnmergeCustom()
    If seper = "" Then seper = Chr(10)
    Dim sRange As Range
    Dim cell As Range
    Dim rmpStr As String
    Dim counter As Long
    Dim arr() As String
    Set sRange = Selection
    Dim element As Variant
    For Each element In sRange.Areas
        If element.MergeCells = False Then GoTo Skip
        TmpStr = CStr(element.Cells(1).Value)
        arr = Split(TmpStr, seper)
        element.ClearContents
        element.MergeCells = False
        counter = 0
        On Error Resume Next
        For Each cell In element.Cells
            cell.Value = CStr(arr(counter))
            counter = counter + 1
        Next cell
Skip:
        TmpStr = ""
    Next
    On Error GoTo 0
End Sub

Sub SelectAllMergedCells()
    Dim c As Range
    Dim mergedCells As Range
    Dim fullRange As Range
    Dim rangeDescription As String
    Dim rng As Range

    If Selection.Cells.count > 1 Then
        Set fullRange = Selection
        rangeDescription = "selected cells"
    Else
        Set fullRange = ActiveSheet.UsedRange
        rangeDescription = "active range"
    End If
    For Each c In fullRange
        If c.MergeCells = True Then
            If mergedCells Is Nothing Then
                Set mergedCells = c
            Else
                Set mergedCells = Union(mergedCells, c)
            End If
        End If
    Next
    If Not mergedCells Is Nothing Then
        mergedCells.Select
    Else
        MsgBox "There are no merged cells in the " _
             & rangeDescription & ": " & fullRange.Address
    End If
End Sub

Sub MoveSelection(Direction As String)
    On Error GoTo eh
    Dim myRange As Range
    Set myRange = Selection
    Select Case Direction
        Case Is = "up"
            myRange.OFFSET(-1, 0).Select
        Case Is = "down"
            myRange.OFFSET(1, 0).Select
        Case Is = "left"
            myRange.OFFSET(0, -1).Select
        Case Is = "right"
            myRange.OFFSET(0, 1).Select
    End Select
eh:
End Sub

Sub GoToBlankFirstInColumn()
    If TypeName(Selection) <> "Range" Or _
                           Selection.Cells.count > 1 Then
        MsgBox "Choose 1 cell"
        Exit Sub
    End If
    Dim sourceCol As Integer, rowCount As Integer, currentRow As Integer
    Dim currentRowValue As String
    sourceCol = ActiveCell.Column
    rowCount = Cells(rows.count, sourceCol).End(xlUp).row
    For currentRow = 1 To rowCount + 1
        currentRowValue = Cells(currentRow, sourceCol).Value
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
            Cells(currentRow, sourceCol).Select
            Exit Sub
        End If
    Next
End Sub

Sub GoToBlankNext()
    If TypeName(Selection) <> "Range" Or _
                           Selection.Cells.count > 1 Then
        MsgBox "Choose 1 cell"
        Exit Sub
    End If
    Dim sourceCol As Integer, rowCount As Integer, currentRow As Integer
    Dim currentRowValue As String
    sourceCol = ActiveCell.Column
    rowCount = Cells(rows.count, sourceCol).End(xlUp).row
    For currentRow = ActiveCell.row + 1 To rowCount + 1
        currentRowValue = Cells(currentRow, sourceCol).Value
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
            Cells(currentRow, sourceCol).Select
            Exit Sub
        End If
    Next
End Sub

Sub GoToBlankLastWithinColumn()
    If TypeName(Selection) <> "Range" Or _
                           Selection.Cells.count > 1 Then
        MsgBox "Choose 1 cell"
        Exit Sub
    End If
    Dim sourceCol As Integer, rowCount As Integer, currentRow As Integer
    Dim currentRowValue As String
    sourceCol = ActiveCell.Column
    rowCount = Cells(rows.count, sourceCol).End(xlUp).row
    For currentRow = rowCount To 1 Step -1
        currentRowValue = Cells(currentRow, sourceCol).Value
        If IsEmpty(currentRowValue) Or currentRowValue = "" Then
            Cells(currentRow, sourceCol).Select
            Exit Sub
        End If
    Next
End Sub

Sub GoToBlankAfterColumn()
    Application.ScreenUpdating = False
    If TypeName(Selection) <> "Range" Or _
                           Selection.Cells.count > 1 Then
        MsgBox "Choose 1 cell"
        Exit Sub
    End If
    Dim LastRow As Long
    LastRow = Cells(rows.count, ActiveCell.Column).End(xlUp).row
    Cells(LastRow, ActiveCell.Column).OFFSET(1, 0).Select
    With ActiveCell.OFFSET(-1)
        If .TEXT = "" Then .Select
    End With
    Application.ScreenUpdating = True
End Sub

Sub Expand(ByVal ExpandRows As Double, ByVal ExpandColumns As Double)
    Application.ScreenUpdating = False
    Dim r As Range
    Dim i As Long
    Dim RowSize As Long
    Dim ColSize As Long
    For i = 1 To Selection.Areas.count
        ColSize = Selection.Areas(i).Columns.count
        RowSize = Selection.Areas(i).rows.count
        If ValidExpand(RowSize + ExpandRows, ColSize + ExpandColumns) Then
            If r Is Nothing Then
                Set r = Range(Selection.Areas(i).RESIZE(RowSize + ExpandRows, ColSize + ExpandColumns).Address)
            Else
                Set r = Application.Union(r, Range(Selection.Areas(i).RESIZE(RowSize + ExpandRows, ColSize + ExpandColumns).Address))
            End If
        End If
    Next i
    If Not r Is Nothing Then
        ActiveWorkbook.ActiveSheet.Range("A1").Select
        r.Select
    End If
    Application.ScreenUpdating = True
End Sub

Function ValidExpand(Optional a As Long, Optional b As Long) As Boolean
    If a >= 1 And b >= 1 Then ValidExpand = True
End Function

Sub ADDcol()
    Call Expand(0, 1)
End Sub

Sub ADDrow()
    Call Expand(1, 0)
End Sub

Sub REMcol()
    Call Expand(0, -1)
End Sub

Sub REMrow()
    Call Expand(-1, 0)
End Sub

Sub ResizeByTextbox(Optional ExpandRows As Long, Optional ExpandColumns As Long)
    Application.ScreenUpdating = False
    Dim r As Range
    Dim i As Long
    Dim RowSize As Long
    Dim ColSize As Long
    For i = 1 To Selection.Areas.count
        If ExpandRows = 0 Then
            ExpandRows = Selection.Areas(i).rows.count
        End If
        If ExpandColumns = 0 Then
            ExpandColumns = Selection.Areas(i).Columns.count
        End If
        If ValidExpand(ExpandRows, ExpandColumns) Then
            If r Is Nothing Then
                Set r = Range(Selection.Areas(i).RESIZE(ExpandRows, ExpandColumns).Address)
            Else
                Set r = Application.Union(r, Range(Selection.Areas(i).RESIZE(ExpandRows, ExpandColumns).Address))
            End If
        End If
    Next i
    If Not r Is Nothing Then
        ActiveWorkbook.ActiveSheet.Range("A1").Select
        r.Select
    End If
    Application.ScreenUpdating = True
End Sub

Sub deleteBlankWorksheets()
    Dim ws As Worksheet
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    For Each ws In Application.Worksheets
        If Application.WorksheetFunction.CountA(ws.UsedRange) = 0 Then
            ws.Delete
        End If
    Next
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
End Sub

Sub CopyWorksheetsIncludingTables()
    Dim TheActiveWindow As Window
    Dim TempWindow As Window

    With ActiveWorkbook
        Set TheActiveWindow = ActiveWindow
        Set TempWindow = .NewWindow
        '.Sheets(Array("Sheet1", "sheet2")).Copy
        '///If you want to copy the selected sheets use
        TheActiveWindow.SelectedSheets.Copy
    End With
    TempWindow.Close
End Sub

Sub SortValuesInCell()
    Dim rng As Range
    Dim cell As Range
    Dim del As String
    Dim arr As Variant

    On Error Resume Next
    Set rng = Application.InputBox(Prompt:="Select a range:", _
                                   title:="Sort values inside cells", _
                                   Default:=Selection.Address, Type:=8)
    del = InputBox(Prompt:="Delimiting character: (leave empty for newline)", _
                   title:="Sort values in a single cell")
    On Error GoTo 0
    If del = "" Then del = vbNewLine
    For Each cell In rng
        arr = Split(cell, del)
        SelectionSort arr
        cell = Join(arr, del)
    Next cell

End Sub

Sub SelectionSort(tempArray As Variant)
    'use in sub

    Dim MaxVal As Variant
    Dim MaxIndex As Integer
    Dim i As Integer, j As Integer

    ' Step through the elements in the array starting with the
    ' last element in the array.
    For i = UBound(tempArray) To 0 Step -1

        ' Set MaxVal to the element in the array and save the
        ' index of this element as MaxIndex.
        MaxVal = tempArray(i)
        MaxIndex = i

        ' Loop through the remaining elements to see if any is
        ' larger than MaxVal. If it is then set this element
        ' to be the new MaxVal.
        For j = 0 To i
            If tempArray(j) > MaxVal Then
                MaxVal = tempArray(j)
                MaxIndex = j
            End If
        Next j

        ' If the index of the largest element is not i, then
        ' exchange this element with element i.
        If MaxIndex < i Then
            tempArray(MaxIndex) = tempArray(i)
            tempArray(i) = MaxVal
        End If
    Next i

End Sub

Public Function LatinToGreek(TEXT4TRANS As String) As String
    'Convert characters to Greek
    Dim strFromLetter, strToLetter As String
    Dim intLetterPosition, intGreekLetterPosition As Long
    strFromLetter = "abgdezhuiklmnjoprstyfxcvABGDEZHUIKLMNJOPRSTYFXCV"
    'strToLetter = "áâãäåæçèéêëìíîïðñóôõö÷øùÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ"
    strToLetter = "ÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÓÔÕÖ×ØÙ"

    For intLetterPosition = 1 To Len(TEXT4TRANS)
        intGreekLetterPosition = InStr(1, strFromLetter, Mid(TEXT4TRANS, intLetterPosition, 1))
        If intGreekLetterPosition > 0 Then
            Mid(TEXT4TRANS, intLetterPosition, 1) = Mid(strToLetter, intGreekLetterPosition, 1)
        End If
    Next
    LatinToGreek = TEXT4TRANS
End Function

Function GreekToLatin(keimeno As String) As String
    'https://varlamis.wordpress.com/2011/05/31/word_GreekToLatin/
    '    Application.Volatile True
    Dim varr As Variant
    Dim inchar As Variant
    Dim exchar As Variant
    Dim pl As Integer, gr As Integer, lu As Integer
    Dim gramma As String
    pl = Len(keimeno)
 
    inchar = Array("Á", "Â", "Ã", "Ä", "Å", "Æ", "Ç", "È", "É", "Ê", "Ë", _
                   "Ì", "Í", "Î", "Ï", "Ð", "Ñ", "Ó", "Ô", "Õ", "Ö", "×", "Ø", "Ù", _
                   "¢", "¸", "¹", "º", "¼", "¾", "¿", "Ú", "Û", "À", "à", "ò")

    exchar = Array("A", "B", "G", "D", "E", "Z", "H", "8", "I", "K", "L", _
                   "M", "N", "KS", "O", "P", "R", "S", "T", "Y", "F", "X", "PS", "W", _
                   "A", "E", "H", "I", "O", "Y", "W", "I", "Y", "I", "Y", "S")
 
    ReDim varr(pl - 1)
    For gr = 1 To pl
        gramma = Mid(keimeno, gr, 1)
        For lu = LBound(inchar) To UBound(inchar)
            If UCase(gramma) = inchar(lu) Then gramma = exchar(lu): Exit For
        Next
        varr(gr - 1) = gramma
    Next
    GreekToLatin = Join(varr, "")

End Function

Function GetColorText(pRange As Range) As String
    Dim xOut As String
    Dim xValue As String
    Dim i As Long
    xValue = pRange.TEXT
    For i = 1 To VBA.Len(xValue)
        'If pRange.Characters(i, 1).Font.Color = vbRed Then
        'If pRange.Characters(i, 1).Font.Bold = true Then
        If pRange.Characters(i, 1).Font.color <> vbBlack Then
            xOut = xOut & VBA.Mid(xValue, i, 1)
        End If
    Next
    GetColorText = xOut
End Function

Function RemoveCharacters(TEXT As String, Remove As String) As String
    Dim X As Long
    RemoveCharacters = TEXT
    For X = 1 To Len(Remove)
        RemoveCharacters = Replace(RemoveCharacters, Mid(Remove, X, 1), "")
    Next
End Function

Private Sub DrawArrows(FromRange As Range, ToRange As Range, Optional RGBcolor As Long, Optional LineType As String)
    '---------------------------------------------------------------------------------------------------
    '---Script: DrawArrows------------------------------------------------------------------------------
    '---Created by: Ryan Wells -------------------------------------------------------------------------
    '---Date: 10/2015-----------------------------------------------------------------------------------
    '---Description: This macro draws arrows or lines from the middle of one cell to the middle --------
    '----------------of another. Custom endpoints and shape colors are suppported ----------------------
    '---------------------------------------------------------------------------------------------------

    Dim dleft1 As Double, dleft2 As Double
    Dim dtop1 As Double, dtop2 As Double
    Dim dheight1 As Double, dheight2 As Double
    Dim dwidth1 As Double, dwidth2 As Double
    dleft1 = FromRange.left
    dleft2 = ToRange.left
    dtop1 = FromRange.top
    dtop2 = ToRange.top
    dheight1 = FromRange.Height
    dheight2 = ToRange.Height
    dwidth1 = FromRange.Width
    dwidth2 = ToRange.Width
 
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, dleft1 + dwidth1 / 2, dtop1 + dheight1 / 2, dleft2 + dwidth2 / 2, dtop2 + dheight2 / 2).Select
    'format line
    Selection.Name = "myArrow"
    With Selection.ShapeRange.line
        .BeginArrowheadStyle = msoArrowheadNone
        .EndArrowheadStyle = msoArrowheadOpen
        .Weight = 1.75
        .Transparency = 0.5
        If UCase(LineType) = "DOUBLE" Then        'double arrows
            .BeginArrowheadStyle = msoArrowheadOpen
        ElseIf UCase(LineType) = "LINE" Then        'Line (no arows)
            .EndArrowheadStyle = msoArrowheadNone
        Else        'single arrow
            'defaults to an arrow with one head
        End If
        'color arrow
        If RGBcolor <> 0 Then
            .ForeColor.RGB = RGBcolor        'custom color
        Else
            .ForeColor.RGB = vbRed        'RGB(228, 108, 10)   'orange (DEFAULT)
        End If
    End With

End Sub

Sub RemoveArrows()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If shp.Name = "myArrow" Then
            shp.Delete
        End If
    Next
End Sub

Sub CircleBoxADD()
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select 1 or more ranges before running the macro."
        Exit Sub
    End If
    Dim MyOval As Shape
    Dim cell As Range
    Dim i As Long
    If oToCells.Value = True Then
        For Each cell In Selection
            addCircle cell
        Next
    Else
        For i = 1 To Selection.Areas.count
            addCircle Selection.Areas(i).Cells
        Next
    End If
End Sub

Sub addCircle(cell As Range)
    On Error GoTo cellisarea
    t = cell.MergeArea.top
    l = cell.MergeArea.left
    h = cell.MergeArea.Height
    w = cell.MergeArea.Width
    GoTo PASS
cellisarea:
    t = cell.top
    l = cell.left
    h = cell.Height
    w = cell.Width
PASS:
    On Error GoTo 0
    Set MyOval = ActiveSheet.Shapes.AddShape(msoShapeOval, l + 2, t + 2, w - 4, h - 4)
    With MyOval
        .Name = "myCircle"
        .Fill.visible = msoFalse
        .line.visible = msoTrue
        .line.ForeColor.RGB = RGB(255, 0, 0)
        .line.Transparency = 0
        .line.visible = msoTrue
        .line.Weight = 0.5
    End With
End Sub

Sub CircleBoxREMOVE()
    Dim shp As Shape
    For Each shp In ActiveSheet.Shapes
        If shp.Name = "myCircle" Then
            shp.Delete
        End If
    Next
End Sub

Sub SquareRange(rng As Range)
    Dim i As Integer
    If rng Is Nothing Then
        If TypeName(Selection) = "Range" Then
            Set rng = Selection
        Else
            Exit Sub
        End If
    End If
    For i = 1 To 4
        With rng
            .Columns.ColumnWidth = _
                                 .Columns("A").ColumnWidth / .Columns("A").Width * _
                                 .rows(1).Height
        End With
    Next
End Sub


