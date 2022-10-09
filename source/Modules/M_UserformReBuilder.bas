Attribute VB_Name = "M_UserformReBuilder"

Rem @Folder FormBuilder Declarations
Public UFSheet As Worksheet

Rem @Folder FormBuilder
Sub ExportUserform(ByRef uForm As VBComponent)
    '#INCLUDE GetModuleText
    '#INCLUDE WorkbookOfModule
    '#INCLUDE SortColumns
    '#INCLUDE PutFramesFirst
    Dim str As String
    Dim counter As Long
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim controlItem As MSForms.control
    Set UFSheet = ThisWorkbook.SHEETS("FormBuilder")
    UFSheet.Range("D:D").ClearContents
    UFSheet.Range("G:ZZ").ClearContents
    Application.ScreenUpdating = False
    RowNumber = 1
    ColumnNumber = 7
    Dim PropertyName As String
    Dim PropertyValue As Range
    Rem otherwise would throw error if control doesn
    On Error Resume Next
    Rem uForm.Controls
    For Each controlItem In uForm.Designer.Controls
        UFSheet.Cells(1, ColumnNumber) = TypeName(controlItem)
        For counter = 2 To 79
            PropertyName = UFSheet.Cells(counter, 6).Value
            Set PropertyValue = UFSheet.Cells(counter, ColumnNumber)
            PropertyValue.Value = CallByName(controlItem, PropertyName, VbGet)
            Rem mark if control is inside frame
            If UCase(PropertyName) = UCase("NAME") Then
                If controlItem.parent.Name <> uForm.Name Then
                    PropertyValue.Value = PropertyValue.Value & "_" & controlItem.parent.Name
                End If
            End If
        Next counter
        ColumnNumber = ColumnNumber + 1
    Next
    Rem export userform properties
    ColumnNumber = 2
    UFSheet.Cells(1, ColumnNumber) = uForm.Name
    UFSheet.Cells(2, ColumnNumber) = uForm.Name & "RC"
    RowNumber = 2
    Dim rng As Range
    Set rng = UFSheet.Range("A3:A" & UFSheet.Range("A" & rows.count).End(xlUp).row)
    For Each cell In rng
        cell.OFFSET(0, 1).Value = uForm.Properties(cell.Value)
    Next
    Rem export code
    ColumnNumber = 4
    UFSheet.Cells(1, 4) = "CODE"
    Dim Code As Variant
    Code = GetModuleText(WorkbookOfModule(uForm).VBProject.VBComponents(uForm.Name))
    Code = Split(Code, vbNewLine)
    UFSheet.Range("D2").RESIZE(UBound(Code) + 1).Value = WorksheetFunction.Transpose(Code)
    UFSheet.Cells.Range("G1:XFD1").EntireColumn.AutoFit
    UFSheet.Cells.NumberFormat = "General"
    With UFSheet.rows(1)
        .Font.Bold = True
        .Font.Size = 14
    End With
    Rem Sort-group controls and put frames first as they need to be created first to hold their controls
    Set rng = UFSheet.Range("F1").CurrentRegion.OFFSET(0, 1)
    Set rng = rng.RESIZE(, rng.Columns.count - 1)
    SortColumns rng
    PutFramesFirst
    Application.ScreenUpdating = True
End Sub

Sub PutFramesFirst()
    '#INCLUDE dp
    Dim ws As Worksheet
    Set ws = SHEETS("FormBuilder")
    Dim rng As Range
    Set rng = ws.Range("F1").CurrentRegion.OFFSET(0, 1).RESIZE(1)
    Set rng = rng.RESIZE(, rng.Columns.count - 1)
    dp rng.Address
    Dim cell As Range
    For Each cell In rng
        If cell Like "Frame" And cell.Column <> Columns("G").Column Then
            cell.EntireColumn.Cut
            ws.Columns("G").Insert
        End If
    Next
    For Each cell In rng
        If cell = "Frame" Then
            If InStr(1, cell.OFFSET(1), "_") = 0 Then
                cell.EntireColumn.Cut
                ws.Columns("G").Insert
            End If
        End If
    Next
End Sub

Public Sub CreateUserForm(Optional TargetWorkbook As Workbook)
    '#INCLUDE ArrayToString
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveWorkbook
    Set UFSheet = ThisWorkbook.SHEETS("FormBuilder")
    On Error Resume Next
    Dim ctr As MSForms.control
    Dim propertyCOUNTER As Long
    Dim counter As Long
    Dim cell As Range
    Dim rng As Range
    Rem ----create userform-------------------------
    Dim uf As VBComponent
    Set uf = TargetWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    uf.Name = UFSheet.Range("B2")
    Set rng = UFSheet.Range("A3:A" & UFSheet.Range("A" & rows.count).End(xlUp).row)
    For Each cell In rng
        uf.Properties(cell.Value) = cell.OFFSET(0, 1).Value
    Next
    Rem ---add controls to userform-------------------
    Dim ControlName As Range
    Dim FrameName As String
    Set rng = UFSheet.Range("G1", UFSheet.Cells(1, Columns.count).End(xlToLeft))
    For Each cell In rng
        Set ControlName = cell.OFFSET(1)
        If InStr(1, ControlName.Value, "_") > 0 Then
            FrameName = Split(ControlName.Value, "_")(1)
            Set ctr = uf.Designer.Controls(FrameName).Add("forms." & cell & ".1")
            CallByName ctr, Split(ControlName.Value, "_")(0), VbLet, UFSheet.Cells(counter, cell.Column).Value
        Else
            Set ctr = uf.Designer.Controls.Add("forms." & cell & ".1")
            CallByName ctr, ControlName.Value, VbLet, UFSheet.Cells(counter, cell.Column).Value
        End If
        With ctr
            For counter = 3 To 71
                CallByName ctr, UFSheet.Cells(counter, 6).Value, VbLet, UFSheet.Cells(counter, cell.Column).Value
            Next counter
        End With
    Next cell
    Rem ---add code to userform--------------------------
    With UFSheet.Range("D1")
        .Value = "CODE"
        .Font.Bold = True
    End With
    Dim ImportCode As String
    ImportCode = ArrayToString(UFSheet.Range("D2:D" & UFSheet.Range("D" & rows.count).End(xlUp).row).Value, vbNewLine)
    uf.CodeModule.AddFromString (ImportCode)
End Sub

