Attribute VB_Name = "F_ListView"
Rem @Folder ListView Declarations
Public Declare PtrSafe Function SendMessage Lib "user32" Alias "SendMessageA" ( _
ByVal hWnd As Long, _
ByVal wMsg As Long, _
ByVal wParam As Long, _
lParam As Any) As Long
Public Const LVM_FIRST = &H1000
Public Const LVM_SETCOLUMNWIDTH = (LVM_FIRST + 30)
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2

Rem @Folder ListView
Public Sub ListviewAutoSizeColumns(LV As ListView, Optional Column As ColumnHeader = Nothing)
    Dim c As ColumnHeader
    If Column Is Nothing Then
        For Each c In LV.ColumnHeaders
            SendMessage LV.hWnd, LVM_FIRST + 30, c.index - 1, ByVal -2
        Next
    Else
        SendMessage LV.hWnd, LVM_FIRST + 30, Column.index - 1, ByVal -2
    End If
    LV.Refresh
End Sub

Sub ListViewEdit(LV As ListView, RowIndex As Long, ColumnIndex As Long, NewValue As Variant)
    Rem base 1 like range
    If ColumnIndex = 1 Then
        LV.ListItems(RowIndex).TEXT = NewValue
    ElseIf ColumnIndex > 1 Then
        LV.ListItems(RowIndex).ListSubItems(ColumnIndex - 1).TEXT = NewValue
    End If
End Sub

Function ListViewSelected(LV As ListView, Optional ColumnIndex As Long = 0, Optional delimeter As String = ",") As Variant
    Rem base 1 like range
retry:
    If ColumnIndex = 0 Then
        Dim s As String
        s = LV.ListItems(LV.SelectedItem.index)
        Dim counter
        For counter = 1 To LV.ColumnHeaders.count - 1
            s = s & delimeter & LV.ListItems(LV.SelectedItem.index).ListSubItems(counter)
        Next
        ListViewSelected = Split(s, delimeter)
    ElseIf ColumnIndex = 1 Then
        ListViewSelected = LV.ListItems(LV.SelectedItem.index)
    ElseIf ColumnIndex > 1 Then
        ListViewSelected = LV.ListItems(ColumnIndex).ListSubItems(ColumnIndex - 1)
    Else
        ColumnIndex = -1
        GoTo retry
    End If
End Function

Sub ListViewPopulateFromArray(LV As ListView, inputArray As Variant)
    Dim vListItem As listItem
    Dim vChildItem As ListSubItem
    Dim vHeader As Variant
    Dim iRows As Long, iColumns As Long
    For iColumns = LBound(inputArray, 2) To UBound(inputArray, 2)
        Set vHeader = LV.ColumnHeaders.Add(, , inputArray(LBound(inputArray, 1), iColumns))
    Next
    For iRows = LBound(inputArray, 1) + 1 To UBound(inputArray, 1)
        Set vListItem = LV.ListItems.Add(, , inputArray(iRows, 1))
        For iColumns = LBound(inputArray, 2) + 1 To UBound(inputArray, 2)
            Set vChildItem = vListItem.ListSubItems.Add(, , inputArray(iRows, iColumns))
        Next
    Next
    LV.View = lvwReport
End Sub

Sub ListViewClear(LV As ListView)
    Dim i As Long
    For i = LV.ListItems.count To 1 Step -1
        LV.ListItems.Remove i
    Next
    LV.Refresh
End Sub


