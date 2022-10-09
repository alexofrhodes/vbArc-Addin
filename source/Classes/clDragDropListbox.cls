VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDragDropListbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'--------------------------------------------------------------------------------------------------------------------------
' Author    : Manfred van den Noort
' Copyright : © 2020 worksheetsvba.com, all rights reserved
' Date      : 2020-12-07
' Version   : 1.0
' Purpose   : Enable ListBox Drag & Drop Advanced Version
'
' You are free to use, change and distribute the code in this class module for your own projects,
'       however this header with copyright notice should remain intact
' You may not publish the code in this class module, for example on a web site, without the explicit consent of the author.
'--------------------------------------------------------------------------------------------------------------------------

Private oParentForm As clDragDropForm
Private WithEvents LstBx As MSForms.ListBox
Attribute LstBx.VB_VarHelpID = -1
Private sngItemHeight As Single
Private lNewListIdx As Long
Private sGroupName As String
Private bDeselectValues As Boolean

Private Sub Class_Initialize()
    lNewListIdx = -1
End Sub

Friend Property Set ParentForm(ByRef oValue As clDragDropForm)
    Set oParentForm = oValue
End Property

Friend Property Set DragDropListBox(ByRef oValue As MSForms.ListBox)
    Set LstBx = oValue
End Property

Friend Property Let GroupName(sValue As String)
    sGroupName = sValue
End Property

Private Sub LstBx_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim oDataObject As DataObject, sItemValue As String, i As Long, j As Long, k As Long, aSelectedItems As Variant, lEffect As Long, lItemClicked As Long, sColValues As String
    If Button = 1 Then
        If LstBx.multiSelect = fmMultiSelectSingle Then
            If LstBx.ListIndex > -1 Then
                For i = 1 To LstBx.columnCount
                    sItemValue = sItemValue & "~" & LstBx.list(LstBx.ListIndex, i - 1)
                Next
                sItemValue = Mid(sItemValue, 2)
            End If
        Else
            lItemClicked = GetMouseOverRow(Y)
            For i = 0 To LstBx.ListCount - 1
                If LstBx.SELECTED(i) Or i = lItemClicked Then
                    LstBx.SELECTED(i) = True
                    sColValues = ""
                    For k = 1 To LstBx.columnCount
                        sColValues = sColValues & "~" & LstBx.list(i, k - 1)
                    Next
                    sItemValue = sItemValue & "~~" & Mid(sColValues, 2)
                    If j = 0 Then
                        ReDim aSelectedItems(0)
                    Else
                        ReDim Preserve aSelectedItems(j)
                    End If
                    aSelectedItems(j) = i
                    j = j + 1
                End If
            Next
            sItemValue = Mid(sItemValue, 3)
        End If
        If sItemValue <> "" Then
            If LstBx.multiSelect = fmMultiSelectSingle Then
                oParentForm.SelectedDragItems = Array(LstBx.ListIndex)
            Else
                oParentForm.SelectedDragItems = aSelectedItems
            End If
            oParentForm.DragSourceGroupName = sGroupName
            Set oParentForm.DragSource = LstBx
            Set oDataObject = New DataObject
            oDataObject.SetText sItemValue
            lEffect = oDataObject.StartDrag
        End If
    Else
        'set ListIndex on new added item in case of fmMultiSelectSingle in same listbox
        If lNewListIdx >= 0 Then
            LstBx.ListIndex = lNewListIdx
            lNewListIdx = -1
        End If
        'reset multiselect dragsource in case of MoveItem
        If IsObject(oParentForm.DragSource) And bDeselectValues Then
            bDeselectValues = False
            If oParentForm.DragSource.multiSelect <> fmMultiSelectSingle Then
                For i = 0 To oParentForm.DragSource.ListCount - 1
                    oParentForm.DragSource.SELECTED(i) = False
                Next
            End If
        End If
    End If
End Sub

Private Sub LstBx_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As Long, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Dim i As Long, iActiveItemRow As Long
    Cancel = True
    If (oParentForm.DragSource Is LstBx And (Not oParentForm.AllowDropInSameListbox Or oParentForm.DropType <> CursorPosition Or LstBx.multiSelect <> fmMultiSelectSingle)) _
        Or LstBx.Locked Or sGroupName <> oParentForm.DragSourceGroupName Or oParentForm.DragSource.columnCount <> LstBx.columnCount Then
        Effect = fmDropEffectNone
    Else
        Effect = oParentForm.DropEffect
        If oParentForm.DropType = CursorPosition Then
            With LstBx
                iActiveItemRow = GetMouseOverRow(Y)
                If oParentForm.AllowDropInSameListbox Then
                    If iActiveItemRow = LstBx.ListCount Then
                        oParentForm.DragSource.SetFocus
                    Else
                        LstBx.SetFocus
                    End If
                End If
                If DragState = fmDragStateLeave Then
                    If oParentForm.AllowDropInSameListbox And LstBx Is oParentForm.DragSource Then
                        'highlight selected items
                        If LstBx.multiSelect = fmMultiSelectSingle Then
                            LstBx.ListIndex = oParentForm.SelectedDragItems(0)
                        Else
                            For i = 0 To LstBx.ListCount - 1
                                LstBx.SELECTED(i) = False
                            Next
                            For i = 0 To UBound(oParentForm.SelectedDragItems)
                                LstBx.SELECTED(oParentForm.SelectedDragItems(i)) = True
                            Next
                        End If
                    Else
                        For i = 0 To .ListCount - 1
                            .SELECTED(i) = False
                        Next
                    End If
                Else
                    'highlight items with mousemove
                    For i = 0 To .ListCount - 1
                        If i = iActiveItemRow Then
                            .SELECTED(i) = True
                        Else
                            .SELECTED(i) = False
                        End If
                    Next
                End If
            End With
        End If
    End If
End Sub

Private Sub LstBx_BeforeDropOrPaste(ByVal Cancel As MSForms.ReturnBoolean, ByVal action As Long, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    Dim i As Long, j As Long, iInsertPosition As Long, aItemValues() As String, aColValues() As String
    Cancel = True
    Effect = oParentForm.DropEffect
    aItemValues = Split(Data.GetText, "~~")
    For i = 0 To UBound(aItemValues)
        aColValues = Split(aItemValues(i), "~")
        With LstBx
            Select Case oParentForm.DropType
                Case Alphabetic
                    For j = 0 To .ListCount - 1
                        If .list(j) > aColValues(0) Then
                            Exit For
                        End If
                    Next
                    iInsertPosition = j
                Case BottomOfList
                    iInsertPosition = .ListCount
                Case CursorPosition
                    If i = 0 Then
                        iInsertPosition = GetMouseOverRow(Y)
                    Else
                        iInsertPosition = iInsertPosition + 1
                    End If
                Case Numeric
                    For j = 0 To .ListCount - 1
                        If val(.list(j)) > val(aColValues(0)) Then
                            Exit For
                        End If
                    Next
                    iInsertPosition = j
            End Select
            .AddItem aColValues(0), iInsertPosition
            For j = 1 To UBound(aColValues)
                .list(iInsertPosition, j) = aColValues(j)
            Next
            If oParentForm.DragSource Is LstBx Then
                For j = 0 To UBound(oParentForm.SelectedDragItems)
                    If iInsertPosition < oParentForm.SelectedDragItems(j) Then
                        oParentForm.IncreaseSelectedItemIndex j
                    End If
                Next
            End If
            .ListIndex = iInsertPosition
            lNewListIdx = iInsertPosition
        End With
    Next
    If oParentForm.DropEffect = MoveItem Then
        With oParentForm.DragSource
            For i = UBound(oParentForm.SelectedDragItems) To 0 Step -1
                .RemoveItem oParentForm.SelectedDragItems(i)
                If oParentForm.SelectedDragItems(i) < lNewListIdx Then
                    lNewListIdx = lNewListIdx - 1
                End If
            Next
            If Not oParentForm.DragSource Is LstBx Then
                'reset lNewListIdx because this is only used for drop in same listbox (because then after drop the listindex is reset to wrong value)
                lNewListIdx = -1
                If oParentForm.DragSource.multiSelect = fmMultiSelectSingle Then
                    .ListIndex = -1
                Else
                    bDeselectValues = True
                End If
            End If
        End With
        LstBx.SetFocus
    End If
    Erase oParentForm.SelectedDragItems
End Sub

Private Function GetMouseOverRow(ByVal Ypos As Single) As Long
    Dim iItemRow As Long
    With LstBx
        If sngItemHeight > 0 Then
            iItemRow = .TopIndex + Ypos / sngItemHeight
        End If
        If iItemRow > .ListCount - 1 Then
            iItemRow = .ListCount
        End If
    End With
    GetMouseOverRow = iItemRow
End Function

Friend Sub SetItemRowHeight()
    'get ItemHeight by setting the TopIndex to Listount - 1
    'if TopIndex > 0 then all rows in the listbox are filled
    'if the listbox is not completely filled, temporary items are added until TopIndex > 0
    Dim lOldTopIndex As Long, lTempTopIndex As Long, i As Long, j As Long
    With LstBx
        If .ListCount > 0 Then
            lOldTopIndex = .TopIndex
            .TopIndex = .ListCount - 1
            lTempTopIndex = .TopIndex
        End If
        Do While lTempTopIndex = 0
            i = i + 1
            .AddItem i
            .TopIndex = .ListCount - 1
            lTempTopIndex = .TopIndex
        Loop
        sngItemHeight = .Height / (.ListCount - .TopIndex)
        For j = 1 To i
            .RemoveItem .ListCount - 1
        Next
        If .ListCount > 0 Then
            .TopIndex = lOldTopIndex
        End If
    End With
End Sub


