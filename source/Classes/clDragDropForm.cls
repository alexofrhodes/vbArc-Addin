VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clDragDropForm"
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

Implements IDragDropForm

Private frmDragDrop As MSForms.UserForm
Private lstbxDragSource As MSForms.ListBox
Private colListBoxes As Collection
Private iDropType As DropTypes, iDropEffect As DropEffects
Private bAllowDropInSameListbox As Boolean
Private aSelectedDragItems As Variant
Private sDragSourceGroupName As String

Public Enum DropEffects
    CopyItem = fmDropEffectCopy
    MoveItem = fmDropEffectMove
End Enum

Public Enum DropTypes
    Alphabetic = 1
    BottomOfList = 2
    CursorPosition = 3
    Numeric = 4
End Enum

Private Sub Class_Initialize()
    Set colListBoxes = New Collection
    'set default values
    iDropEffect = MoveItem
    iDropType = BottomOfList
    bAllowDropInSameListbox = False
End Sub

Private Property Set IDragDropForm_DragDropForm(ByRef oFrm As Object)
    Set frmDragDrop = oFrm
End Property

Private Property Let IDragDropForm_DropEffect(ByVal iVal As DropEffects)
    iDropEffect = iVal
End Property

Friend Property Get DropEffect() As DropEffects
    DropEffect = iDropEffect
End Property

Private Property Let IDragDropForm_DropType(ByVal iVal As DropTypes)
    iDropType = iVal
End Property

Friend Property Get DropType() As DropTypes
    DropType = iDropType
End Property

Private Property Let IDragDropForm_AllowDropInSameListbox(ByVal bValue As Boolean)
    bAllowDropInSameListbox = bValue
End Property

Friend Property Get AllowDropInSameListbox() As Boolean
    AllowDropInSameListbox = bAllowDropInSameListbox
End Property

Friend Property Let SelectedDragItems(ByVal aValue As Variant)
    aSelectedDragItems = aValue
End Property

Friend Property Get SelectedDragItems() As Variant
    SelectedDragItems = aSelectedDragItems
End Property

Friend Property Set DragSource(ByRef oListBox As MSForms.ListBox)
    Set lstbxDragSource = oListBox
End Property

Friend Property Get DragSource()
    Set DragSource = lstbxDragSource
End Property

Friend Property Let DragSourceGroupName(sValue As String)
    sDragSourceGroupName = sValue
End Property

Friend Property Get DragSourceGroupName() As String
    DragSourceGroupName = sDragSourceGroupName
End Property

Private Sub IDragDropForm_GetListboxes(ParamArray aListBoxGroupNames() As Variant)
    Dim ctl As control, oDragDropListBox As clDragDropListbox, iTabIndex As Integer, sActiveListbox As String, i As Integer
    iTabIndex = (2 ^ 15) - 1
    For Each ctl In frmDragDrop.Controls
        If TypeName(ctl) = "ListBox" Then
            If ctl.Enabled And Not ctl.Locked Then
                'SetFocus is used, otherwise the first Y in the MouseMove event is not always correct
                ctl.SetFocus
                If ctl.TabIndex < iTabIndex Then
                    iTabIndex = ctl.TabIndex
                    sActiveListbox = ctl.Name
                End If
            End If
            Set oDragDropListBox = New clDragDropListbox
            With oDragDropListBox
                Set .ParentForm = Me
                Set .DragDropListBox = ctl
                .SetItemRowHeight
                For i = 0 To UBound(aListBoxGroupNames)
                    If InStr(aListBoxGroupNames(i), "~") > 0 Then
                        If Split(aListBoxGroupNames(i), "~")(0) = ctl.Name Then
                            .GroupName = Split(aListBoxGroupNames(i), "~")(1)
                        End If
                    End If
                Next
            End With
            colListBoxes.Add oDragDropListBox
        End If
    Next
    If sActiveListbox <> "" Then
        frmDragDrop.Controls(sActiveListbox).SetFocus
    End If
End Sub

Friend Sub IncreaseSelectedItemIndex(iIdx As Long)
    aSelectedDragItems(iIdx) = aSelectedDragItems(iIdx) + 1
End Sub

Private Sub IDragDropForm_TerminateDragDrop()
    Set colListBoxes = Nothing
End Sub

