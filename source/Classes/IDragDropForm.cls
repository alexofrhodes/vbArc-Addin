VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IDragDropForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Set DragDropForm(ByRef oFrm As Object)
End Property

Public Property Let DropEffect(ByVal iVal As DropEffects)
End Property

Public Property Let DropType(ByVal iVal As DropTypes)
End Property

Public Property Let AllowDropInSameListbox(ByVal bValue As Boolean)
End Property

Public Sub GetListboxes(ParamArray aListBoxGroupNames() As Variant)
End Sub

Public Sub TerminateDragDrop()
End Sub

