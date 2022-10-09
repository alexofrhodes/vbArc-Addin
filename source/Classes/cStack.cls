VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module cStack
'Uses a VBA Collection as the stack data structure

Option Explicit

Dim pStack As Collection

Private Sub Class_Initialize()
    Set pStack = New Collection
End Sub

Private Sub Class_Terminate()
    Set pStack = Nothing
End Sub

Public Function Push(newItem As cStackItem) As cStackItem
    With pStack
        .Add newItem
        Set Push = .item(.count)
    End With
End Function

Public Function Pop() As cStackItem
    With pStack
        If .count > 0 Then
            Set Pop = .item(.count)
            .Remove .count
        End If
    End With
End Function

Public Function count() As Long
    count = pStack.count
End Function

Public Sub Dump()
    Dim i As Long
    Dim stackItem As cStackItem
    With pStack
        If .count = 0 Then Debug.Print "Empty"
        For i = .count To 1 Step -1
            Set stackItem = .item(i)
            With stackItem
                Debug.Print i; .baseCell.OFFSET(.rowOffset).Address(False, False), Hex(.hwndParent), "Offset " & .rowOffset
            End With
        Next
    End With
End Sub

