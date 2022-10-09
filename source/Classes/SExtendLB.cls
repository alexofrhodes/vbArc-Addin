VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SExtendLB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Rem in userform
Rem'''''''''''''''''''''''''''''''
Rem Dim LBCollection As SECollect
Rem
Rem Private Sub UserForm_Initialize()
Rem Set LBCollection = New SECollect
Rem LBCollection.Add ListBox1
Rem LBCollection.Add ListBox2
Rem End Sub

Option Explicit
Private Declare PtrSafe Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Public WithEvents LB As MSForms.ListBox
Attribute LB.VB_VarHelpID = -1

Private ItemsSelected As Variant

Private Sub LB_Change()
    Dim i As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With LB
            For i = 0 To .ListCount - 1
                If ItemsSelected(i) = False And .SELECTED(i) = True Then
                    ItemsSelected(i) = True
                End If
            Next i
        End With
    End If
End Sub

Private Sub LB_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With LB
            ReDim ItemsSelected(0 To .ListCount - 1) As Boolean
            For i = 0 To .ListCount - 1
                ItemsSelected(i) = .SELECTED(i)
            Next i
        End With
    End If
End Sub

Private Sub LB_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    If CBool(GetKeyState(&HA0) And &HFF80) Then
        With LB
            For i = 0 To .ListCount - 1
                .SELECTED(i) = ItemsSelected(i)
            Next i
        End With
    End If
End Sub

Public Sub Add(ByVal NewLB As MSForms.ListBox)
    Set LB = NewLB
End Sub

