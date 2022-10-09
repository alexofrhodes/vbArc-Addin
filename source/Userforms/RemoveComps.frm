VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RemoveComps 
   Caption         =   "Remove Code or Components"
   ClientHeight    =   5736
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7572
   OleObjectBlob   =   "RemoveComps.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RemoveComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : RemoveComps
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Remover_Click()
    If LComponents.ListCount = 0 Then Exit Sub
    If pmWorkbook Is Nothing Then Set pmWorkbook = ActiveWorkbook
    Dim i As Long
    For i = 0 To LComponents.ListCount - 1
        If LComponents.SELECTED(i) Then
            If oCode.Value = True Then
                ClearComponent pmWorkbook.VBProject.VBComponents(LComponents.list(i, 1))
            ElseIf oComps.Value = True Then
                DeleteComponent pmWorkbook.VBProject.VBComponents(LComponents.list(i, 1))
            End If
        End If
    Next i
    addCompsList
End Sub

Private Sub UserForm_Initialize()
    If pmWorkbook Is Nothing Then Set pmWorkbook = ActiveWorkbook
    addCompsList
    Me.Caption = "Comps of " & pmWorkbook.Name
End Sub

Sub addCompsList()
    '#INCLUDE SortListboxOnColumn
    '#INCLUDE ComponentTypeToString
    '#INCLUDE GetSheetByCodeName
    '#INCLUDE ResizeUserformToFitControls
    '#INCLUDE ResizeControlColumns
    LComponents.clear
    Dim vbComp As VBComponent
    For Each vbComp In pmWorkbook.VBProject.VBComponents
        If vbComp.Name <> "ThisWorkbook" Then
            LComponents.AddItem
            LComponents.list(LComponents.ListCount - 1, 0) = ComponentTypeToString(vbComp.Type)
            LComponents.list(LComponents.ListCount - 1, 1) = vbComp.Name
            If vbComp.Type = vbext_ct_Document Then
                LComponents.list(LComponents.ListCount - 1, 2) = GetSheetByCodeName(pmWorkbook, vbComp.Name).Name
            End If
        End If
    Next
    SortListboxOnColumn LComponents, 0
    ResizeControlColumns LComponents
    ResizeUserformToFitControls Me
    Me.Repaint
End Sub

