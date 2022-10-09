VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} RenameComps 
   Caption         =   "Rename Components"
   ClientHeight    =   7104
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   5700
   OleObjectBlob   =   "RenameComps.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "RenameComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : RenameComps
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

Private Sub RenameComponents_Click()
    If pmWorkbook Is Nothing Then Set pmWorkbook = ActiveWorkbook
    Dim NewNames As Variant
    Dim i As Long
    NewNames = Split(textboxNewName, vbNewLine)
    For i = 0 To UBound(NewNames)
        If NewNames(i) = vbNullString Then
            NewNames(i) = LRenameListbox.list(i)
        End If
    Next i
    For i = 0 To UBound(NewNames)
continue:
        On Error GoTo eh
        Select Case LRenameListbox.list(i, 0)
            Case Is = "Module", "Class", "UserForm"
                If LRenameListbox.list(i, 1) <> NewNames(i) Then
                    pmWorkbook.VBProject.VBComponents(LRenameListbox.list(i, 1)).Name = NewNames(i)
                End If
            Case Is = "Document"
                If LRenameListbox.list(i, 1) <> NewNames(i) Then
                    pmWorkbook.SHEETS(LRenameListbox.list(i, 1)).Name = NewNames(i)
                End If
        End Select
    Next
    For i = 0 To LRenameListbox.ListCount - 1
        LRenameListbox.list(i, 1) = NewNames(i)
    Next i
    textboxNewName.TEXT = vbNullString
    Dim str As String
    str = Join(NewNames, vbNewLine)
    textboxNewName.TEXT = str
    MsgBox "Components renamed"
    Exit Sub
eh:
    NewNames(i) = NewNames(i) & i + 1
    Resume continue
End Sub

Private Sub UserForm_Initialize()
    If pmWorkbook Is Nothing Then Set pmWorkbook = ActiveWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In pmWorkbook.VBProject.VBComponents
        If vbComp.Name <> "ThisWorkbook" Then
            LRenameListbox.AddItem
            LRenameListbox.list(LRenameListbox.ListCount - 1, 0) = ComponentTypeToString(vbComp.Type)
            If vbComp.Type <> vbext_ct_Document Then
                LRenameListbox.list(LRenameListbox.ListCount - 1, 1) = vbComp.Name
            Else
                LRenameListbox.list(LRenameListbox.ListCount - 1, 1) = GetSheetByCodeName(pmWorkbook, vbComp.Name).Name
            End If
        End If
    Next
    SortListboxOnColumn LRenameListbox, 0
    Dim str As String
    str = LRenameListbox.list(0, 1)
    For i = 1 To LRenameListbox.ListCount - 1
        str = str & vbNewLine & LRenameListbox.list(i, 1)
    Next
    textboxNewName.TEXT = str
End Sub

