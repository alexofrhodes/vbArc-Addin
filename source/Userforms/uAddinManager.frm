VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAddinManager 
   Caption         =   "ADDINS MANAGER"
   ClientHeight    =   7584
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4896
   OleObjectBlob   =   "uAddinManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uAddinManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uAddinManager
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub CommandButton1_Click()
    On Error Resume Next
    For i = 0 To body.ListCount - 1
        If body.SELECTED(i) = True Then
            If Not AddIns(body.list(i, 0)).IsOpen Then
                Workbooks.Open (AddIns(body.list(i, 0)).FullName)
                AddIns(body.list(i, 0)).Installed = True
            Else
                Workbooks(AddIns(body.list(i, 0)).Name).Close True
                AddIns(body.list(i, 0)).Installed = False
            End If
        End If
        '    If body.Selected(i) = True Then AddIns(body.List(i, 0)).Installed = Not AddIns(body.List(i, 0)).Installed
    Next
    LoadAddins
End Sub

Private Sub CommandButton2_Click()
    Dim ans As Long
    ans = MsgBox("Irreversible. Proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    AddIns(body.list(body.ListIndex, 0)).Installed = False
    Kill AddIns(body.list(body.ListIndex, 0)).FullName
End Sub

Private Sub CommandButton3_Click()
    SortListboxOnColumn body, 0
End Sub

Private Sub CommandButton4_Click()
    SortListboxOnColumn body, 1
End Sub

Private Sub CommandButton5_Click()
    vbArcAddinsForm.Show
    Unload Me
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Label1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    uUpdateFiles.Show
End Sub

Private Sub UserForm_Initialize()
    LoadAddins
End Sub

Sub LoadAddins()
    '#INCLUDE SortListboxOnColumn
    body.clear
    Dim ad As AddIn
    On Error Resume Next
    For Each ad In AddIns
        body.AddItem
        body.list(body.ListCount - 1, 0) = left(ad.Name, InStr(1, ad.Name, ".") - 1)
        body.list(body.ListCount - 1, 1) = IIf(ad.Installed, " ENABLED", "-")
    Next
    SortListboxOnColumn body, 1
End Sub

Sub SortListboxOnColumn(lBox As MSForms.ListBox, Optional OnColumn As Long = 0)
    Dim vntData As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    vntData = lBox.list
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                For lngSubItemIndex = 0 To lBox.columnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    lBox.clear
    lBox.list = vntData
End Sub

