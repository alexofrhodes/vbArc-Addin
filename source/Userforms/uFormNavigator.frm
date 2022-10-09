VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFormNavigator 
   Caption         =   "double click to open a form"
   ClientHeight    =   4440
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2916
   OleObjectBlob   =   "uFormNavigator.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFormNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uFormNavigator
'* Created    : 06-10-2022 10:35
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub CommandButton1_Click()
    ListBox1.ListIndex = -1
End Sub

Private Sub Image2_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub oActive_Click()
    LoadForms
End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim FormName As String
    FormName = ListBox1.list(ListBox1.ListIndex)
    ShowUserform FormName
End Sub

Private Sub UserForm_Initialize()
    LoadForms
End Sub

Sub LoadForms()
    ListBox1.clear
    Dim wb As Workbook
    '    If oActive.Value = True Then
    Set wb = ActiveWorkbook
    '    Else
    '        Set wb = ThisWorkbook
    '    End If
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_MSForm Then
            ListBox1.AddItem vbComp.Name
        End If
    Next
End Sub

