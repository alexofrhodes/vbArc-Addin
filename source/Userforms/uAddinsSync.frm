VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAddinsSync 
   Caption         =   "REPLACE OLD FILE WITH NEW VERSION"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "uAddinsSync.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uAddinsSync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uAddinsSync
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub CommandButton1_Click()
    UpdateFiles
End Sub

Private Sub CommandButton2_Click()
    Dim c As MSForms.control
    For Each c In Me.Controls
        If UCase(TypeName(c)) = UCase("CheckBox") Then
            c.Value = True
        End If
    Next
End Sub

Private Sub CommandButton3_Click()
    Dim c As MSForms.control
    For Each c In Me.Controls
        If UCase(TypeName(c)) = UCase("CheckBox") Then
            c.Value = False
        End If
    Next
End Sub

Private Sub UserForm_Activate()
    ResizeUserformToFitControls Me
End Sub

Private Sub UserForm_Initialize()
    AddinsModified
End Sub

