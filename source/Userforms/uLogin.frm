VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uLogin 
   Caption         =   "Login Form"
   ClientHeight    =   4680
   ClientLeft      =   48
   ClientTop       =   396
   ClientWidth     =   5604
   OleObjectBlob   =   "uLogin.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Label1_Click()
    Login
End Sub

Sub Login()
    '#INCLUDE CheckUser
    ThisWorkbook.SHEETS("uLoginSettings").Range("B6").Value = Me.Password.Value
    ThisWorkbook.SHEETS("uLoginSettings").Range("B5").Value = Me.UserNames.Value
    CheckUser
End Sub

Private Sub Label3_Click()
    Unload Me
End Sub

Private Sub UserForm_Initialize()
    MakeFormBorderless Me
    MakeFormTransparent Me, vbYellow
    ListUsernames
End Sub

Sub ListUsernames()
    UserNames.clear
    Dim rng As Range
    Set rng = ThisWorkbook.SHEETS("uLoginSettings").Range("E4").CurrentRegion.RESIZE(, 1)
    Set rng = rng.OFFSET(1).RESIZE(rng.rows.count - 1)
    Dim cell As Range
    For Each cell In rng
        UserNames.AddItem cell
    Next
    UserNames.ListIndex = 0
End Sub

Private Sub UserNames_DropButtonClick()
    UserNames.ForeColor = vbBlack
End Sub

Private Sub UserNames_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    UserNames.ForeColor = vbWhite
End Sub

