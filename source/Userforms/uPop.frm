VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uPop 
   Caption         =   "UserForm1"
   ClientHeight    =   5244
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   8028
   OleObjectBlob   =   "uPop.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uPop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uPop
'* Created    : 06-10-2022 10:39
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Rem call this form with the function    -   uMSG()
Private Sub UserForm_Initialize()
    MakeFormBorderless Me
    MakeFormTransparent Me
    UserformOnTop Me
End Sub

Sub Init(Optional Caption As Variant = "vbArc", _
         Optional SecondsPerMessage As Long = 3, _
         Optional ImagePath As String, _
         Optional TextSize As Long = 12, _
         Optional FontBold As Boolean = True, _
         Optional TextColor As Long = vbBlack, _
         Optional CounterColor As Long = vbBlack)
    '#INCLUDE CountDown
    SecondsPerMessage = SecondsPerMessage * 100
    Label3.Font.Size = TextSize
    Label3.ForeColor = TextColor
    Label3.Font.Bold = FontBold
    Label2.ForeColor = CounterColor
    If ImagePath <> "" Then
        Image1.Picture = LoadPicture(ImagePath)
        Image1.PictureSizeMode = fmPictureSizeModeStretch
    End If
    Me.Show
    On Error GoTo LoopEnd
    Application.EnableCancelKey = xlErrorHandler
    Dim element As Variant
    If TypeName(Caption) = "String" Then
        Label3.Caption = Caption
        CountDown SecondsPerMessage
    Else
        For Each element In Caption
            Label3.Caption = element
            CountDown SecondsPerMessage
        Next
    End If
LoopEnd:
    Application.EnableCancelKey = xlInterrupt
    Unload Me
End Sub

Sub CountDown(PopSleep As Long)
    '#INCLUDE Pop
    Dim i As Long
    i = 0
    Do While i < PopSleep / 100
        Label2.Caption = Round(PopSleep / 100, 0) - i
        DoEvents
        Sleep 1000
        i = i + 1
    Loop
End Sub


