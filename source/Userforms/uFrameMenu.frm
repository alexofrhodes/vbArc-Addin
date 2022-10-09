VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFrameMenu 
   ClientHeight    =   5364
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9432.001
   OleObjectBlob   =   "uFrameMenu.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFrameMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uFrameMenu
'* Created    : 06-10-2022 10:36
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


Rem See the video of at https://youtu.be/l9b6DvCig5E


Rem @TODO create from sheet data

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Private Sub Emitter_LabelMouseOut(label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then
        If label.BackColor <> &H80B91E Then label.BackColor = &H534848
    End If
End Sub

Private Sub Emitter_LabelMouseOver(label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then
        If label.BackColor <> &H80B91E Then label.BackColor = &H808080
    End If
End Sub

Sub Emitter_LabelClick(ByRef label As MSForms.label)
    If InStr(1, label.Tag, "reframe", vbTextCompare) > 0 Then Reframe Me, label
End Sub

Private Sub UserForm_Initialize()
    Dim anc As MSForms.control

    For Each c In Me.Controls
        If TypeName(c) = "Frame" Then
            'c.Caption = ""
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.visible = False
                If InStr(1, c.Tag, "anchor") > 0 Then
                    On Error Resume Next
                    Set anc = Me.Controls("Anchor" & Mid(c.Tag, InStr(1, c.Tag, "Anchor", vbTextCompare) + Len("Anchor"), 2))
                    If anc Is Nothing Then Stop
                    On Error GoTo 0
                    c.top = anc.top        'Anchor01.Top
                    c.left = anc.left        ' Anchor01.Left
                    Set anc = Nothing
                End If
            End If
        End If
    Next
    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
End Sub

