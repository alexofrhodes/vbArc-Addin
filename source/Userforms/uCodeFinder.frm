VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uCodeFinder 
   Caption         =   "UserForm1"
   ClientHeight    =   8964.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4212
   OleObjectBlob   =   "uCodeFinder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uCodeFinder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uCodeFinder
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim CalledFromModule As VBComponent
Dim CalledFromProcedure As String

Private Sub CommandButton2_Click()
    ReturnToCaller
End Sub

Private Sub UserForm_Initialize()
    '#INCLUDE ImageListLoadProjectIcons
    MakeUserFormChildOfVBEditor uCodeFinder.Caption

    With TreeView1
        .Sorted = True
        .Appearance = ccFlat
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Font.Size = 10
        .Indentation = 2
    End With

    ImageListLoadProjectIcons ImageList1, TreeView1
    
    Set CalledFromModule = ActiveModule
    CalledFromProcedure = ActiveProcedure
    
End Sub

Sub ReturnToCaller()
    On Error GoTo HELL
    GoToModule CalledFromModule
    Dim i As Long
    For i = 1 To Module.CodeModule.CountOfLines
        If InStr(1, Module.CodeModule.Lines(i, 1), "Sub " & CalledFromProcedure) > 0 Or _
                                                                                 InStr(1, Module.CodeModule.Lines(i, 1), "Function " & CalledFromProcedure) > 0 Then
            Module.CodeModule.CodePane.SetSelection i, 1, i, 1
            Exit Sub
        End If
    Next
HELL:
End Sub

Private Sub CommandButton1_Click()
    '#INCLUDE TreeviewClear
    '#INCLUDE FindCodeEverywhere
    '#INCLUDE TreeviewAssignProjectImages
    '#INCLUDE TreeviewExpandAllNodes
    Dim tvtop As Long, tvleft As Long
    
    'TreeView1.Visible = False
    TreeviewClear TreeView1
    FindCodeEverywhere TextBox1, TreeView1
    TreeviewAssignProjectImages TreeView1
    TreeviewExpandAllNodes TreeView1

    'TreeView1.Visible = True
    TreeView1.Nodes(1).Expanded = True
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then
        CommandButton1_Click
    End If
End Sub

Private Sub TreeView1_DblClick()
    '#INCLUDE TreeviewGotoProjectElement
    TreeviewGotoProjectElement TreeView1
End Sub


