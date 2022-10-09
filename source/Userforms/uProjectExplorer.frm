VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uProjectExplorer 
   Caption         =   "Project Explorer"
   ClientHeight    =   9264.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3636
   OleObjectBlob   =   "uProjectExplorer.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uProjectExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uProjectExplorer
'* Created    : 06-10-2022 10:39
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub TreeView1_Click()
    '#INCLUDE
    TreeviewGotoProjectElement TreeView1
End Sub

Private Sub UserForm_Initialize()
    InitializeProjectExplorer
End Sub

Sub InitializeProjectExplorer()
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveModule
    '#INCLUDE ProjectExplorer
    '#INCLUDE TreeviewAllProjects
    '#INCLUDE ImageListLoadProjectIcons
    '#INCLUDE TreeviewAssignProjectImages
    '#INCLUDE TreeviewSelectNodes
    '#INCLUDE MakeUserFormChildOfVBEditor
    Application.VBE.MainWindow.visible = True
    MakeUserFormChildOfVBEditor Me.Caption
    TreeviewAllProjects TreeView1
    With TreeView1
        .Sorted = True
        .Appearance = ccFlat
        .LineStyle = tvwRootLines
        .Style = tvwTreelinesPlusMinusPictureText
        .Font.Size = 10
        .Indentation = 2
    End With
    ImageListLoadProjectIcons ImageList1, TreeView1
    TreeviewAssignProjectImages TreeView1
    If Application.VBE.MainWindow.visible = False Then
        Set TargetWorkbook = ActiveWorkbook
    Else
        Set TargetWorkbook = ActiveCodepaneWorkbook
    End If
    TreeviewSelectNodes TreeView1, True, TargetWorkbook.Name, Array(ActiveModule.Name)
End Sub

