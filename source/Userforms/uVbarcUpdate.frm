VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uVbarcUpdate 
   Caption         =   "vbArc Addin Update"
   ClientHeight    =   8484.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7848
   OleObjectBlob   =   "uVbarcUpdate.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uVbarcUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uVbarcUpdate
'* Created    : 06-10-2022 10:41
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim changeLog As String

Private Sub Label1_Click()
    SkipThisVersion
End Sub

Private Sub Label2_Click()
    FollowLink ("https://github.com/alexofrhodes")
End Sub

Private Sub Label3_Click()
    Dim TargetFolder As String
    TargetFolder = getFolder
    If TargetFolder = "" Then Exit Sub
    Dim SaveFileFullPath As String
    SaveFileFullPath = TargetFolder & "\" & getFilePartName(PROJECT_DOWNLAOD_URL, True)
    DownloadFile PROJECT_DOWNLAOD_URL, SaveFileFullPath
    Do While Not FileExists(SaveFileFullPath)
        DoEvents
    Loop
    FollowLink TargetFolder
    If openAfterDownload Then Workbooks.Open (SaveFileFullPath)
End Sub

Private Sub UserForm_Initialize()
    If GetInternetConnectedState = False Then
        MsgBox "Seems Internet is not available"
        Unload Me
    End If
    
    changeLog = TXTReadFromUrl(PROJECT_CHANGELOG_URL)
    TextBox1.TEXT = changeLog
    
    
End Sub

Sub DownloadLatestVersion()
    
End Sub

