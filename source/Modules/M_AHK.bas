Attribute VB_Name = "M_AHK"
Rem @Folder AHK
Sub AHK_vbeMenu()
    '#INCLUDE dp
    '#INCLUDE RunAHKScript
    If Application.VBE.MainWindow.visible = True Then dp "The hotkey will be CTRL + SHIFT + H"
    RunAHKScript "C:\Users\acer\Dropbox\SOFTWARE\AHK\0 EXCEL\vbaMenu\vbaMenu.ahk"
End Sub

Public Function RunAHKScript(AHKFilePath As String, Optional AHKExePath As String = vbNullString) As Long
    '#INCLUDE GetAppPath
    If AHKExePath = vbNullString Then AHKExePath = GetAppPath("AutoHotkey.exe")
    Const QUOTES As String = """"
    RunAHKScript = Shell(AHKExePath & Space(1) & QUOTES & AHKFilePath & QUOTES)
End Function

Public Function GetAppPath(AppName As String) As String
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    Dim ReadFrom As String
    Const APP_PATH_REG_LOCATION As String = "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\"
    ReadFrom = APP_PATH_REG_LOCATION & AppName & "\"
    On Error Resume Next
    GetAppPath = WshShell.RegRead(ReadFrom)
End Function

Public Sub StopAHKScript()
    Shell "taskkill /pid " & AHKSheet.Range("B2").Value & " /pid " & AHKSheet.Range("B3").Value
End Sub

