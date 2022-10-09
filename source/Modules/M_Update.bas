Attribute VB_Name = "M_Update"

Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwFlags As Long, ByVal dwReserved As Long) As Long
#Else
    Private Declare Function InternetGetConnectedState Lib "wininet.dll" (ByRef dwflags As Long, ByVal dwReserved As Long) As Long
#End If

Function DayAfterCheck() As Long
    Dim TbRange     As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    DayAfterCheck = TbRange.Cells(1, 5)
End Function

Public Function GetInternetConnectedState() As Boolean
    GetInternetConnectedState = InternetGetConnectedState(0&, 0&)
End Function

Rem * Created    : 15-09-2019 15:48
Rem * Author     : VBATools
Rem * Contacts   : http://vbatools.ru/ https://vk.com/vbatools
Rem * Copyright  : VBATools.ru
Public Sub StartUpdate()
    '#INCLUDE InvalidateControl
    '#INCLUDE SetControlValue
    '#INCLUDE GetInternetConnectedState
    '#INCLUDE ShowUpdateMsg
    '#INCLUDE isUpdateAvailable
    If Not GetInternetConnectedState Then Exit Sub
    Dim TbRange     As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    Dim ShowUpdateAvailable As Range
    Set ShowUpdateAvailable = TbRange.Cells(1, 4)
    ShowUpdateAvailable.Value = isUpdateAvailable
    SetControlValue "MainButtonUpdate", "visible", ShowUpdateAvailable.Value
    InvalidateControl "MainButtonUpdate"
    If ShowUpdateAvailable.Value = True Then ShowUpdateMsg
End Sub

Private Sub ShowUpdateMsg()
    '#INCLUDE DayAfterCheck
    '#INCLUDE SkipThisVersion
    On Error GoTo ErrorHandler
    Dim TbRange     As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    Dim TextUpdate  As String
    TextUpdate = TbRange.Cells(1, 3).Value2
    If TextUpdate <> vbNullString And TbRange.Cells(1, 2).Value2 + DayAfterCheck < Now() Then
        Rem @TODO Move this IF block to a userform which will also show by Update button in vbArc Tab
        If MsgBox("Greetings!" & vbNewLine & _
                  "Update " & TextUpdate & "is available for " & PROJECT_NAME & vbNewLine & vbNewLine & _
                  "To update, go to the website " & AUTHOR_GITHUB & vbNewLine & _
                  "or use the now visible UPDATE icon in your vbArc Tab" & vbNewLine & _
                  "This message will not show again until the next update check." & vbNewLine & _
                  "Skip this version?" & vbNewLine, vbYesNo, "Updating " & PROJECT_NAME) = vbYes Then
            SkipThisVersion
        End If
        TbRange.Cells(1, 2).Value2 = Now()
        ThisWorkbook.Save
    End If
    If TextUpdate <> TbRange.Cells(1, 1).Value2 And TextUpdate <> vbNullString Then TbRange.Cells(1, 4).Value = True
    Exit Sub
ErrorHandler:
    Select Case err.Number
        Case 1004:
        Case Else:
            Debug.Print "Mistake! in ShowUpdateMsg" & vbLf & err.Number & vbLf & err.Description & vbCrLf & "in the line" & Erl
    End Select
    err.clear
End Sub

Private Function isUpdateAvailable() As Boolean
    '#INCLUDE TXTReadFromUrl
    '#INCLUDE DayAfterCheck
    '#INCLUDE ChekDateUpdate
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Dim TbRange     As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    Dim NewVersion As String, CurentVersion As String
    If ChekDateUpdate Then
        NewVersion = TXTReadFromUrl(PROJECT_CHANGELOG_URL)
        NewVersion = Split(NewVersion, vbLf)(0)
        NewVersion = Split(NewVersion, " ")(UBound(Split(NewVersion, " ")))
        NewVersion = Trim(Replace(NewVersion, vbLf, ""))
    End If
    If NewVersion <> vbNullString Then
        CurentVersion = TbRange.Cells(1, 1).Value
        TbRange.Cells(1, 3).Value2 = NewVersion
        If CurentVersion < NewVersion Then
            isUpdateAvailable = True
            ThisWorkbook.Save
        Else
            GoTo SaveLabel
        End If
    Else
        GoTo SaveLabel
    End If
    Application.DisplayAlerts = True
    Exit Function
SaveLabel:
    isUpdateAvailable = False
    TbRange.Cells(1, 2).Value2 = Now() + DayAfterCheck
    ThisWorkbook.Save
    Application.DisplayAlerts = True
    Exit Function
ErrorHandler:
    Select Case err.Number
        Case 1004, -2146697211:
        Case Else:
            Debug.Print "Mistake! in isUpdateAvailable" & vbLf & err.Number & vbLf & err.Description & vbCrLf & "in the line" & Erl
    End Select
    err.clear
    isUpdateAvailable = False
End Function

Private Function ChekDateUpdate() As Boolean
    '#INCLUDE DayAfterCheck
    ChekDateUpdate = False
    Dim TbRange     As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    Dim LastCheckedDate As Date
    LastCheckedDate = CDate(TbRange.Cells(1, 2).Value2)
    If Now > LastCheckedDate + DayAfterCheck Then ChekDateUpdate = True
End Function

Sub SkipThisVersion()
    '#INCLUDE TXTReadFromUrl
    '#INCLUDE InvalidateControl
    '#INCLUDE SetControlValue
    Dim TbRange As Range
    Set TbRange = ThisWorkbook.SHEETS("vbArc_Addin_Settings").ListObjects("Update").DataBodyRange
    Dim latestVersion As String
    Dim changeLog As String
    changeLog = TXTReadFromUrl(PROJECT_CHANGELOG_URL)
    latestVersion = Split(changeLog, vbLf)(0)
    latestVersion = Split(latestVersion, " ")(UBound(Split(latestVersion, " ")))
    Rem to skip this version @TODO add new column?
    TbRange.Cells(1, 1).Value2 = latestVersion
    Rem Hide update button
    SetControlValue "MainButtonUpdate", "visible", "FALSE"
    InvalidateControl "MainButtonUpdate"
End Sub

Private Function ResponseTextHttp(ByVal URL As String) As String
    Dim oHttp       As Object
    On Error Resume Next
    Set oHttp = CreateObject("MSXML2.XMLHTTP")
    If err.Number <> 0 Then
        Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
    End If
    On Error GoTo 0
    If oHttp Is Nothing Then
        ResponseTextHttp = vbNullString
        Exit Function
    End If
    With oHttp
        .Open "GET", URL, False
        .send
        If .Status = 200 Then
            ResponseTextHttp = .responseText
        Else
            ResponseTextHttp = vbNullString
        End If
    End With
    Set oHttp = Nothing
End Function


