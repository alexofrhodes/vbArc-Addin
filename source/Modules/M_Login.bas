Attribute VB_Name = "M_Login"
Option Explicit

Rem @Folder Login
Sub ListRightsSheets()
    Dim coll As New Collection
    Dim uLoginSettingsSheet As Worksheet
    Set uLoginSettingsSheet = ThisWorkbook.SHEETS("uLoginSettings")
    Dim rng As Range
    Dim ws As Worksheet
    Dim cell As Range
    Set cell = uLoginSettingsSheet.Range("H4")
    Set rng = Union(cell, cell.End(xlToRight))
    rng.ClearContents
    For Each ws In ThisWorkbook.Worksheets
        cell = ws.Name
        Set cell = cell.OFFSET(0, 1)
    Next
End Sub

Sub CheckUser()
    On Error Resume Next
    Dim UserRow, SheetCol As Long, SheetNm As String
    With ThisWorkbook.SHEETS("uLoginSettings")
        .Calculate
        If .Range("B8").Value = Empty Then
            MsgBox "Please enter a correct user name"
            Exit Sub
        End If
        If .Range("B7").Value <> True Then
            MsgBox "Pleae enter a correct password"
            Exit Sub
        End If
        uLogin.Hide
        UserRow = .Range("B8").Value
        For SheetCol = 8 To 13
            SheetNm = .Cells(4, SheetCol).Value
            If .Cells(UserRow, SheetCol).Value = .[B1] Then
                SHEETS(SheetNm).Unprotect "123"
                SHEETS(SheetNm).visible = xlSheetVisible
            End If
            If .Cells(UserRow, SheetCol).Value = .[b2] Then
                SHEETS(SheetNm).Protect "123"
                SHEETS(SheetNm).visible = xlSheetVisible
            End If
            If .Cells(UserRow, SheetCol).Value = .[b3] Then SHEETS(SheetNm).visible = xlVeryHidden
        Next SheetCol
    End With
End Sub

Sub UserLogOff()
    '#INCLUDE HideWorksheets
    Dim ans As Long
    ans = MsgBox("Log Off?", vbYesNo)
    If ans = vbYes Then
        SHEETS("uLoginSettings").Range("B5,B6").ClearContents
        HideWorksheets
    End If
End Sub

Sub UserLoginStart()
    uLogin.Show
End Sub

Sub HideWorksheets()
    Dim WkSht As Worksheet
    ThisWorkbook.SHEETS("uLogin").Activate
    For Each WkSht In ThisWorkbook.Worksheets
        If WkSht.Name <> "uLogin" Then WkSht.visible = xlSheetVeryHidden
    Next WkSht
End Sub

