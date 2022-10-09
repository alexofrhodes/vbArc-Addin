VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSessions 
   Caption         =   "vbaCodeArchive ~ Sessions Manager"
   ClientHeight    =   6072
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11496
   OleObjectBlob   =   "uSessions.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uSessions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uSessions
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Function GetOpenFolders() As Collection
    Dim coll As Collection
    Set coll = New Collection
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            Debug.Print Wnd.document.Folder.Self.Path
            coll.Add Wnd.document.Folder.Self.Path
        End If
    Next Wnd
    Set GetOpenFolders = coll
    Set coll = Nothing
End Function

Private Sub chIncludeWorkbooks_Click()
    If chIncludeWorkbooks.Value = True Then
        chClose.visible = True
        chClose.Value = False
    Else
        chClose.visible = False
        chClose.Value = False
    End If
End Sub

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub UserForm_Initialize()
    cbSessionName.AddItem "LAST SESSION"
    cbSessionName.ListIndex = -1
    cbSessionName.TEXT = "<SESSION NAME>"
    PopulateSessions
End Sub

Private Sub cDELETE_Click()
    If LSessions.list(LSessions.ListIndex) = "LAST SESSION" Then
        MsgBox "You can overwrite Last Session but not delete it"
    Else
        SessionDelete
    End If
    ThisWorkbook.Save
End Sub

Private Sub cDeleteItems_Click()
    RemoveItemFromSessionBooks
End Sub

Private Sub cLOAD_Click()
    If LSessions.ListIndex = -1 Then Exit Sub
    SessionOpen
End Sub

Private Sub cSAVE_Click()
    Dim newSessionName As String
    newSessionName = cbSessionName.TEXT
    If newSessionName = vbNullString Or newSessionName = "<SESSION NAME>" Then cbSessionName.ListIndex = 0
    SessionSave
    ThisWorkbook.Save
End Sub

Private Sub LSessions_Click()
    PopulateSessionBooks
    cbSessionName.TEXT = LSessions.list(LSessions.ListIndex)
End Sub

Sub RemoveItemFromSessionBooks()
    Dim r As Long, c As Long
    Dim i As Long
    c = ThisWorkbook.SHEETS("uSessions").rows(1).Find(uSessions.LSessions.list(uSessions.LSessions.ListIndex), LookAt:=xlWhole).Column
    For r = uSessions.LSessionBooks.ListCount - 1 To 0 Step -1
        If uSessions.LSessionBooks.SELECTED(r) Then
            ThisWorkbook.SHEETS("uSessions").Cells(r + 2, c).Delete Shift:=xlUp
            uSessions.LSessionBooks.RemoveItem r
        End If
    Next r
    ThisWorkbook.Save
End Sub

Sub SessionSave()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("uSessions")
    Dim foundRange As Range
    Dim listCol As Long
    Set foundRange = ws.rows(1).Find(uSessions.cbSessionName.TEXT, LookAt:=xlWhole)
    If foundRange Is Nothing Then
        listCol = ws.Cells(1, 1).CurrentRegion.Columns.count + 1
        ws.Cells(1, listCol).Value = uSessions.cbSessionName.TEXT
    Else
        listCol = foundRange.Column
        If Application.WorksheetFunction.CountA(ws.Columns(listCol)) > 1 Then
            Dim LastRow As Long
            LastRow = ws.Cells(rows.count, listCol).End(xlUp).row
            ws.Range(ws.Cells(2, listCol), ws.Cells(LastRow, listCol)).ClearContents
        End If
    End If
    Dim RowNum As Long
    RowNum = 1
    Dim counterFolders As Long
    Dim counterFiles As Long
    If chIncludeWorkbooks.Value = True Then
        Dim wbO As Workbook
        For Each wbO In Application.Workbooks
            counterFiles = counterFiles + 1
            RowNum = RowNum + 1
            ws.Cells(RowNum, listCol).Value = wbO.FullName
            If uSessions.chClose.Value = True Then
                If wbO.Name <> ActiveWorkbook.Name Then
                    wbO.Close savechanges:=True
                End If
            End If
        Next wbO
    End If
    If uSessions.chIncludeFolders.Value = True Then
        Dim element As Variant
        For Each element In GetOpenFolders
            counterFolders = counterFolders + 1
            RowNum = RowNum + 1
            ws.Cells(RowNum, listCol).Value = element
        Next element
    End If
    If counterFiles + counterFolders > 0 Then
        If uSessions.cbSessionName.TEXT <> "LAST SESSION" Then
            uSessions.LSessions.AddItem uSessions.cbSessionName.TEXT
        End If
        If uSessions.LSessions.ListCount = 1 Then
            uSessions.LSessions.ListIndex = -1
            uSessions.LSessions.ListIndex = 0
        End If
        MsgPOP "Session " & UCase(uSessions.cbSessionName.TEXT) & " saved with" & vbNewLine & _
                                                                counterFiles & " Workbooks and " & counterFolders & " Folders", 2
    Else
        MsgPOP "Only folders option: No open folders found"
    End If
End Sub

Sub SessionOpen()
    Dim element As Variant
    If ListboxSelectedCount(uSessions.LSessionBooks) = 0 Then
        Dim i As Long
        For i = 0 To uSessions.LSessionBooks.ListCount - 1
            FollowLink (uSessions.LSessionBooks.list(i))
        Next i
    Else
        For Each element In ListboxSelectedValues(uSessions.LSessionBooks)
            FollowLink (element)
        Next element
    End If
    ThisWorkbook.Activate
    MsgPOP "Selected session items loaded"
End Sub

Sub SessionDelete()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("uSessions")
    Dim foundRange As Range
    Set foundRange = ws.rows(1).Find(uSessions.cbSessionName.TEXT, LookAt:=xlWhole)
    foundRange.EntireColumn.Delete
    uSessions.LSessions.RemoveItem (uSessions.LSessions.ListIndex)
    MsgPOP "Session " & UCase(uSessions.cbSessionName.TEXT) & " deleted"
End Sub

Sub PopulateSessions()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("uSessions")
    Dim listCol As Long
    For listCol = 1 To ws.Cells(1, 1).CurrentRegion.Columns.count
        uSessions.LSessions.AddItem ws.Cells(1, listCol).TEXT
    Next listCol
End Sub

Sub PopulateSessionBooks()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("uSessions")
    Dim foundRange As Range
    Dim listCol As Long
    Set foundRange = ws.rows(1).Find(uSessions.LSessions.list(uSessions.LSessions.ListIndex), LookAt:=xlWhole)
    listCol = foundRange.Column
    Dim i As Long
    Dim LastRow As Long
    LastRow = ws.Cells(ws.rows.count, listCol).End(xlUp).row
    uSessions.LSessionBooks.clear
    For i = 2 To LastRow
        uSessions.LSessionBooks.AddItem ws.Cells(i, listCol).Value
    Next i
End Sub


