VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uMemoryKnots 
   Caption         =   "Notebooks"
   ClientHeight    =   6300
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7656
   OleObjectBlob   =   "uMemoryKnots.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uMemoryKnots"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uMemoryKnots
'* Created    : 06-10-2022 10:38
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *


'Public WithEvents WorksheetSelectionChangeCheck As Excel.Worksheet

Dim memoPath As String
Dim FolderToZip As String
Dim MemoryKnotsWB As Workbook
Dim MemoryKnotsWS As Worksheet
Dim cell As Range
Dim i As Long
Dim str As String
Dim strTMP As String
Dim tmpWS As Worksheet
Dim KeepLoading As Boolean
Dim RestoreTo As Worksheet

Private Sub AddBook_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    str = InputBox("New NoteBook name")
    If str = "" Then Exit Sub
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS.Add(After:=MemoryKnotsWB.SHEETS(MemoryKnotsWB.SHEETS.count))
    noteBOOKS.AddItem ("o" & UCase(str))
    With MemoryKnotsWS
        .Name = "o" & UCase(str)
        .[A1] = "DATE"
        .[B1] = "NOTES"
        .[A1:B1].Font.Bold = True
        '        .Visible = xlSheetHidden
    End With
End Sub

Sub AddNotes()
    '    noteLIST.ListIndex = -1
    Call ListboxClearSelection(noteLIST)
    noteBOX.TEXT = ""
    Me.Width = 400
    noteBOX.SetFocus
End Sub

Private Sub AddNote_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call AddNotes
End Sub

Private Sub Clear_FilterNoteBooks_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Dim tmpl As String
    tmpl = noteBOOKS.list(noteBOOKS.ListIndex)
    FilterNoteBooks.TEXT = ""
    For i = 0 To noteBOOKS.ListCount - 1
        If noteBOOKS.list(i) = tmpl Then
            noteBOOKS.SELECTED(i) = True
        End If
    Next i
End Sub

Private Sub Clear_FilterNoteList_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FilterNoteList.TEXT = ""
    noteLIST.SetFocus

    '    For i = 0 To noteLIST.ListCount - 1
    '        If noteLIST.List(i) = tmpl Then
    '            noteLIST.Selected(i) = True
    '        End If
    '    Next i
End Sub

Private Sub CloseNoteBook_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If noteBOOKS.ListIndex < 0 Then
        MsgBox "No selection"
        Exit Sub
    End If
    
    MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).visible = False
    If Not tmpWS Is Nothing Then
        tmpWS.Activate
    End If
End Sub

Private Sub cmdCopyAll_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    Dim msg As String
    Dim i As Long, j As Long
    With noteLIST
        For i = 0 To .ListCount - 1
            msg = msg & .list(i) & vbCrLf
        Next i
    End With
    msg = left(msg, Len(msg) - 2)
    Call CLIP(msg)
    MsgBox "Selection copied"

End Sub

Private Sub cmdCopyAllLinked_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If
   
    If Listbox_Selected(noteLIST, 1) < 2 Then
        MsgBox "What's the sound of one hand clapping?" & Chr(10) & Chr(10) & _
                                                                            Space(50) & "~Fortune Cookie"
        Exit Sub
    End If
   
    Dim link As String
    link = InputBox("Choose link")
    If link = "" Then
        MsgBox "No input"
        Exit Sub
    End If
    
    Dim msg As String
    Dim i As Long, j As Long
    With noteLIST
        For i = 0 To .ListCount - 1
            msg = msg & .list(i) & link
        Next i
    End With
    msg = left(msg, Len(msg) - 1)
    Call CLIP(msg)
End Sub

Private Sub cmdCopySelected_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    Dim answer As Long
    answer = MsgBox("(YES) link with line break" & Chr(10) & _
                    "(NO) link with your choice", vbYesNoCancel)

    If answer = vbCancel Then Exit Sub

    Dim msg As String
    Dim i As Long, j As Long
    With noteLIST
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                If answer = vbYes Then
                    msg = msg & .list(i) & vbCrLf
                Else
                    msg = msg & .list(i) & link
                End If
            End If
        Next i
    End With
    
    If answer = vbYes Then
        msg = left(msg, Len(msg) - 2)
    Else
        msg = left(msg, Len(msg) - 1)
    End If
    
    Call CLIP(msg)
    MsgBox "Selection copied"
End Sub

Private Sub cmdExport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExportNotes
End Sub

Private Sub cmdExportAsImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExportAsImage
End Sub

Private Sub cmdExportAsImageMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ExportAsImage
End Sub

Sub ExportAsImage()
    If Not TypeName(Selection) = "Range" Then Exit Sub

    Dim JPGfolder As String
    JPGfolder = memoPath        '& "JPG"
    'On Error Resume Next
    'MkDir JPGfolder
    'On Error GoTo 0

    Dim action As Long

    'If Selection.Areas.count > 1 Then
    action = MsgBox("(YES) = for each area in selection" & Chr(10) & _
                    "(NO) = for each cell in selection", vbYesNoCancel)
    If action = vbCancel Then Exit Sub
    'Else
    '    action = vbNo
    'End If

    On Error Resume Next        'goto 0
    Application.DisplayAlerts = False

    Dim JPGcell As Range
    Dim result As String
    Dim ImageExtension As String
    ImageExtension = InputBox("Choose extension" & Chr(10) & _
                              "(1) = jpg" & Chr(10) & _
                              "(2) = bmp" & Chr(10) & _
                              "(3) = gib" & Chr(10) & _
                              "(4) = ico" & Chr(10) & _
                              "(5) = cur" & Chr(10) & _
                              "(6) = wmf", Default:=2)
    If Not IsNumeric(ImageExtension) Or ImageExtension < 1 Or ImageExtension > 6 Then
        Exit Sub
    End If

    Select Case ImageExtension
        Case Is = 1
            ImageExtension = "jpg"
        Case Is = 2
            ImageExtension = "bmp"
        Case Is = 3
            ImageExtension = "gib"
        Case Is = 4
            ImageExtension = "ico"
        Case Is = 5
            ImageExtension = "cur"
        Case Is = 6
            ImageExtension = "wmf"
    End Select


    Select Case action

        Case Is = vbNo
            For Each JPGcell In Selection
                Call ExportRangeAsImage(ActiveSheet, JPGcell, JPGfolder, JPGcell.Value, ImageExtension)
                Application.Wait (Now + TimeValue("0:00:01"))

                Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
                Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
                cell = Now()
                cell.OFFSET(0, 1) = JPGcell.Value & "." & ImageExtension
                noteLIST.AddItem JPGcell.Value & "." & ImageExtension

            Next JPGcell

        Case Is = vbYes

            For i = 1 To Selection.Areas.count
                result = ""
                result = InputBox("name for image of area: " & Selection.Areas(i).Address)

                Call ExportRangeAsImage(ActiveSheet, Selection.Areas(i), JPGfolder, result, ImageExtension)
                Application.Wait (Now + TimeValue("0:00:01"))

                Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
                Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
                cell = Now()
                cell.OFFSET(0, 1) = result & "." & ImageExtension
                noteLIST.AddItem result & "." & ImageExtension
    
            Next i

    End Select

    Application.DisplayAlerts = True
End Sub

Private Sub cmdImport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ImportNotes
End Sub

Private Sub cmdInsertComment_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    If Selection.Cells.count > 1 Then
        MsgBox ("choose 1 cell to insert comment")
        Exit Sub
    End If

    Dim answer As Long

    answer = MsgBox("(YES) link with line break" & Chr(10) & _
                    "(NO) link with your choice", vbYesNoCancel)

    If answer = vbCancel Then Exit Sub

    If answer = vbNo Then
        Dim link As String
        link = InputBox("Choose link")
    End If
    
    Dim msg As String
    With noteLIST
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                If answer = vbYes Then
                    msg = msg & .list(i) & vbCrLf
                Else
                    msg = msg & .list(i) & link
                End If
            End If
        Next i
    End With
    
    If answer = vbYes Then
        msg = left(msg, Len(msg) - 2)
    Else
        msg = left(msg, Len(msg) - 1)
    End If
    
    Dim action As Long
    If ActiveCell.Comment Is Nothing Then
        ActiveCell.AddComment Format(Now(), "yymmdd hhmm") & " " & msg
    Else
        action = MsgBox("(YES) pretend comment" & Chr(10) & _
                        "(NO) replace comment", vbYesNoCancel)
        If action = vbCancel Then Exit Sub
        If action = vbYes Then
            msg = Format(Now(), "yy/mm/dd hh:mm") & Chr(10) & Chr(10) & msg & Chr(10) & Chr(10) & ActiveCell.Comment.TEXT
            ActiveCell.Comment.Delete
            ActiveCell.AddComment msg
        Else
            ActiveCell.Comment.Delete
            ActiveCell.AddComment Format(Now(), "yy/mm/dd hh:mm") & Chr(10) & Chr(10) & msg
        End If
    End If
    ActiveCell.Comment.Shape.TextFrame.AutoSize = True
    
    '    Cells(ActiveCell.Row + 1, ActiveCell.Column).Activate

End Sub

Private Sub cmdInsertSelectedToCells_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Not TypeName(Selection) = "Range" Then Exit Sub
    
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If


    Dim j As Long

    Dim answer As Long
    answer = MsgBox("(YES) insert vertically" & Chr(10) & _
                    "(NO) insert horizontally", vbYesNoCancel)

    If answer = vbCancel Then Exit Sub

    Dim result As Long

    If answer = vbYes Then
        If Selection.Cells.count = 1 Then
            j = 0
            With noteLIST
                For i = 0 To .ListCount - 1
                    If .SELECTED(i) Then
                        Cells(ActiveCell.row + j, ActiveCell.Column).Value = .list(i)
                        j = j + 1
                    End If
                Next i
            End With
        Else
            result = MsgBox("You've selected " & Selection.Cells.count & " cells." & Chr(10) & Chr(10) & _
                            "Proceed to insert selected items vertically for each cell?", vbYesNo)
            If result = vbNo Then Exit Sub
            If result = vbYes Then
                For Each cell In Selection.Cells
                    j = 0
                    With noteLIST
                        For i = 0 To .ListCount - 1
                            If .SELECTED(i) Then
                                Cells(cell.row + j, cell.Column).Value = .list(i)
                                j = j + 1
                            End If
                        Next i
                    End With
                Next cell
            End If
    
        End If
    Else
        If Selection.Cells.count = 1 Then
            j = 0
            With noteLIST
                For i = 0 To .ListCount - 1
                    If .SELECTED(i) Then
                        Cells(ActiveCell.row, ActiveCell.Column + j).Value = .list(i)
                        j = j + 1
                    End If
                Next i
            End With
        Else
            result = MsgBox("You've selected " & Selection.Cells.count & " cells." & Chr(10) & Chr(10) & _
                            "Proceed to insert selected items vertically for each cell?", vbYesNo)
            If result = vbYes Then
                For Each cell In Selection.Cells
                    j = 0
                    With noteLIST
                        For i = 0 To .ListCount - 1
                            If .SELECTED(i) Then
                                Cells(cell.row, cell.Column + j).Value = .list(i)
                                j = j + 1
                            End If
                        Next i
                    End With
                Next cell
            End If
        End If
    End If
End Sub

Private Sub cmdInsertSelectedMerged_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If TypeName(Selection) <> "Range" Then Exit Sub
    
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    Dim answer As Long
    answer = MsgBox("(YES) link with line break" & Chr(10) & _
                    "(NO) link with your choice", vbYesNoCancel)

    If answer = vbCancel Then Exit Sub

    If answer = vbNo Then
        Dim link As String
        link = InputBox("Choose link")
    End If
    
    Dim msg As String
    With noteLIST
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                If answer = vbYes Then
                    msg = msg & .list(i) & vbCrLf
                Else
                    msg = msg & .list(i) & link
                End If
            End If
        Next i
    End With
    
    If answer = vbYes Then
        ActiveCell.Value = left(msg, Len(msg) - 2)
    Else
        ActiveCell.Value = left(msg, Len(msg) - 1)
    End If
    '    Cells(ActiveCell.Row + 1, ActiveCell.Column).Activate

End Sub

Private Sub cmdMail_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If
    
    Dim WshShell As Object
    Set WshShell = CreateObject("WScript.Shell")
    '   Dim FolderToZip As String
    FolderToZip = WshShell.SpecialFolders("MyDocuments")
    FolderToZip = FolderToZip & "\FolderToZip\"
    
    Call FoldersCreate("FolderToZip")
    
    Dim AttachPath As String
    AttachPath = MailAttachments
    
    On Error Resume Next
    Kill memoPath & "zipNotes.zip"
    On Error GoTo 0
    
    Dim msg As String
    With noteLIST
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                msg = msg & .list(i) & vbCrLf
            End If
        Next i
    End With
    
    If OutlookCheck = True Then
        'For Tips see: http://www.rondebruin.nl/win/winmail/Outlook/tips.htm
        'Working in Office 2000-2016

        Dim OutApp As Object
        Dim OutMail As Object

        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)

        On Error Resume Next
        With OutMail
            .To = ""
            .CC = ""
            .BCC = ""
            .Subject = "MemoryKnots"
            .body = msg
            If AttachPath <> "" Then
                .Attachments.Add (MailAttachments)
            End If
            .Display
            '        .Send
        End With
        
        Set OutMail = Nothing
        Set OutApp = Nothing
    Else
        Call CLIP(msg)
        MsgBox "Outlook was not found" & Chr(10) & _
                                                 "Notes have been COPIED" & Chr(10) & _
                                                 "Please go to your mail and PASTE" & Chr(10) & _
                                                 "If you've included wav/image notes, they have been zipped at " & Chr(10) & _
                                                 memoPath & "zipNotes.zip"
    End If
    
    '   On Error GoTo 0
    On Error Resume Next
    Kill FolderToZip & "*.*"
    RmDir FolderToZip

End Sub

Function MailAttachments()
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")

    Dim FileToCopy As String
    With noteLIST
        For i = 0 To .ListCount - 1
            If .SELECTED(i) = True Then
                If (Right(.list(i), 3) = "wav" _
                    Or Right(.list(i), 3) = "jpg" Or _
                    Right(.list(i), 3) = "bmp" Or _
                    Right(.list(i), 3) = "gib" Or _
                    Right(.list(i), 3) = "ico" Or _
                    Right(.list(i), 3) = "cur" Or _
                    Right(.list(i), 3) = "wmf") Then
                    FileToCopy = memoPath & .list(i)
                    fso.CopyFile FileToCopy, FolderToZip
                End If
            End If
        Next i
    End With

    If MailAttachmentsCount > 0 Then
        Call Zip(FolderToZip, memoPath & "zipNotes.zip")
        '        Call ZipFolder(FolderToZip, MemoPath & "wav.zip")
        MailAttachments = memoPath & "zipNotes.zip"
    Else
        MailAttachments = ""
    End If

End Function

Function MailAttachmentsCount()
    Dim Path As String, count As Integer

    Path = FolderToZip & "*.*"

    fileName = Dir(Path)

    Do While fileName <> ""
        count = count + 1
        fileName = Dir()
    Loop
    MailAttachmentsCount = count
End Function

Private Sub cmdMoveNote_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 0 Then Exit Sub
    '    Set cell = MemoryKnotsWB.Sheets(noteBOOKS.List(noteBOOKS.ListIndex)).Columns("B:B").Find( _
    '        What:=noteLIST.List(noteLIST.ListIndex), _
    '        LookIn:=xlFormulas, _
    '        LookAt:=xlWhole, _
    '        SearchOrder:=xlByRows, _
    '        SearchDirection:=xlNext, _
    '        MatchCase:=False, _
    '        SearchFormat:=False)
    
    noteSELECT.clear
    For i = 0 To noteBOOKS.ListCount - 1
        noteSELECT.AddItem noteBOOKS.list(i)
    Next i
    noteSELECT.ListIndex = -1
    FrameSelect.Width = 84
    FrameSelect.ZOrder (0)
    FrameSelect.visible = True
      
End Sub

Private Sub cmdMoveOK_Click()

    'If noteSELECT.ListIndex <> -1 Then
    '    strTMP = noteSELECT.List(noteSELECT.ListIndex)
    '    cell.EntireRow.Copy MemoryKnotsWB.Sheets(strTMP).Range("A" & Rows.Count).End(xlUp).Offset(1, 0)
    '    cell.EntireRow.Delete
    '    noteBOX.Value = ""
    '    noteLIST.RemoveItem (noteLIST.ListIndex)
    '
    '    noteBOOKS.ListIndex = -1
    '    For i = 0 To noteBOOKS.ListCount - 1
    '        If noteBOOKS.List(i) = strTMP Then
    '            noteBOOKS.Selected(i) = True
    '        Exit For
    '        End If
    '    Next i
    '
    '    noteSELECT.Clear
    '    FrameSelect.Visible = False
    'Else
    '    strTMP = ""
    '    noteSELECT.Clear
    '    FrameSelect.Visible = False
    'End If

    If noteSELECT.ListIndex = -1 Then
        MsgBox "List empty or no item selected"
        noteSELECT.clear
        FrameSelect.visible = False
        Exit Sub
    Else

        Dim moveWhat As Variant
        moveWhat = Split(Listbox_Selected(noteLIST, 2), ",")
        Dim moveTo As String
        moveTo = noteSELECT.list(noteSELECT.ListIndex)
            
        For i = UBound(moveWhat) To LBound(moveWhat) Step -1
        
            Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(moveWhat(i)), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=True, _
        SearchFormat:=False)
            
            cell.EntireRow.Copy MemoryKnotsWB.SHEETS(moveTo).Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
            cell.EntireRow.Delete
    
        Next i
        noteSELECT.clear
        FrameSelect.visible = False
        
    End If
            
    For i = UBound(moveWhat) To LBound(moveWhat) Step -1
        noteLIST.RemoveItem (i)
    Next i
        
    If ToggleExtra.Value = False Then
        FrameExtra.Width = 5
        FrameExtra.visible = False
    End If
End Sub

Private Sub cmdMoveX_Click()
    '    strTMP = ""
    noteSELECT.clear
    FrameSelect.visible = False
End Sub

Private Sub cmdNewNoteFromSelection_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ListboxClearSelection(noteLIST)
    noteBOX.TEXT = SelectionValues(Chr(10))
    Me.Width = 400
    noteBOX.SetFocus
End Sub

Private Sub cmdNewNoteFromSelectionMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ListboxClearSelection(noteLIST)
    noteBOXmini.TEXT = SelectionValues(Chr(10))
    noteBOXmini.SetFocus
End Sub

Private Sub cmdNewNotesFromSelection_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    Dim varr As Variant
    varr = Split(SelectionValues("``"), "``")
    Call ListboxClearSelection(noteLIST)
    noteBOX.TEXT = ""
    For i = 0 To UBound(varr)
        noteBOX.TEXT = varr(i)
        Call NoteSave
    Next
End Sub

Private Sub cmdNewNotesFromSelectionMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim i As Long
    Dim varr As Variant
    varr = Split(SelectionValues("``"), "``")
    Call ListboxClearSelection(noteLIST)
    noteBOXmini.TEXT = ""
    For i = 0 To UBound(varr)
        noteBOXmini.TEXT = varr(i)
        Call NoteSaveMini
    Next
End Sub

Private Sub cmdPlayWAV_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Listbox_Selected(noteLIST, 1) = 1 Then
        If Right(noteLIST.list(noteLIST.ListIndex), 3) = "wav" Then
            VoiceToWav.play memoPath & noteLIST.list(noteLIST.ListIndex)
        Else
            MsgBox "Not wav file, or file was manually deleted"
        End If
    End If
End Sub

Private Sub cmdSettingsNotebooks_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim result As String
    result = InputBox("Change noteBOOKS font size", Default:=Me.noteBOOKS.Font.Size)
    If result = "" Then Exit Sub
    If Not IsNumeric(result) Then
        MsgBox "Font size must be numeric"
        Exit Sub
    End If
    ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteBooksFontSize").Value = result
    Me.noteBOOKS.Font.Size = result
End Sub

Private Sub cmdSettingsNotebox_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim result As String
    result = InputBox("Change noteLIST font size", Default:=Me.noteBOX.Font.Size)
    If result = "" Then Exit Sub
    If Not IsNumeric(result) Then
        MsgBox "Font size must be numeric"
        Exit Sub
    End If
    ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteBoxFontSize").Value = result
    Me.noteBOX.Font.Size = result
End Sub

Private Sub cmdSettingsNotelist_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Dim result As String
    result = InputBox("Change noteLIST font size", Default:=Me.noteLIST.Font.Size)
    If result = "" Then Exit Sub
    If Not IsNumeric(result) Then
        MsgBox "Font size must be numeric"
        Exit Sub
    End If
    ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteListFontSize").Value = result
    Me.noteLIST.Font.Size = result
End Sub

Private Sub cmdSpeechToWavStartRecording_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VoiceNoteRecord
End Sub

Sub VoiceNoteRecord()
    VoiceToWav.StartRecord Bits16, Sampels32000, Mono
    cmdSpeechToWavStopRecording.visible = True
    cmdSpeechToWavStartRecording.visible = False
End Sub

Private Sub cmdSpeechToWavStartRecordingMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VoiceNoteRecord
End Sub

Private Sub cmdSpeechToWavStopRecording_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VoiceNoteSave
End Sub

Sub VoiceNoteSave()
    VoiceToWav.SaveRecord Environ("USERPROFILE") & "\Desktop\" & "tmp.wav"
    
    result = ""
    result = InputBox("Filename")
    If result = "" Then
        Exit Sub
    End If
    result = result & ".wav"
    
    Name Environ("USERPROFILE") & "\Desktop\" & "tmp.wav" As memoPath & result

    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
    cell = Now()
    cell.OFFSET(0, 1) = result
    noteLIST.AddItem result

    cmdSpeechToWavStopRecording.visible = False
    cmdSpeechToWavStartRecording.visible = True

End Sub

Private Sub cmdSpeechToWavStopRecordingMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call VoiceNoteSave
End Sub

Private Sub cmdTextToWav_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SaveTextToWav(1)
End Sub

Sub SaveTextToWav(NormalOrMini As Long)
    'run this to make a wav file from a text input
    Dim sP As String, sFN As String, sStr As String, sFP As String
    Dim result As String

    'set parameter values - insert your own profile name first
    'paths
    result = InputBox("Record title?" & Chr(10) & Chr(10) & "Overwrites if file name same")
    If result = "" Then
        Exit Sub
    End If
    sFN = result & ".wav"
    
    Select Case NormalOrMini
        Case Is = 1
            result = noteBOX.TEXT
        Case Is = 2
            result = noteBOXmini.TEXT
    End Select

    If result = "" Then
        '        Kill ThisWorkbook.Path & "\" & result
        Exit Sub
    End If
    'string to use for the recording
    sStr = result
    
    'make voice wav file from string
    '"My.wav" 'overwrites if file name same
    sP = memoPath
    sFP = sP & sFN
    StringToWavFile sStr, sFP
    
    Application.ScreenUpdating = False
    Dim tmpSheet As Worksheet
    Set tmpSheet = ActiveSheet
    Dim i As Integer
    Dim j As Variant
        
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
    cell = Now()
    cell.OFFSET(0, 1) = sFN
    noteLIST.AddItem sFN
    
    Call ListboxClearSelection(noteLIST)
    Select Case NormalOrMini
        Case Is = 1
            noteBOX.TEXT = ""
        Case Is = 2
            noteBOXmini.TEXT = ""
    End Select

End Sub

Private Sub cmdTextToWavMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call SaveTextToWav(2)
End Sub

Private Sub DeleteBook_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If noteBOOKS.ListIndex < 0 Then
        MsgBox "Notebooks list empty or none selected"
        Exit Sub
    End If
    If noteBOOKS.list(noteBOOKS.ListIndex) = "o NOTES" _
                                             Or noteBOOKS.list(noteBOOKS.ListIndex) = "o RESOLVED" Then
        Exit Sub
    End If
    Application.DisplayAlerts = False
    MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Delete
    noteBOOKS.RemoveItem (noteBOOKS.ListIndex)
    Application.DisplayAlerts = True
End Sub

Private Sub DeleteNote_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    '    If noteLIST.ListIndex < 0 Then
    '        MsgBox "Note list empty or none selected"
    '        Exit Sub
    '    End If
    '    Set cell = MemoryKnotsWB.Sheets(noteBOOKS.List(noteBOOKS.ListIndex)).Columns("B:B").Find( _
    '        What:=noteLIST.List(noteLIST.ListIndex), _
    '        LookIn:=xlFormulas, _
    '        LookAt:=xlWhole, _
    '        SearchOrder:=xlByRows, _
    '        SearchDirection:=xlNext, _
    '        MatchCase:=False, _
    '        SearchFormat:=False)
    '    cell.EntireRow.Delete
    '    noteBOX.Value = ""
    '    noteLIST.RemoveItem (noteLIST.ListIndex)

    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    'remove item
    Dim deleteWhat As Variant
    deleteWhat = Split(Listbox_Selected(noteLIST, 2), ",")
            
    'move backwards when deleting
    For i = UBound(deleteWhat) To LBound(deleteWhat) Step -1
    
        Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(deleteWhat(i)), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
        cell.EntireRow.Delete
        On Error Resume Next
        With noteLIST
            If Right(.list(i), 3) = "wav" Or Right(.list(i), 3) = "jpg" Or Right(.list(i), 3) = "bmp" Or Right(.list(i), 3) = "gib" _
                                                                                                                              Or Right(.list(i), 3) = "ico" Or Right(.list(i), 3) = "cur" Or Right(.list(i), 3) = "wmf" Then
                Kill memoPath & noteLIST.list(deleteWhat(i))
            End If
        End With
        noteLIST.RemoveItem (deleteWhat(i))
    
    Next i
    
    noteBOX.TEXT = ""
    If ToggleExtra.Value = False Then
        FrameExtra.Width = 5
        FrameExtra.visible = False
    End If
    
    Me.Width = 174
End Sub

Private Sub DynamicImage_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Shell("explorer.exe" & " " & memoPath & noteLIST.list(noteLIST.ListIndex), vbNormalFocus)
End Sub

Private Sub FilterNoteBooks_Change()
    On Error GoTo eh
    'Reload list so if you type and delete you'll get the items back
    Call LoadNoteBooks
    
    Dim i               As Long
    Dim n               As Long
    Dim str             As String
    Dim sTemp           As String
   
    'Equals is always case sensitive
    'Remove LCase if you want it to be case sensitive
    str = LCase(FilterNoteBooks.TEXT)
   
    n = noteBOOKS.ListCount
   
    For i = n - 1 To 0 Step -1        'Work backwards when deleting items
        'Equals is always case sensitive
        'Remove LCase if you want it to be case sensitive
        sTemp = LCase(noteBOOKS.list(i))
       
        If InStr(sTemp, str) = 0 Then
            noteBOOKS.RemoveItem (i)
            'Exit Sub   'Uncomment to Exit if value found
        End If
    Next i
    
    noteBOOKS.SELECTED(0) = True
    Exit Sub
eh:
End Sub

Private Sub FilterNoteList_Change()
    On Error GoTo eh
    'Reload list so if you type and delete you'll get the items back
    Call LoadNoteList
    
    Dim i               As Long
    Dim n               As Long
    Dim str             As String
    Dim sTemp           As String
   
    'Equals is always case sensitive
    'Remove LCase if you want it to be case sensitive
    str = LCase(FilterNoteList.TEXT)
   
    n = noteLIST.ListCount
   
    For i = n - 1 To 0 Step -1        'Work backwards when deleting items
        'Equals is always case sensitive
        'Remove LCase if you want it to be case sensitive
        sTemp = LCase(noteLIST.list(i))
       
        If InStr(sTemp, str) = 0 Then
            noteLIST.RemoveItem (i)
            'Exit Sub   'Uncomment to Exit if value found
        End If
    Next i
    
    noteLIST.SELECTED(0) = True
    Exit Sub
eh:
End Sub

Private Sub cmdResolved_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)

    If noteBOOKS.list(noteBOOKS.ListIndex) = "o RESOLVED" Then Exit Sub

    If Listbox_Selected(noteLIST, 1) = 0 Or noteLIST.ListCount = 0 Then
        MsgBox "List empty or no item selected"
        Exit Sub
    End If

    Dim resolveWhat As Variant
    resolveWhat = Split(Listbox_Selected(noteLIST, 2), ",")
            
    For i = LBound(resolveWhat) To UBound(resolveWhat)
    
        Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(resolveWhat(i)), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
            
        cell.EntireRow.Copy MemoryKnotsWB.SHEETS("o RESOLVED").Range("A" & rows.count).End(xlUp).OFFSET(1, 0)
        cell.EntireRow.Delete
    
    Next i
    For i = UBound(resolveWhat) To LBound(resolveWhat) Step -1
        noteLIST.RemoveItem resolveWhat(i)
    Next i
    
    noteBOX.Value = ""
    
    If ToggleExtra.Value = False Then
        FrameExtra.Width = 5
        FrameExtra.visible = False
    End If
    
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub noteBOOKS_Change()
    On Error GoTo eh
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    noteLIST.clear
    Call LoadNoteList
eh:
End Sub

Private Sub noteBOOKS_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If noteBOOKS.ListIndex = -1 Then Exit Sub
    If noteBOOKS.list(noteBOOKS.ListIndex) = "o NOTES" Or _
                                             noteBOOKS.list(noteBOOKS.ListIndex) = "o RESOLVED" Then
        MsgBox "Can't touch this" & Chr(10) & _
                                            "          ~mc Hammer"
        Exit Sub
    End If

    Dim result As String
    result = InputBox("New NoteBook name")
    If result = "" Then Exit Sub
    If WorksheetExists("o" & UCase(result)) Then
        MsgBox "Name taken"
        Exit Sub
    End If
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    
    With MemoryKnotsWS
        .Name = "o" & UCase(result)
    End With
    
    noteBOOKS.list(noteBOOKS.ListIndex) = "o" & UCase(result)
    Set MemoryKnotsWS = Nothing
End Sub

Private Sub noteBOX_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.Width = 174
End Sub

Private Sub noteLIST_Change()
    DynamicImage.visible = False

    If Listbox_Selected(noteLIST, 1) <= 1 Then        'noteLIST.ListCount > 0 And noteLIST.ListIndex >= 0 And
        noteBOX.TEXT = noteLIST.list(noteLIST.ListIndex)
        If ToggleExtra.Value = False Then
            FrameExtra.Width = 5
            FrameExtra.visible = False
        End If
    Else
        noteBOX.TEXT = ""
        FrameExtra.Width = 23
        FrameExtra.ZOrder (0)
        FrameExtra.visible = True
    End If
      
    If Listbox_Selected(noteLIST, 1) = 1 Then
        If Right(noteLIST.list(Listbox_Selected(noteLIST, 2)), 3) = "wav" Then
            cmdPlayWAV.visible = True
            cmdPlayWAV.ZOrder (0)
            DynamicImage.visible = False
            Me.Width = 174
        ElseIf Right(noteLIST.list(noteLIST.ListIndex), 3) = "jpg" Or _
                                                             Right(noteLIST.list(noteLIST.ListIndex), 3) = "bmp" Or _
                                                             Right(noteLIST.list(noteLIST.ListIndex), 3) = "gib" Or _
                                                             Right(noteLIST.list(noteLIST.ListIndex), 3) = "ico" Or _
                                                             Right(noteLIST.list(noteLIST.ListIndex), 3) = "cur" Or _
                                                             Right(noteLIST.list(noteLIST.ListIndex), 3) = "wmf" Then
            cmdPlayWAV.visible = False
            DynamicImage.visible = True
            '                On Error Resume Next
            '                DynamicImage.Picture = LoadPicture(MemoPath & noteLIST.List(noteLIST.ListIndex))
            '                DynamicImage.ZOrder (0)
            '                Me.Width = 400
        Else
            cmdPlayWAV.visible = False
            DynamicImage.visible = False

        End If
    Else
        cmdPlayWAV.visible = False
        DynamicImage.visible = False
        Me.Width = 174
    End If
End Sub

Private Sub noteLIST_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If noteLIST.ListCount = 0 Then Exit Sub
    If Right(noteLIST.list(noteLIST.ListIndex), 3) = "wav" Then
        str = InputBox("Rename WAV file" & Chr(10) & Chr(10) & "Will replace file with same name.")
        If str = "" Then
            Exit Sub
        Else
            On Error Resume Next
            Name memoPath & noteLIST.list(noteLIST.ListIndex) As _
                                                              memoPath & str & ".wav"
        
            Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
            Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(noteLIST.ListIndex), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
            cell.OFFSET(0, -1) = Now()
            cell = str & ".wav"
        
            noteLIST.list(noteLIST.ListIndex) = str & ".wav"
            
        End If
    ElseIf Right(noteLIST.list(noteLIST.ListIndex), 3) = "jpg" Or Right(noteLIST.list(noteLIST.ListIndex), 3) = "bmp" Or Right(noteLIST.list(noteLIST.ListIndex), 3) = "gib" _
                                                                                                                                                                       Or Right(noteLIST.list(noteLIST.ListIndex), 3) = "ico" Or Right(noteLIST.list(noteLIST.ListIndex), 3) = "cur" Or Right(noteLIST.list(noteLIST.ListIndex), 3) = "wmf" Then

        str = InputBox("Rename file" & Chr(10) & Chr(10) & "Will replace file with same name.")
        If str = "" Then
            Exit Sub
        Else
            On Error Resume Next
            Name memoPath & noteLIST.list(noteLIST.ListIndex) As _
                                                              memoPath & str & Right(noteLIST.list(noteLIST.ListIndex), 3)
        
            Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
            Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(noteLIST.ListIndex), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
            cell.OFFSET(0, -1) = Now()
            cell = str & Right(noteLIST.list(noteLIST.ListIndex), 3)
        
            noteLIST.list(noteLIST.ListIndex) = str & Right(noteLIST.list(noteLIST.ListIndex), 3)
            noteBOX.TEXT = str & Right(noteLIST.list(noteLIST.ListIndex), 3)
        End If
    Else
        Me.Width = 400
        noteBOX.SetFocus
    End If
End Sub

Private Sub OpenNoteBook_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If noteBOOKS.ListIndex < 0 Then
        MsgBox "No selection"
        Exit Sub
    End If

    Set tmpWS = ActiveSheet
    With MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
        .visible = True
        .Activate
    End With

End Sub

Private Sub OpenBakupFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Shell("explorer.exe" & " " & memoPath, vbNormalFocus)
End Sub

Private Sub cmdNoteSave_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoteSave
End Sub

Sub NoteSave()
    If noteBOX.TEXT = "" Then
        MsgBox "Write a note first"
        Exit Sub
    End If

    Dim msg As String
    '    If ToggleRangeNote.Value = True And Selection.Cells.count = 1 Then
    '        msg = ActiveCell.Address(False, False) & " " & noteBOX.Text
    '    Else
    msg = noteBOX.TEXT
    '    End If
    
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    
    If Listbox_Selected(noteLIST, 1) = 0 Then
        Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
        cell = Now()
        cell.OFFSET(0, 1) = msg
        noteLIST.AddItem (msg)
    Else
        Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(noteLIST.ListIndex), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
        cell.OFFSET(0, -1) = Now()
        cell = msg
        
        noteLIST.list(noteLIST.ListIndex) = msg
    End If
    
    Call ListboxClearSelection(noteLIST)
    noteBOX.TEXT = ""
End Sub

Private Sub cmdNoteSaveMini_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call NoteSaveMini
End Sub

Sub NoteSaveMini()
    If noteBOXmini.TEXT = "" Then
        MsgBox "Write a note first"
        Exit Sub
    End If

    Dim msg As String
    '    If ToggleRangeNote.Value = True And Selection.Cells.count = 1 Then
    '        msg = ActiveCell.Address(False, False) & " " & noteBOXmini.Text
    '    Else
    msg = noteBOXmini.TEXT
    '    End If
    
    Set MemoryKnotsWS = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex))
    
    If Listbox_Selected(noteLIST, 1) = 0 Then
        Set cell = MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).OFFSET(1, 0)
        cell = Now()
        cell.OFFSET(0, 1) = msg
        noteLIST.AddItem (msg)
    Else
        Set cell = MemoryKnotsWB.SHEETS(noteBOOKS.list(noteBOOKS.ListIndex)).Columns("B:B").Find( _
        What:=noteLIST.list(noteLIST.ListIndex), _
        LookIn:=xlFormulas, _
        LookAt:=xlWhole, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlNext, _
        MatchCase:=False, _
        SearchFormat:=False)
        
        cell.OFFSET(0, -1) = Now()
        cell = msg
        
        noteLIST.list(noteLIST.ListIndex) = msg
    End If
    
    noteBOXmini.TEXT = ""
End Sub

Private Sub cmdSortNoteBooksAZ_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Dim tmpl As String
    tmpl = noteBOOKS.list(noteBOOKS.ListIndex)
    Call ListboxSortAZ(noteBOOKS)
    For i = 0 To noteBOOKS.ListCount - 1
        If noteBOOKS.list(i) = tmpl Then
            noteBOOKS.SELECTED(i) = True
        End If
    Next i

End Sub

Private Sub cmdSortNoteBooksZA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    Dim tmpl As String
    tmpl = noteBOOKS.list(noteBOOKS.ListIndex)
    Call ListboxSortZA(noteBOOKS)
    '    Call ListboxClearSelection(notebOoks)
    For i = 0 To noteBOOKS.ListCount - 1
        If noteBOOKS.list(i) = tmpl Then
            noteBOOKS.SELECTED(i) = True
        End If
    Next i
End Sub

Private Sub cmdSortNoteListAZ_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    '    Dim tmpl As String
    '    tmpl = noteLIST.List(noteLIST.ListIndex)
    Call ListboxSortAZ(noteLIST)
    Call ListboxClearSelection(noteLIST)
    '    For i = 0 To noteLIST.ListCount - 1
    '        If noteLIST.List(i) = tmpl Then
    '            noteLIST.Selected(i) = True
    '        End If
    '    Next i
End Sub

Private Sub cmdSortNoteListZA_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    '    Dim tmpl As String
    '    tmpl = noteLIST.List(noteLIST.ListIndex)
    Call ListboxSortZA(noteLIST)
    Call ListboxClearSelection(noteLIST)
    '    For i = 0 To noteLIST.ListCount - 1
    '        If noteLIST.List(i) = tmpl Then
    '            noteLIST.Selected(i) = True
    '        End If
    '    Next i
End Sub

Private Sub Toggle_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    'If notbooks.ListIndex = -1 Then
    '    noteBOOKS.Selected(0) = True
    'End If
    Call ListboxClearSelection(noteLIST)
    If Me.Height > 100 Then
        Me.Height = 64        '50
        Me.Width = 174
        Me.Caption = noteBOOKS.list(noteBOOKS.ListIndex)
        cmdNoteSaveMini.visible = True
        noteBOXmini.visible = True
        cmdSpeechToWavStartRecordingMini.visible = True
        cmdSpeechToWavStopRecordingMini.visible = True
        cmdTextToWavMini.visible = True
        cmdNewNoteFromSelectionMini.visible = True
        cmdNewNotesFromSelectionMini.visible = True
        cmdExportAsImageMini.visible = True
        noteBOXmini.SetFocus
    Else
        Me.Height = 336
        'Me.Width = 174
        Me.Caption = "Notebooks"
        cmdNoteSaveMini.visible = False
        noteBOXmini.visible = False
        cmdSpeechToWavStartRecordingMini.visible = False
        cmdSpeechToWavStopRecordingMini.visible = False
        cmdTextToWavMini.visible = False
        cmdNewNoteFromSelectionMini.visible = False
        cmdNewNotesFromSelectionMini.visible = False
        cmdExportAsImageMini.visible = False
        noteBOOKS.SetFocus
    End If
End Sub

Private Sub ToggleExtra_Click()
    If ToggleExtra.Value = True Then
        FrameExtra.Width = 23
        FrameExtra.ZOrder (0)
        FrameExtra.visible = True
    Else
        FrameExtra.Width = 5
        FrameExtra.visible = False
    End If
End Sub

Private Sub UserForm_Initialize()
    '/load position
    If GetSetting("My Settings Folder", Me.Name, "Left Position") = "" _
                                                                    And GetSetting("My Settings Folder", Me.Name, "Top Position") = "" Then
        Me.StartUpPosition = 1        ' CenterOwner
    Else
        Me.left = GetSetting("My Settings Folder", Me.Name, "Left Position")
        Me.top = GetSetting("My Settings Folder", Me.Name, "Top Position")
    End If
    'load position/

    Me.Width = 174
    Me.noteBOOKS.Font.Size = ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteBooksFontSize").Value
    Me.noteLIST.Font.Size = ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteListFontSize").Value
    Me.noteBOX.Font.Size = ThisWorkbook.SHEETS("MemoryKnots_Settings").Range("noteBoxFontSize").Value
    Me.noteBOXmini.Font.Size = 8        'ThisWorkbook.Sheets("MemoryKnots_Settings").Range("noteBoxFontSize").Value
    

    Set MemoryKnotsWB = Workbooks(ThisWorkbook.Name)

    '    Set MemoryKnotsWS = MemoryKnotsWB.Sheets(1)
        
    Call LoadNoteBooks

    noteBOOKS.SELECTED(0) = True

    Set WorksheetSelectionChangeCheck = ActiveSheet
    
    memoPath = Environ$("USERPROFILE") & "\My Documents\vbArc\MemoryKnots\"
    FoldersCreate memoPath
End Sub

'Sub WorksheetSelectionChangeCheck_SelectionChange(ByVal Target As Range)
'    If ToggleRangeNote.Value = False Then Exit Sub
'    FilterNoteBooks.Text = ""
'    FilterNoteList.Text = ""
'
'    Dim noteLISTvalues As String
'
'    For i = 0 To noteBOOKS.ListCount - 1
'        noteLISTvalues = noteLISTvalues & " " & noteBOOKS.List(i)
'        Debug.Print noteLISTvalues
'    Next i
'
'    If InStr(LCase(noteBOOKS.List(noteBOOKS.ListIndex)), LCase(ActiveSheet.Name)) > 0 _
'    And Selection.Cells.count = 1 Then
'    For i = 0 To noteLIST.ListCount - 1
'        noteLISTvalues = noteLISTvalues & " " & noteLIST.List(i)
'        Debug.Print noteLISTvalues
'    Next i
'    If InStr(noteLISTvalues, ActiveCell.Address(False, False)) > 0 Then
'        FilterNoteList.Text = ActiveCell.Address(False, False)
'    Else
'        FilterNoteList.Text = ""
'    End If
'End If
'End Sub

'Sub ListboxValues(LBox As MSForms.ListBox)
'
'End Sub
'
'Sub CheckUpdate()
'    On Error GoTo eh
'    If UCase(noteLIST.List(0)) = "UPDATE" Then
'        Call ImportNotes
'    End If
'eh:
'End Sub

Private Sub LoadNoteList()
    noteLIST.clear
    Dim i As Long
    For i = 2 To MemoryKnotsWS.Cells(rows.count, 1).End(xlUp).row
        noteLIST.AddItem MemoryKnotsWS.Cells(i, 2)
    Next
End Sub

Private Sub LoadNoteBooks()
    noteBOOKS.clear
    For i = 1 To MemoryKnotsWB.SHEETS.count
        If MemoryKnotsWB.SHEETS(i).Name Like "o *" Then
            noteBOOKS.AddItem MemoryKnotsWB.SHEETS(i).Name
        End If
    Next
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    '/save position
    'must have uf position set to manual
    SaveSetting "My Settings Folder", Me.Name, "Left Position", Me.left
    SaveSetting "My Settings Folder", Me.Name, "Top Position", Me.top


End Sub


