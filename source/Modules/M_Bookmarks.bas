Attribute VB_Name = "M_Bookmarks"
Rem @Folder Bookmarks
Sub BookmarkSave(Optional index As Long = 0)
    '#INCLUDE CommandBarBuilder
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE UpdateBookmarkLabel
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("VbeBookmarks", ThisWorkbook)
    Dim cell As Range
    If index = 0 Then
        Set cell = ws.Range("A9999").End(xlUp)
        If cell <> "" Then Set cell = cell.OFFSET(1, 0)
        index = cell.row
    Else
        Set cell = ws.Cells(index, 1)
    End If
    Dim delim As String
    delim = " | "
    Dim Procedure As String
    On Error Resume Next
    Procedure = ActiveProcedure
    On Error GoTo 0
    If Procedure = "" Then Procedure = "N/A"
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim BookmarkLine As String
    BookmarkLine = ActiveCodepaneWorkbook.Name & delim & _
                   Module.Name & delim & _
                   Procedure & delim & _
                   Module.CodeModule.Lines(CodePaneSelectionStartLine, 1)
    cell = cell.row
    cell.OFFSET(0, 1) = BookmarkLine
    If index < 11 Then
        UpdateBookmarkLabel index, IIf(Procedure <> "N/A", Procedure, Module.Name)
        CommandBarBuilder ThisWorkbook.SHEETS("BAR_Bookmarks")
    End If
End Sub

Sub UpdateBookmarkLabel(index As Long, newLabel As String)
    Application.EnableEvents = False
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("BAR_Bookmarks")
    Dim cell As Range
    Set cell = ws.Columns(3).SpecialCells(xlCellTypeConstants).Find("bmSave" & index, LookAt:=xlWhole)
    cell.OFFSET(0, -1).Value = newLabel
    Set cell = ws.Columns(3).SpecialCells(xlCellTypeConstants).Find("bmLoad" & index, LookAt:=xlWhole)
    cell.OFFSET(0, -1).Value = newLabel
    Application.EnableEvents = True
End Sub

Sub ResetBookmarkLabels()
    '#INCLUDE CommandBarBuilder
    '#INCLUDE UpdateBookmarkLabel
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("BAR_Bookmarks")
    Dim cell As Range
    Dim index As Long
    For index = 1 To 10
        UpdateBookmarkLabel index, CStr(index)
    Next
    CommandBarBuilder ws
End Sub

Sub bmSave1()
    '#INCLUDE BookmarkSave
    BookmarkSave 1
End Sub

Sub bmSave2()
    '#INCLUDE BookmarkSave
    BookmarkSave 2
End Sub

Sub bmSave3()
    '#INCLUDE BookmarkSave
    BookmarkSave 3
End Sub

Sub bmSave4()
    '#INCLUDE BookmarkSave
    BookmarkSave 4
End Sub

Sub bmSave5()
    '#INCLUDE BookmarkSave
    BookmarkSave 5
End Sub

Sub bmSave6()
    '#INCLUDE BookmarkSave
    BookmarkSave 6
End Sub

Sub bmSave7()
    '#INCLUDE BookmarkSave
    BookmarkSave 7
End Sub

Sub bmSave8()
    '#INCLUDE BookmarkSave
    BookmarkSave 8
End Sub

Sub bmSave9()
    '#INCLUDE BookmarkSave
    BookmarkSave 9
End Sub

Sub bmSave10()
    '#INCLUDE BookmarkSave
    BookmarkSave 10
End Sub

Sub BookmarkList()
    '#INCLUDE dp
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("VbeBookmarks", ThisWorkbook)
    If ws.Cells(1, 1) = "" Then
        MsgBox "No bookmarks found"
        Exit Sub
    End If
    Dim cell As Range, rng As Range
    For Each cell In ws.Columns(1).SpecialCells(xlCellTypeConstants)
        If rng Is Nothing Then
            Set rng = Union(cell, cell.OFFSET(0, 1))
        Else
            Set rng = Union(rng, Union(cell, cell.OFFSET(0, 1)))
        End If
    Next
    dp rng
End Sub

Sub BookmarkLoad(Optional index As Long = 0, Optional movingForward As Boolean)
    '#INCLUDE GoToModule
    '#INCLUDE ProcedureExists
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE GetProcText
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("VbeBookmarks")
    Dim lr As Long
    lr = ws.Cells(9999, 1).End(xlUp).row
    Dim cell As Range
retry:
    If index = 0 Then
        Set cell = ws.Cells(lr, 1)
        index = cell.row
    Else
        Set cell = ws.Cells(index, 1)
    End If
    If cell = "" Then Exit Sub
    Dim delim As String
    delim = " | "
    Dim var
    var = Split(cell.OFFSET(0, 1), delim)
    Dim targetworkbookname As String
    targetworkbookname = var(0)
    Dim ModuleName As String
    ModuleName = var(1)
    Dim Procedure As String
    Procedure = var(2)
    Dim BookmarkLine As String
    BookmarkLine = var(3)
    If targetworkbookname <> ActiveCodepaneWorkbook.Name Then
        index = index + IIf(movingForward = True, 1, -1)
        GoTo retry
    End If
    Dim wb As Workbook
    Dim Module As VBComponent
    On Error Resume Next
    Set wb = Workbooks(targetworkbookname)
    Set Module = wb.VBProject.VBComponents(ModuleName)
    If Module Is Nothing Then Set Module = Workbooks(targetworkbookname).VBProject.VBComponents(ModuleName)
    On Error GoTo 0
    If wb Is Nothing Then
        index = index + IIf(movingForward = True, 1, -1)
        GoTo retry
    End If
    ws.Range("O1").Value = index
    If Module Is Nothing Then Exit Sub
    GoToModule Module
    If Procedure = "N/A" Then
    ElseIf ProcedureExists(Procedure, wb) Then
        ProcFirstline = ProcedureStartLine(Module, Procedure)
        Module.CodeModule.CodePane.SetSelection ProcFirstline, 1, ProcFirstline, 1
        If BookmarkLine <> "" Then
            If InStr(1, GetProcText(Module, Procedure), BookmarkLine) > 0 Then
                Dim i As Long
                For i = ProcedureStartLine(Module, Procedure) To ProcedureEndLine(Module, Procedure)
                    If InStr(1, Module.CodeModule.Lines(i, 1), BookmarkLine, vbTextCompare) > 0 Then
                        Module.CodeModule.CodePane.SetSelection i, 1, i, 1
                        Exit Sub
                    End If
                Next
            End If
        End If
    Else
        Debug.Print "Procedure " & Procedure & " not found in workbook " & targetworkbookname
    End If
End Sub

Rem @TODO unidentified error
Rem         causes the subs to run multiple times
Rem         when calling from vbe button but not
Rem         when calling from immediate window
Rem     Dim ws As Worksheet
Rem     Set ws = ThisWorkbook.Sheets("VbeBookmarks")
Rem
Rem     Dim LastBookmark As Range
Rem     Set LastBookmark = ws.Range("O1")
Rem         LastBookmark.Value = LastBookmark.Value + 1
Rem
Rem     Dim index As Long
Rem         index = LastBookmark.Value
Rem
Rem     Dim lr As Long
Rem         lr = ws.Cells(9999, 1).End(xlUp).row
Rem
Rem     If index > lr Then
Rem         LastBookmark.Value = 1
Rem         index = 1
Rem     End If
Rem
Rem     Dim LoadThisBookmark As Range
Rem     Set LoadThisBookmark = ws.Cells(index, 1)
Rem
Rem     Do While LoadThisBookmark.TEXT = vbNullString
Rem         index = index + 1
Rem         LastBookmark.Value = index
Rem         If index > lr Then
Rem             LastBookmark.Value = vbNullString
Rem             Exit Sub
Rem         End If
Rem         Set LoadThisBookmark = ws.Cells(index, 1)
Rem     Loop
Rem     BookmarkLoad index
Rem End Sub
Rem
Rem Sub BookMarkPrevious()
Rem     Dim ws As Worksheet
Rem     Set ws = ThisWorkbook.Sheets("VbeBookmarks")
Rem
Rem     Dim LastBookmark As Range
Rem     Set LastBookmark = ws.Range("O1")
Rem         LastBookmark.Value = LastBookmark.Value - 1
Rem
Rem     Dim index As Long
Rem         index = LastBookmark.Value
Rem
Rem     Dim lr As Long
Rem         lr = ws.Cells(9999, 1).End(xlUp).row
Rem
Rem     If index < 1 Then
Rem         LastBookmark.Value = lr
Rem         index = lr
Rem     End If
Rem
Rem     Dim LoadThisBookmark As Range
Rem     Set LoadThisBookmark = ws.Cells(index, 1)
Rem
Rem     Do While LoadThisBookmark.TEXT = vbNullString
Rem         index = index - 1
Rem         LastBookmark.Value = index
Rem         If index < 1 Then
Rem             LastBookmark.Value = vbNullString
Rem             Exit Sub
Rem         End If
Rem         Set LoadThisBookmark = ws.Cells(index, 1)
Rem     Loop
Rem     BookmarkLoad index, False
Rem End Sub
Sub bmload1()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 1
End Sub

Sub bmload2()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 2
End Sub

Sub bmload3()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 3
End Sub

Sub bmload4()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 4
End Sub

Sub bmload5()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 5
End Sub

Sub bmload6()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 6
End Sub

Sub bmload7()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 7
End Sub

Sub bmload8()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 8
End Sub

Sub bmload9()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 9
End Sub

Sub bmload10()
    '#INCLUDE BookmarkLoad
    BookmarkLoad 10
End Sub

Sub BookmarkDelete(Optional index As Long)
    '#INCLUDE CommandBarBuilder
    '#INCLUDE UpdateBookmarkLabel
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("VbeBookmarks", ThisWorkbook)
    Dim cell As Range
    If index = 0 Then
        Set cell = ws.Cells(9999, 1).End(xlUp)
        cell.RESIZE(1, 2).clear
    Else
        ws.Cells(index, 1).RESIZE(1, 2).clear
    End If
    If index > 0 And index <= 10 Then UpdateBookmarkLabel index, CStr(index)
    CommandBarBuilder ws
End Sub

Sub BookmarkReset()
    '#INCLUDE ResetBookmarkLabels
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("VbeBookmarks", ThisWorkbook)
    ws.Cells.clear
    ResetBookmarkLabels
End Sub


