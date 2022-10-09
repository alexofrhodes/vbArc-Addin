VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFileManager 
   Caption         =   "File Manager"
   ClientHeight    =   9504.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   10140
   OleObjectBlob   =   "uFileManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uFileManager
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub CommandButton10_Click()
    OleVbaRun ListboxSelectedValues(ListBox1)
    MsgBox "Done"
End Sub

Private Sub CommandButton11_Click()
    Dim out As New Collection
    Set out = ListboxSelectedValues(ListBox1)
    Dim element
    For Each element In out
        SplitATextFileintoIndividualOnes element
    Next
    MsgBox "Done"
End Sub

Private Sub CommandButton12_Click()
    Dim out As New Collection
    Set out = ListboxSelectedValues(ListBox1)
    Dim element
    For Each element In out
        TxtRemoveComments element
    Next
    MsgBox "Done"
End Sub

Private Sub CommandButton3_Click()
    ListboxToRangeSelect ListBox1
End Sub

Private Sub CommandButton4_Click()
    '    TextBox1.TEXT = ""
    SelectDeselectAll ListBox1, True
End Sub

Private Sub CommandButton5_Click()
    '    TextBox1.TEXT = ""
    SelectDeselectAll ListBox1, False
End Sub

Private Sub CommandButton6_Click()
    Dim out As New Collection
    Set out = ListboxSelectedValues(ListBox1)
    Dim Path As String
    Path = Environ$("USERPROFILE") & "\Documents\vbArc\MergedTXT\"
    FoldersCreate Path
    MergeFileText out, Path & "Merged " & Format(Now, "YY-MM-DD HHNN") & ".txt"
    
    MsgBox "Done"
    FollowLink Path
    
End Sub

Private Sub CommandButton7_Click()
    Dim element As Variant
    For Each element In ListboxSelectedValues(ListBox1)
        PretendListOfContainedProceduresInTXT CStr(element)
    Next
    MsgBox "Done"
End Sub

Private Sub CommandButton8_Click()
    Dim out As New Collection
    Set out = ListboxSelectedValues(ListBox1)
    Dim element
    For Each element In out
        TxtRemoveBlankLines element
    Next
    MsgBox "Done"
End Sub

Private Sub CommandButton9_Click()

End Sub

Private Sub ListViewControl_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim dbBefore As Long: dbBefore = ListBox1.ListCount
    Dim FileFullPath As String
    Dim fileItem As Long
    Dim objFSO As Scripting.FileSystemObject
    Dim objTopFolder As Scripting.Folder
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim var As Variant, element As Variant
    For fileItem = 1 To Data.Files.count
        FileFullPath = Data.Files(fileItem)
        If oLogFiles = True Then
            If LCase(isFDU(FileFullPath)) = "f" Then
                var = Split(TextBox2.TEXT, ",")
                On Error Resume Next
                If left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
                For Each element In var
                    If InStr(1, FileFullPath, element, vbTextCompare) > 0 And left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" Then
PASS:
                        If ListboxContains(ListBox1, FileFullPath) = False Then
                            AddToListBox ListBox1, FileFullPath
                        End If
                    End If
                Next
            Else        'if drag dropped folder
                Set objTopFolder = objFSO.getFolder(FileFullPath)
                FileRecursive objTopFolder, oSearchInSubfolders.Value
            End If
        End If
        If oLogFolders = True Then
            If UCase(isFDU(FileFullPath)) = "D" Then
                Set objTopFolder = objFSO.getFolder(FileFullPath)
                If ListboxContains(ListBox1, objTopFolder.Path & "\") = False Then
                    AddToListBox ListBox1, objTopFolder.Path
                End If
                FolderRecursive objTopFolder, oSearchInSubfolders.Value
            End If
        End If
    Next fileItem
    
    If ListBox1.ListCount - dbBefore > 0 Then ListboxToDatabaseSheet
    
    Set objFSO = Nothing
    Set objTopFolder = Nothing
End Sub

Private Sub ListboxToDatabaseSheet()
    Dim rng As Range
    Set rng = ThisWorkbook.SHEETS("FileManager_DB").Range("A1")
    rng.CurrentRegion.Cells.clear
    ListboxToRange ListBox1, rng
End Sub

Sub AddToListBox(lBox As MSForms.ListBox, FileOrFolderPath As String)
    If UCase(isFDU(FileOrFolderPath)) = "F" Then
        lBox.AddItem
        lBox.list(ListBox1.ListCount - 1, 0) = Mid(FileOrFolderPath, InStrRev(FileOrFolderPath, "\") + 1)
        lBox.list(ListBox1.ListCount - 1, 1) = FileOrFolderPath
    Else
        lBox.AddItem
        lBox.list(ListBox1.ListCount - 1, 0) = UCase(Mid(FileOrFolderPath, InStrRev(FileOrFolderPath, "\") + 1)) & "\"
        lBox.list(ListBox1.ListCount - 1, 1) = FileOrFolderPath
    End If
End Sub

Private Function FileRecursive(objFolder As Scripting.Folder, IncludeSubFolders As Boolean)
    Dim objFile As Scripting.file
    Dim objSubFolder As Scripting.Folder
    Dim var As Variant
    Dim element As Variant
    Dim FileFullPath As String
    For Each objFile In objFolder.Files
        FileFullPath = objFile.Path
        'If objFile.DateCreated > Range("afterdate") Then
        var = Split(TextBox2.TEXT, ",")
        On Error Resume Next
        If left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
        For Each element In var
            If InStr(1, FileFullPath, element) > 0 And left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" Then
PASS:
                If ListboxContains(ListBox1, FileFullPath) = False Then
                    AddToListBox ListBox1, FileFullPath
                End If
            End If
        Next
    Next objFile
    If IncludeSubFolders Then
        For Each objSubFolder In objFolder.SubFolders
            Call FileRecursive(objSubFolder, True)
        Next objSubFolder
    End If
End Function

Private Function FolderRecursive(objFolder As Scripting.Folder, IncludeSubFolders As Boolean)
 
    Dim objSubFolder As Scripting.Folder
    Dim var As Variant
    Dim element As Variant
    Dim FileFullPath As String

    For Each objSubFolder In objFolder.SubFolders
        FileFullPath = objSubFolder.Path
        var = Split(TextBox2.TEXT, ",")
        On Error Resume Next
        If left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" And (var(0) = "*" Or var(0) = "") Then GoTo PASS
        For Each element In var
            If InStr(1, FileFullPath, element) > 0 And left(Mid(FileFullPath, InStrRev(FileFullPath, "\") + 1), 1) <> "~" Then
PASS:
                If ListboxContains(ListBox1, FileFullPath & "\") = False Then
                    AddToListBox ListBox1, FileFullPath
                End If
            End If
        Next
    Next objSubFolder
    If IncludeSubFolders Then
        For Each objSubFolder In objFolder.SubFolders
            Call FolderRecursive(objSubFolder, True)
        Next objSubFolder
    End If
End Function

Private Sub TextBox1_Change()
    '    SelectControItemsByFilter ListBox1, TextBox1.TEXT
    LoadListbox
    FilterListboxByColumn ListBox1, TextBox1.TEXT, 0
End Sub

Private Sub UserForm_Activate()
    ResizeUserformToFitControls Me
End Sub

Private Sub UserForm_Initialize()
    '    MakeFormChildOfNothing
    '    UserformOnTop Me
    LoadListbox
End Sub

Private Sub LoadListbox()
    ListBox1.clear
    Dim rng As Range
    Set rng = ThisWorkbook.SHEETS("FileManager_DB").Range("A1")
    If rng.Value = "" Then Exit Sub
    Dim var As Variant
    var = rng.CurrentRegion
    ListBox1.list = var
End Sub

Private Sub CommandButton1_Click()
    Dim element As Variant
    For Each element In ListboxSelectedValues(ListBox1)
        If CStr(element) Like "*.zip" Then
            UnzipToOwnFolder CStr(element), oDeleteExistingFolder, oDeleteZip
        End If
    Next
    MsgBox "Done"
End Sub

Private Sub CommandButton2_Click()
    RemoveItems
End Sub

Private Sub RemoveItems()
    Dim i As Long
    Dim rng As Range

    For i = ListBox1.ListCount - 1 To 0 Step -1
        Dim lookInRng As Range
        Set lookInRng = ThisWorkbook.SHEETS("FileManager_DB").Columns(1)
        If ListBox1.SELECTED(i) = True Then
            If rng Is Nothing Then
                Set rng = lookInRng.Find(ListBox1.list(i), LookIn:=xlValues)
            Else
                Set rng = Union(rng, lookInRng.Find(ListBox1.list(i), LookIn:=xlValues))
            End If
            ListBox1.RemoveItem (i)
        End If
    Next i
    rng.EntireRow.Delete
End Sub

Sub RemoveSelectedFromListbox(lBox As MSForms.ListBox)
    Dim i As Long
    Dim coll As New Collection
    Set coll = ListboxSelectedIndexes(lBox)
    For i = coll.count To 1 Step -1
        lBox.RemoveItem coll(i)
    Next
End Sub

Private Sub info_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Rem file convert
Private Sub oExcelFiles_Click()
    WordOutput.visible = False
    ExcelOutput.visible = True
End Sub

Private Sub oWordFiles_Click()
    WordOutput.visible = True
    WordOutput.left = fFileType.left
    ExcelOutput.visible = False
End Sub

Private Sub Convert_Click()
    Dim element As Variant
    For Each element In ListboxSelectedValues(ListBox1)
        UnzipToOwnFolder CStr(element), oDeleteExistingFolder, oDeleteZip
    Next
End Sub

Sub convertFile(vPath As String)
    If oExcelFiles.Value = True Then
        If vPath Like "*.xl*" Then
            Select Case UCase(whichOption(Me.ExcelOutput, "OptionButton").Caption)
                Case "XLSB"
                    XLS_ConvertFileFormat vPath, xlExcel12, Me.oDelete
                Case "XLSM"
                    XLS_ConvertFileFormat vPath, xlOpenXMLWorkbookMacroEnabled, Me.oDelete
                Case "XLSX"
                    XLS_ConvertFileFormat vPath, xlWorkbookDefault, Me.oDelete
                Case "CSV"
                    XLS_ConvertFileFormat vPath, xlCSV, Me.oDelete
                Case "XLAM"
                    XLS_ConvertFileFormat vPath, xlOpenXMLAddIn, Me.oDelete
                Case "PDF"
                    ExcelToPDF vPath, cSeparateSheets.Value, True
            End Select
        End If
    Else
        If vPath Like "*.doc*" Then
            Select Case whichOption(Me.WordOutput, "OptionButton").Caption
                Case "DOCX"
                    Word_ConvertFileFormat vPath, wdFormatDocumentDefault, Me.oDelete
                Case "TXT"
                    Word_ConvertFileFormat vPath, wdFormatText, Me.oDelete
                Case "DOCM"
                    Word_ConvertFileFormat vPath, wdFormatXMLDocumentMacroEnabled, Me.oDelete
                Case "PDF"
                    Word_ConvertFileFormat vPath, wdFormatPDF, Me.oDelete
            End Select
        End If
    End If
End Sub

