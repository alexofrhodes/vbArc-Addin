Attribute VB_Name = "U_MemoryKnots"

Rem  Procedure : ExportRangeAsImage
Rem  Author    : Daniel Pineault, CARDA Consultants Inc.
Rem  Website   : http://www.cardaconsultants.com
Rem  Purpose   : Capture a picture of a worksheet range and save it to disk
Rem                Returns True if the operation is successful
Rem  Note      : *** Overwrites files, if already exists, without any warning! ***
Rem  Copyright : The following is release as Attribution-ShareAlike 4.0 International
Rem              (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
Rem  Reqrem d Refs: Uses Late Binding, so none required
Rem
Rem  Input Variables:
Rem  ~~~~~~~~~~~~~~~~
Rem  ws            : Worksheet to capture the image of the range from
Rem  rng           : Range to capture an image of
Rem  sPath         : Fully qualified path where to export the image to
Rem  sFileName     : filename to save the image to WITHOUT the extension, just the name
Rem  sImgExtension : The image file extension, commonly: JPG, GIF, PNG, BMP
Rem                    If omitted will be JPG format
Rem
Rem  Usage:
Rem  ~~~~~~
Rem  ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01". "JPG")
Rem  ? ExportRangeAsImage(Sheets("Products"), Range("D5:F23"), "C:\Temp\Charts", "test02")
Rem  ? ExportRangeAsImage(Sheets("Sheet1"), Range("A1"), "C:\Temp\Charts\", "test01", "PNG")
Rem

Rem  Revision History:
Rem  Rev       Date(yyyy/mm/dd)        Description
Rem  **************************************************************************************
Rem  1         2020-04-06              Initial Release
Function ExportRangeAsImage(ws As Worksheet, _
                            rng As Range, _
                            sPath As String, _
                            sFilename As String, _
                            Optional sImgExtension As String = "JPG") As Boolean
    Dim oChart                As ChartObject
    On Error GoTo Error_Handler
    If Right(sPath, 1) <> "\" Then sPath = sPath & "\"
    Application.ScreenUpdating = False
    ws.Activate
    rng.CopyPicture xlScreen, xlPicture
    Set oChart = ws.ChartObjects.Add(0, 0, rng.Width, rng.Height)
    oChart.Activate
    With oChart.Chart
        .Paste
        .Export sPath & sFilename & "." & LCase(sImgExtension), sImgExtension
    End With
    oChart.Delete
    ExportRangeAsImage = True
Error_Handler_Exit:
    On Error Resume Next
    Application.ScreenUpdating = True
    If Not oChart Is Nothing Then Set oChart = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: ExportRangeAsImage" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Sub ListboxSortAZ(myListBox As MSForms.ListBox, Optional resetMacro As String)
    Dim j As Long
    Dim i As Long
    Dim temp As Variant
    If resetMacro <> "" Then
        Run resetMacro, myListBox
    End If
    With myListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) > LCase(.list(i + 1)) Then
                    temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

Sub ListboxSortZA(myListBox As MSForms.ListBox, Optional resetMacro As String)
    Dim j As Long
    Dim i As Long
    Dim temp As Variant
    If resetMacro <> "" Then
        Run resetMacro, myListBox
    End If
    With myListBox
        For j = 0 To .ListCount - 2
            For i = 0 To .ListCount - 2
                If LCase(.list(i)) < LCase(.list(i + 1)) Then
                    temp = .list(i)
                    .list(i) = .list(i + 1)
                    .list(i + 1) = temp
                End If
            Next i
        Next j
    End With
End Sub

Function SelectionValues(link As String)
    Dim c As Range
    If TypeName(Selection) = "Range" _
                             And Selection.Cells.count = 1 Then
        SelectionValues = Selection.Value
    ElseIf TypeName(Selection) = "Range" Then
        For Each c In Selection.SpecialCells(xlCellTypeVisible)
            If Len(c.Value) <> 0 Then
                If SelectionValues = "" Then
                    SelectionValues = c.Value
                Else
                    SelectionValues = SelectionValues & link & c.Value
                End If
            End If
        Next c
    End If
End Function

Function Listbox_Selected(lBox As MSForms.ListBox, Count_Indexes_Values As Integer)
    Dim SelectedIndexes As String
    Dim SelectedValues As String
    Dim SelectedCount As Integer
    Dim i As Long
    With lBox
        For i = 0 To .ListCount - 1
            If .SELECTED(i) Then
                SelectedCount = SelectedCount + 1
                SelectedIndexes = SelectedIndexes & i & ","
                SelectedValues = SelectedValues & .list(i) & ","
            End If
        Next i
    End With
    If SelectedCount = 0 Then
        Listbox_Selected = 0
        Exit Function
    End If
    SelectedIndexes = left(SelectedIndexes, Len(SelectedIndexes) - 1)
    SelectedValues = left(SelectedValues, Len(SelectedValues) - 1)
    Select Case Count_Indexes_Values
        Case Is = 1
            Listbox_Selected = SelectedCount
        Case Is = 2
            Listbox_Selected = SelectedIndexes
        Case Is = 3
            Listbox_Selected = SelectedValues
    End Select
End Function

Sub ListboxClearSelection(lBox As MSForms.ListBox)
    On Error Resume Next
    For i = 0 To lBox.ListCount
        lBox.SELECTED(i) = False
    Next i
End Sub

Sub ListboxSelectValue(lBox As MSForms.ListBox, str As String, Optional clr As Boolean = True)
    '#INCLUDE ListboxClearSelection
    If clr = True Then ListboxClearSelection (lBox)
    For i = 0 To lBox.ListCount - 1
        If lBox.list(i) = str Then
            lBox.SELECTED(i) = True
            Exit Sub
        End If
    Next i
End Sub

Rem  Zip a file or a folder to a zip file/folder using Windows Explorer.
Rem  Default behaviour is similar to right-clicking a file/folder and selecting:
Rem    Send to zip file.
Rem
Rem  Parameters:
Rem    Path:
Rem        Valid (UNC) path to the file or folder to zip.
Rem    Destination:
Rem        (Optional) Valid (UNC) path to file with zip extension or other extension.
Rem    Overwrite:
Rem        (Optional) Leave (default) or overwrite an existing zip file.
Rem        If False, the created zip file will be versioned: Example.zip, Example (2).zip, etc.
Rem        If True, an existing zip file will first be deleted, then recreated.
Rem
Rem    Path and Destination can be relative paths. If so, the current path is used.
Rem
Rem    If success, 0 is returned, and Destination holds the full path of the created zip file.
Rem    If error, error code is returned, and Destination will be zero length string.
Rem
Rem  Early binding requires references to:
Rem
Rem    Shell:
Rem        Microsoft Shell Controls And Automation
Rem
Rem    Scripting.FileSystemObject:
Rem        Microsoft Scripting Runtime
Rem
Rem  2017-10-22. Gustav Brock. Cactus Data ApS, CPH.
Public Function Zip( _
       ByVal Path As String, _
       Optional ByRef Destination As String, _
       Optional ByVal Overwrite As Boolean) _
        As Long
    '#INCLUDE FileExists
    '#INCLUDE FolderExists
    '#INCLUDE OpenTextFile
    '#INCLUDE getFolder
    #If EarlyBinding Then
        Dim FileSystemObject    As Scripting.FileSystemObject
        Dim ShellApplication    As Shell
        Set FileSystemObject = New Scripting.FileSystemObject
        Set ShellApplication = New Shell
    #Else
        Dim FileSystemObject    As Object
        Dim ShellApplication    As Object
        Set FileSystemObject = CreateObject("Scripting.FileSystemObject")
        Set ShellApplication = CreateObject("Shell.Application")
    #End If
    Const ZipExtensionName  As String = "zip"
    Const ZipExtension      As String = "." & ZipExtensionName
    Const ErrorPathFile     As Long = 75
    Const ErrorOther        As Long = -1
    Const ErrorNone         As Long = 0
    Const MaxZipVersion     As Integer = 1000
    Dim ZipHeader           As String
    Dim ZipPath             As String
    Dim ZipName             As String
    Dim ZipFile             As String
    Dim ZipBase             As String
    Dim ZipTemp             As String
    Dim Version             As Integer
    Dim result              As Long
    If FileSystemObject.FileExists(Path) Then
        ZipName = FileSystemObject.GetBaseName(Path) & ZipExtension
        ZipPath = FileSystemObject.GetFile(Path).ParentFolder
    ElseIf FileSystemObject.FolderExists(Path) Then
        ZipName = FileSystemObject.GetBaseName(Path) & ZipExtension
        ZipPath = FileSystemObject.getFolder(Path).ParentFolder
    Else
    End If
    If ZipName = "" Then
        Destination = ""
    Else
        If Destination <> "" Then
            If FileSystemObject.GetExtensionName(Destination) = "" Then
                ZipPath = Destination
            Else
                ZipName = FileSystemObject.GetFileName(Destination)
                ZipPath = FileSystemObject.GetParentFolderName(Destination)
            End If
        Else
        End If
        ZipFile = FileSystemObject.BuildPath(ZipPath, ZipName)
        If FileSystemObject.FileExists(ZipFile) Then
            If Overwrite = True Then
                FileSystemObject.DeleteFile ZipFile, True
            Else
                ZipBase = FileSystemObject.GetBaseName(ZipFile)
                Version = Version + 1
                Do
                    Version = Version + 1
                    ZipFile = FileSystemObject.BuildPath(ZipPath, ZipBase & Format(Version, " \(0\)") & ZipExtension)
                Loop Until FileSystemObject.FileExists(ZipFile) = False Or Version > MaxZipVersion
                If Version > MaxZipVersion Then
                    err.Raise ErrorPathFile, "Zip Create", "File could not be created."
                End If
            End If
        End If
        ZipTemp = FileSystemObject.BuildPath(ZipPath, FileSystemObject.GetBaseName(FileSystemObject.GetTempName()) & ZipExtension)
        ZipHeader = Chr(80) & Chr(75) & Chr(5) & Chr(6) & String(18, vbNullChar)
        With FileSystemObject.OpenTextFile(ZipTemp, ForWriting, True)
            .Write ZipHeader
            .Close
        End With
        ZipTemp = FileSystemObject.GetAbsolutePathName(ZipTemp)
        Path = FileSystemObject.GetAbsolutePathName(Path)
        With ShellApplication
            Debug.Print Timer, "Zipping started . ";
            .Namespace(CVar(ZipTemp)).CopyHere CVar(Path)
            On Error Resume Next
            Do Until .Namespace(CVar(ZipTemp)).items.count = 1
                Application.Wait (Now + TimeValue("0:00:01"))
                Debug.Print ".";
            Loop
            Debug.Print
            On Error GoTo 0
            Debug.Print Timer, "Zipping finished."
        End With
        FileSystemObject.MoveFile ZipTemp, ZipFile
    End If
    Set ShellApplication = Nothing
    Set FileSystemObject = Nothing
    If err.Number <> ErrorNone Then
        Destination = ""
        result = err.Number
    ElseIf Destination = "" Then
        result = ErrorOther
    End If
    Zip = result
End Function

Sub ExportNotes()
    Dim fso As Object
    Set fso = VBA.CreateObject("Scripting.FileSystemObject")
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Dim strPath As String
    strPath = memoPath
    If strPath = "" Then Exit Sub
    On Error Resume Next
    fso.CopyFile Workbooks("MemoryKnots.xlam").FullName, strPath
    Workbooks("MemoryKnots.xlam").SHEETS.Copy
    ActiveWorkbook.SHEETS("MemoryKnots_Settings").Delete
    ActiveWorkbook.SaveAs fileName:=strPath & "MemoryKnots.xlsx", _
                          FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWorkbook.Saved = True
    ActiveWorkbook.Close
    CreateObject("WScript.Shell").PopUp "Successfully exported to " & Chr(10) & strPath, 1
    ActiveWindow.Close
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

Function PickFolder()
    With Application.FileDialog(msoFileDialogFolderPicker)
        .initialFileName = Environ("USERprofile") & "\Desktop\"
        If .Show = -1 Then
            PickFolder = .SelectedItems(1) & "\"
        Else
            Exit Function
        End If
    End With
End Function

Sub ImportNotes()
    '#INCLUDE IsWorkBookOpen
    Dim strADMIN As String
    Dim answer As Integer
    answer = MsgBox("ATTENTION!" & Chr(10) & Chr(10) & _
                    "Present Notebooks will be DELETED and REPLACED from IMPORT file" & Chr(10) & Chr(10) & _
                    "Proceed? (YES) or Cancel import? (NO)", _
                    vbYesNo)
    If answer = vbYes Then
    Else
        Exit Sub
    End If
    Dim AddinWorkbook As Workbook
    Set AddinWorkbook = Workbooks("MemoryKnots.xlam")
    Dim ImportWorkbook As Workbook
    If IsWorkBookOpen("MemoryKnots.xlsx") Then
        Set ImportWorkbook = Workbooks("MemoryKnots.xlsx")
    Else
        If Dir(memoPath & "MemoryKnots.xlsx") > Len(memoPath) Then
            Set ImportWorkbook = Workbooks.Open(memoPath & "MemoryKnots.xlsx")
        Else
            CreateObject("WScript.Shell").PopUp "MemoryKnots.xlsx not found. Use EXPORT first.", 1
            GoTo cleanup
        End If
    End If
    ImportWorkbook.Save
    Application.ScreenUpdating = False
    AddinWorkbook.IsAddin = False
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In AddinWorkbook.Worksheets
        If ws.Name <> "MemoryKnots_Settings" Then ws.Delete
    Next ws
    Dim cell As Range
    For Each ws In ImportWorkbook.Worksheets
        If left(ws.Name, 1) = "o" Then
            For Each cell In ws.Columns(2)
                If cell.Value = "" Then Exit For
                If cell.OFFSET(0, -1) = "" Then
                    cell.OFFSET(0, -1) = Now()
                End If
            Next cell
        End If
    Next ws
    For Each ws In ImportWorkbook.Worksheets
        If ws.Name <> "MemoryKnots_Settings" Then
            ws.Copy After:=AddinWorkbook.SHEETS(AddinWorkbook.SHEETS.count)
        End If
    Next ws
cleanup:
    Application.DisplayAlerts = True
    AddinWorkbook.IsAddin = True
    Set ws = Nothing
    Set ImportWorkbook = Nothing
    Set AddinWorkbook = Nothing
    Application.ScreenUpdating = True
End Sub

Sub testgetfile()
    '#INCLUDE GetFileToImport
    Debug.Print GetFileToImport("xlsx", False)
End Sub

Function GetFileToImport(Optional fileType As String, Optional multiSelect As Boolean) As String
    Dim blArray As Boolean
    Dim strErrMsg As String, strTitle As String
    strTitle = "Import Notebooks"
    If IsMissing(fileType) Then
        Exit Function
    End If
    If strErrMsg = vbNullString Then
        With Application.FileDialog(msoFileDialogFilePicker)
            .initialFileName = "MemoryKnots.xlsx"
            .AllowMultiSelect = multiSelect
            .Filters.clear
            If blArray Then .Filters.Add "File type", "*." & fileType
            .title = strTitle
            If .Show <> 0 Then
                GetFileToImport = .SelectedItems(1)
            End If
        End With
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Function IsWorkBookOpen(Name As String) As Boolean
    Dim xWb As Workbook
    On Error Resume Next
    Set xWb = Application.Workbooks.item(Name)
    IsWorkBookOpen = (Not xWb Is Nothing)
End Function


