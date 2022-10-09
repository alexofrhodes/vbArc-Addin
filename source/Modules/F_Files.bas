Attribute VB_Name = "F_Files"
Option Compare Text
Rem @Folder Files
Rem @Subfolder Files>Convert Declarations
'#Const EarlyBind = True 'Use Early Binding, Req. Reference Library
#Const EarlyBind = False        'Use Late Binding
Enum XlFileFormat
    'Ref: https://msdn.microsoft.com/en-us/vba/excel-vba/articles/xlfileformat-enumeration-excel
    xlAddIn = 18        'Microsoft Excel 97-2003 Add-In *.xla
    xlAddIn8 = 18        'Microsoft Excel 97-2003 Add-In *.xla
    xlCSV = 6        'CSV *.csv
    xlCSVMac = 22        'Macintosh CSV *.csv
    xlCSVMSDOS = 24        'MSDOS CSV *.csv
    xlCSVWindows = 23        'Windows CSV *.csv
    xlCurrentPlatformText = -4158        'Current Platform Text *.txt
    xlDBF2 = 7        'Dbase 2 format *.dbf
    xlDBF3 = 8        'Dbase 3 format *.dbf
    xlDBF4 = 11        'Dbase 4 format *.dbf
    xlDIF = 9        'Data Interchange format *.dif
    xlExcel12 = 50        'Excel Binary Workbook *.xlsb
    xlExcel2 = 16        'Excel version 2.0 (1987) *.xls
    xlExcel2FarEast = 27        'Excel version 2.0 far east (1987) *.xls
    xlExcel3 = 29        'Excel version 3.0 (1990) *.xls
    xlExcel4 = 33        'Excel version 4.0 (1992) *.xls
    xlExcel4Workbook = 35        'Excel version 4.0. Workbook format (1992) *.xlw
    xlExcel5 = 39        'Excel version 5.0 (1994) *.xls
    xlExcel7 = 39        'Excel 95 (version 7.0) *.xls
    xlExcel8 = 56        'Excel 97-2003 Workbook *.xls
    xlExcel9795 = 43        'Excel version 95 and 97 *.xls
    xlHtml = 44        'HTML format *.htm; *.html
    xlIntlAddIn = 26        'International Add-In No file extension
    xlIntlMacro = 25        'International Macro No file extension
    xlOpenDocumentSpreadsheet = 60        'OpenDocument Spreadsheet *.ods
    xlOpenXMLAddIn = 55        'Open XML Add-In *.xlam
    xlOpenXMLStrictWorkbook = 61        '(&;H3D) Strict Open XML file *.xlsx
    xlOpenXMLTemplate = 54        'Open XML Template *.xltx
    xlOpenXMLTemplateMacroEnabled = 53        'Open XML Template Macro Enabled *.xltm
    xlOpenXMLWorkbook = 51        'Open XML Workbook *.xlsx
    xlOpenXMLWorkbookMacroEnabled = 52        'Open XML Workbook Macro Enabled *.xlsm
    xlSYLK = 2        'Symbolic Link format *.slk
    xlTemplate = 17        'Excel Template format *.xlt
    xlTemplate8 = 17        ' Template 8 *.xlt
    xlTextMac = 19        'Macintosh Text *.txt
    xlTextMSDOS = 21        'MSDOS Text *.txt
    xlTextPrinter = 36        'Printer Text *.prn
    xlTextWindows = 20        'Windows Text *.txt
    xlUnicodeText = 42        'Unicode Text No file extension; *.txt
    xlWebArchive = 45        'Web Archive *.mht; *.mhtml
    xlWJ2WD1 = 14        'Japanese 1-2-3 *.wj2
    xlWJ3 = 40        'Japanese 1-2-3 *.wj3
    xlWJ3FJ3 = 41        'Japanese 1-2-3 format *.wj3
    xlWK1 = 5        'Lotus 1-2-3 format *.wk1
    xlWK1ALL = 31        'Lotus 1-2-3 format *.wk1
    xlWK1FMT = 30        'Lotus 1-2-3 format *.wk1
    xlWK3 = 15        'Lotus 1-2-3 format *.wk3
    xlWK3FM3 = 32        'Lotus 1-2-3 format *.wk3
    xlWK4 = 38        'Lotus 1-2-3 format *.wk4
    xlWKS = 4        'Lotus 1-2-3 format *.wks
    xlWorkbookDefault = 51        'Workbook default *.xlsx
    xlWorkbookNormal = -4143        'Workbook normal *.xls
    xlWorks2FarEast = 28        'Microsoft Works 2.0 far east format *.wks
    xlWQ1 = 34        'Quattro Pro format *.wq1
    xlXMLSpreadsheet = 46        'XML Spreadsheet *.xml
    TypePDF = 47
End Enum

Enum WdSaveFormat
    'Ref: https://msdn.microsoft.com/en-us/vba/word-vba/articles/wdsaveformat-enumeration-word
    wdFormatDocument = 0        'Microsoft Office Word 97 - 2003 binary file format.
    wdFormatDOSText = 4        'Microsoft DOS text format.  *.txt
    wdFormatDOSTextLineBreaks = 5        'Microsoft DOS text with line breaks preserved.  *.txt
    wdFormatEncodedText = 7        'Encoded text format.  *.txt
    wdFormatFilteredHTML = 10        'Filtered HTML format.
    wdFormatFlatXML = 19        'Open XML file format saved as a single XML file.
    '    wdFormatFlatXML = 20                                                    'Open XML file format with macros enabled saved as a single XML file.
    wdFormatFlatXMLTemplate = 21        'Open XML template format saved as a XML single file.
    wdFormatFlatXMLTemplateMacroEnabled = 22        'Open XML template format with macros enabled saved as a single XML file.
    wdFormatOpenDocumentText = 23        'OpenDocument Text format. *.odt
    wdFormatHTML = 8        'Standard HTML format. *.html
    wdFormatRTF = 6        'Rich text format (RTF). *.rtf
    wdFormatStrictOpenXMLDocument = 24        'Strict Open XML document format.
    wdFormatTemplate = 1        'Word template format.
    wdFormatText = 2        'Microsoft Windows text format. *.txt
    wdFormatTextLineBreaks = 3        'Windows text format with line breaks preserved. *.txt
    wdFormatUnicodeText = 7        'Unicode text format. *.txt
    wdFormatWebArchive = 9        'Web archive format.
    wdFormatXML = 11        'Extensible Markup Language (XML) format. *.xml
    wdFormatDocument97 = 0        'Microsoft Word 97 document format. *.doc
    wdFormatDocumentDefault = 16        'Word default document file format. For Word, this is the DOCX format. *.docx
    wdFormatPDF = 17        'PDF format. *.pdf
    wdFormatTemplate97 = 1        'Word 97 template format.
    wdFormatXMLDocument = 12        'XML document format.
    wdFormatXMLDocumentMacroEnabled = 13        'XML document format with macros enabled.
    wdFormatXMLTemplate = 14        'XML template format.
    wdFormatXMLTemplateMacroEnabled = 15        'XML template format with macros enabled.
    wdFormatXPS = 18        'XPS format. *.xps
End Enum

Rem @Subfolder Files>Recycle Declarations
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias _
"SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare PtrSafe Function PathIsNetworkPath Lib "shlwapi.dll" _
Alias "PathIsNetworkPathA" ( _
ByVal pszPath As String) As Long
Private Declare PtrSafe Function GetSystemDirectory Lib "kernel32" _
Alias "GetSystemDirectoryA" ( _
ByVal lpBuffer As String, _
ByVal nSize As Long) As Long
Private Declare PtrSafe Function SHEmptyRecycleBin _
Lib "Shell32" Alias "SHEmptyRecycleBinA" _
(ByVal hWnd As Long, _
ByVal pszRootPath As String, _
ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function PathIsDirectory Lib "shlwapi" (ByVal pszPath As String) As Long
Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10
Private Const MAX_PATH As Long = 260
Private Type SHFILEOPSTRUCT
    hWnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As Long
    lpszProgressTitle As String
End Type

Rem @Subfolder Files>Recycle
Public Function RecycleFile(fileName As String) As Boolean
    '#INCLUDE Recycle
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    If Dir(fileName, vbNormal) = vbNullString Then
        RecycleFile = True
        Exit Function
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = fileName
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        RecycleFile = True
    Else
        RecycleFile = False
    End If
End Function

Public Function Recycle(filespec As String, Optional ErrText As String) As Boolean
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    Dim sFileSpec As String
    ErrText = vbNullString
    sFileSpec = filespec
    If InStr(1, filespec, ":", vbBinaryCompare) = 0 Then
        ErrText = "'" & filespec & "' is not a fully qualified name on the local machine"
        Recycle = False
        Exit Function
    End If
    If Dir(filespec, vbDirectory) = vbNullString Then
        ErrText = "'" & filespec & "' does not exist"
        Recycle = False
        Exit Function
    End If
    If Right(sFileSpec, 1) = "\" Then
        sFileSpec = left(sFileSpec, Len(sFileSpec) - 1)
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileSpec
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        Recycle = True
    Else
        Recycle = False
    End If
End Function

Public Function RecycleSafe(filespec As String, Optional ByRef ErrText As String) As Boolean
    Dim ThisWorkbookFullName As String
    Dim ThisWorkbookPath As String
    Dim WindowsFolder As String
    Dim SystemFolder As String
    Dim ProgramFiles As String
    Dim MyDocuments As String
    Dim Desktop As String
    Dim ApplicationPath As String
    Dim pos As Long
    Dim ShellObj As Object
    Dim sFileSpec As String
    Dim SHFileOp As SHFILEOPSTRUCT
    Dim Res As Long
    Dim FileNum As Integer
    sFileSpec = filespec
    If InStr(1, filespec, ":", vbBinaryCompare) = 0 Then
        RecycleSafe = False
        ErrText = "'" & filespec & "' is not a fully qualified name on the local machine"
        Exit Function
    End If
    If Dir(filespec, vbDirectory) = vbNullString Then
        RecycleSafe = False
        ErrText = "'" & filespec & "' does not exist"
        Exit Function
    End If
    If Right(sFileSpec, 1) = "\" Then
        sFileSpec = left(sFileSpec, Len(sFileSpec) - 1)
    End If
    ThisWorkbookFullName = ThisWorkbook.FullName
    ThisWorkbookPath = ThisWorkbook.Path
    SystemFolder = String$(MAX_PATH, vbNullChar)
    GetSystemDirectory SystemFolder, Len(SystemFolder)
    SystemFolder = left(SystemFolder, InStr(1, SystemFolder, vbNullChar, vbBinaryCompare) - 1)
    pos = InStrRev(SystemFolder, "\")
    If pos > 0 Then
        WindowsFolder = left(SystemFolder, pos - 1)
    End If
    pos = InStr(1, Application.Path, "\", vbBinaryCompare)
    pos = InStr(pos + 1, Application.Path, "\", vbBinaryCompare)
    ProgramFiles = left(Application.Path, pos - 1)
    ApplicationPath = Application.Path
    On Error Resume Next
    err.clear
    Set ShellObj = CreateObject("WScript.Shell")
    If ShellObj Is Nothing Then
        RecycleSafe = False
        ErrText = "Error Creating WScript.Shell. " & CStr(err.Number) & ": " & err.Description
        Exit Function
    End If
    MyDocuments = ShellObj.SpecialFolders("MyDocuments")
    Desktop = ShellObj.SpecialFolders("Desktop")
    Set ShellObj = Nothing
    If (sFileSpec Like "?*:") Or (sFileSpec Like "?*:\") Then
        RecycleSafe = False
        ErrText = "File Specification is a root directory."
        Exit Function
    End If
    If (InStr(1, sFileSpec, "*", vbBinaryCompare) > 0) Or (InStr(1, sFileSpec, "?", vbBinaryCompare) > 0) Then
        RecycleSafe = False
        ErrText = "File specification contains wildcard characters"
        Exit Function
    End If
    If StrComp(sFileSpec, ThisWorkbookFullName, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the same as this workbook."
        Exit Function
    End If
    If StrComp(sFileSpec, ThisWorkbookPath, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is this workbook's path"
        Exit Function
    End If
    If StrComp(ThisWorkbook.FullName, sFileSpec, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is this workbook."
        Exit Function
    End If
    If StrComp(sFileSpec, SystemFolder, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the System Folder"
        Exit Function
    End If
    If StrComp(sFileSpec, WindowsFolder, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is the Windows folder"
        Exit Function
    End If
    If StrComp(sFileSpec, Application.Path, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is Application Path"
        Exit Function
    End If
    If StrComp(sFileSpec, MyDocuments, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is MyDocuments"
        Exit Function
    End If
    If StrComp(sFileSpec, Desktop, vbTextCompare) = 0 Then
        RecycleSafe = False
        ErrText = "File specification is Desktop"
        Exit Function
    End If
    If (GetAttr(sFileSpec) And vbSystem) <> 0 Then
        RecycleSafe = False
        ErrText = "File specification is a System entity"
        Exit Function
    End If
    If PathIsDirectory(sFileSpec) = 0 Then
        FileNum = FreeFile()
        On Error Resume Next
        err.clear
        Open sFileSpec For Input Lock Read As #FileNum
        If err.Number <> 0 Then
            Close #FileNum
            RecycleSafe = False
            ErrText = "File in use: " & CStr(err.Number) & "  " & err.Description
            Exit Function
        End If
        Close #FileNum
    End If
    With SHFileOp
        .wFunc = FO_DELETE
        .pFrom = sFileSpec
        .fFlags = FOF_ALLOWUNDO
        .fFlags = FOF_ALLOWUNDO + FOF_NOCONFIRMATION
    End With
    Res = SHFileOperation(SHFileOp)
    If Res = 0 Then
        RecycleSafe = True
    Else
        RecycleSafe = False
    End If
End Function

Public Function EmptyRecycleBin(Optional DriveRoot As String = vbNullString) As Boolean
    '#INCLUDE Recycle
    Const SHERB_NOCONFIRMATION = &H1
    Const SHERB_NOPROGRESSUI = &H2
    Const SHERB_NOSOUND = &H4
    Dim Res As Long
    If DriveRoot <> vbNullString Then
        If PathIsNetworkPath(DriveRoot) <> 0 Then
            MsgBox "You can't empty the Recycle Bin of a network drive."
            Exit Function
        End If
    End If
    Res = SHEmptyRecycleBin(hWnd:=0&, _
                            pszRootPath:=DriveRoot, _
                            dwFlags:=SHERB_NOCONFIRMATION + _
                                      SHERB_NOPROGRESSUI + _
                                      SHERB_NOSOUND)
    If Res = 0 Then
        EmptyRecycleBin = True
    Else
        EmptyRecycleBin = False
    End If
End Function

Rem End of Recycle
Rem @Subfolder Files>ConvertWord
'---------------------------------------------------------------------------------------
' Procedure : Word_ConvertFileFormat
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Converts a Word compatible file format to another format
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sOrigFile     : String - Original file path, name and extension to be converted
' lNewFileFormat: New File format to save the original file as
' bDelOrigFile  : True/False - Should the original file be deleted after the conversion
'
' Usage:
' ~~~~~~
' Convert a doc file into a docx file but retain the original copy
'   Call Word_ConvertFileFormat("C:\Users\Daniel\Documents\Resume.doc", wdFormatPDF)
' Convert a doc file into a docx file and delete the original doc once converted
'   Call Word_ConvertFileFormat("C:\Users\Daniel\Documents\Resume.doc", wdFormatPDF, True)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2018-02-27              Initial Release
'---------------------------------------------------------------------------------------
Function Word_ConvertFileFormat(ByVal sOrigFile As String, _
                                Optional lNewFileFormat As WdSaveFormat = wdFormatDocumentDefault, _
                                Optional bDelOrigFile As Boolean = False) As Boolean
    '#INCLUDE Recycle
    '#INCLUDE XLS_ConvertFileFormat
    #If EarlyBind = True Then
        Dim oWord             As Word.Application
        Dim oDoc              As Word.document
    #Else
        Dim oWord             As Object
        Dim oDoc              As Object
    #End If
    Dim bWordOpened           As Boolean
    Dim sOrigFileExt          As String
    Dim sNewFileExt           As String
    Select Case lNewFileFormat
        Case wdFormatDocument
            sNewFileExt = "."
        Case wdFormatDOSText, wdFormatDOSTextLineBreaks, wdFormatEncodedText, wdFormatOpenDocumentText, wdFormatText, wdFormatTextLineBreaks, wdFormatUnicodeText
            sNewFileExt = ".txt"
        Case wdFormatFilteredHTML, wdFormatHTML
            sNewFileExt = ".html"
        Case wdFormatFlatXML, wdFormatXML, wdFormatXMLDocument
            sNewFileExt = ".xml"
        Case wdFormatFlatXMLTemplate
            sNewFileExt = "."
        Case wdFormatFlatXMLTemplateMacroEnabled
            sNewFileExt = "."
        Case wdFormatRTF
            sNewFileExt = ".rtf"
        Case wdFormatStrictOpenXMLDocument
            sNewFileExt = "."
        Case wdFormatTemplate
            sNewFileExt = "."
        Case wdFormatWebArchive
            sNewFileExt = "."
        Case wdFormatDocument97
            sNewFileExt = ".doc"
        Case wdFormatDocumentDefault
            sNewFileExt = ".docx"
        Case wdFormatPDF
            sNewFileExt = ".pdf"
        Case wdFormatTemplate97
            sNewFileExt = "."
        Case wdFormatXMLDocumentMacroEnabled
            sNewFileExt = ".docm"
        Case wdFormatXMLTemplate
            sNewFileExt = ".doct"
        Case wdFormatXMLTemplateMacroEnabled
            sNewFileExt = "."
        Case wdFormatXPS
            sNewFileExt = ".xps"
    End Select
    sOrigFileExt = "." & Right(sOrigFile, Len(sOrigFile) - InStrRev(sOrigFile, "."))
    On Error Resume Next
    Set oWord = GetObject(, "Word.Application")
    If err.Number <> 0 Then
        err.clear
        On Error GoTo Error_Handler
        Set oWord = CreateObject("Word.Application")
    Else
        bWordOpened = True
    End If
    On Error GoTo Error_Handler
    oWord.visible = False
    Set oDoc = oWord.Documents.Open(sOrigFile)
    oDoc.SaveAs2 Replace(sOrigFile, sOrigFileExt, sNewFileExt), lNewFileFormat
    Word_ConvertFileFormat = True
    oDoc.Close False
    If bWordOpened = False Then
        oWord.Quit
    Else
        oWord.visible = True
    End If
    If bDelOrigFile = True Then Recycle (sOrigFile)
Error_Handler_Exit:
    On Error Resume Next
    Set oDoc = Nothing
    Set oWord = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: XLS_ConvertFileFormat" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    oWord.visible = True
    Resume Error_Handler_Exit
End Function

Rem @Subfolder Files>ConvertExcel
'---------------------------------------------------------------------------------------
' Procedure : XLS_ConvertFileFormat
' Author    : Daniel Pineault, CARDA Consultants Inc.
' Website   : http://www.cardaconsultants.com
' Purpose   : Converts an Excel compatible file format to another format
' Copyright : The following is release as Attribution-ShareAlike 4.0 International
'             (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
' Req'd Refs: Uses Late Binding, so none required
'
' Input Variables:
' ~~~~~~~~~~~~~~~~
' sOrigFile     : String - Original file path, name and extension to be converted
' lNewFileFormat: New File format to save the original file as
' bDelOrigFile  : True/False - Should the original file be deleted after the conversion
'
' Usage:
' ~~~~~~
' Convert an xls file into a txt file and delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.xls", xlTextWindows)
' Convert an xls file into a xlsx file and NOT delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.xls",, False)
' Convert a csv file into a xlsx file and delete the xls once completed
'   Call XLS_ConvertFileFormat("C:TempTest.csv", xlWorkbookDefault, True)
'
' Revision History:
' Rev       Date(yyyy/mm/dd)        Description
' **************************************************************************************
' 1         2018-02-27              Initial Release
' 2         2020-12-31              Fixed typo xlDBF24 -> xlDBF4
'---------------------------------------------------------------------------------------
Function XLS_ConvertFileFormat(ByVal sOrigFile As String, _
                               Optional lNewFileFormat As XlFileFormat = xlOpenXMLWorkbook, _
                               Optional bDelOrigFile As Boolean = False) As Boolean
    '#INCLUDE Recycle
    #If EarlyBind = True Then
        Dim oExcel            As Excel.Application
        Dim oExcelWrkBk       As Excel.Workbook
    #Else
        Dim oExcel            As Object
        Dim oExcelWrkBk       As Object
    #End If
    Dim bExcelOpened          As Boolean
    Dim sOrigFileExt          As String
    Dim sNewXLSFileExt        As String
    Select Case lNewFileFormat
        Case xlAddIn, xlAddIn8
            sNewFileExt = ".xla"
        Case xlCSV, xlCSVMac, xlCSVMSDOS, xlCSVWindows
            sNewFileExt = ".csv"
        Case xlCurrentPlatformText, xlTextMac, xlTextMSDOS, xlTextWindows, xlUnicodeText
            sNewFileExt = ".txt"
        Case xlDBF2, xlDBF3, xlDBF4
            sNewFileExt = ".dbf"
        Case xlDIF
            sNewFileExt = ".dif"
        Case xlExcel12 = 50
            sNewFileExt = ".xlsb"
        Case xlExcel2, xlExcel2FarEast, xlExcel3, xlExcel4, xlExcel5, xlExcel7, _
             xlExcel8, xlExcel9795, xlWorkbookNormal
            sNewFileExt = ".xls"
        Case xlExcel4Workbook = 35
            sNewFileExt = ".xlw"
        Case xlHtml = 44
            sNewFileExt = ".html"
        Case xlIntlAddIn, xlIntlMacro
            sNewFileExt = ""
        Case xlOpenDocumentSpreadsheet
            sNewFileExt = ".ods"
        Case xlOpenXMLAddIn
            sNewFileExt = ".xlam"
        Case xlOpenXMLStrictWorkbook, xlOpenXMLWorkbook, xlWorkbookDefault = 51
            sNewFileExt = ".xlsx"
        Case xlOpenXMLTemplate
            sNewFileExt = ".xltx"
        Case xlOpenXMLTemplateMacroEnabled
            sNewFileExt = ".xltm"
        Case xlOpenXMLWorkbookMacroEnabled
            sNewFileExt = ".xlsm"
        Case xlSYLK
            sNewFileExt = ".slk"
        Case xlTemplate, xlTemplate8
            sNewFileExt = ".xlt"
        Case xlTextPrinter
            sNewFileExt = ".prn"
        Case xlWebArchive
            sNewFileExt = ".mhtml"
        Case xlWJ2WD1
            sNewFileExt = ".wj2"
        Case xlWJ3, xlWJ3FJ3
            sNewFileExt = ".wj3"
        Case xlWK1, xlWK1ALL, xlWK1FMT
            sNewFileExt = ".wk1"
        Case xlWK3, xlWK3FM3
            sNewFileExt = ".wk3"
        Case xlWK4
            sNewFileExt = ".wk4"
        Case xlWKS, xlWorks2FarEast
            sNewFileExt = ".wks"
        Case xlWQ1
            sNewFileExt = ".wq1"
        Case xlXMLSpreadsheet
            sNewFileExt = ".xml"
        Case TypePDF
            sNewFileExt = ".pdf"
    End Select
    sOrigFileExt = "." & Right(sOrigFile, Len(sOrigFile) - InStrRev(sOrigFile, "."))
    On Error Resume Next
    Set oExcel = GetObject(, "Excel.Application")
    If err.Number <> 0 Then
        err.clear
        On Error GoTo Error_Handler
        Set oExcel = CreateObject("Excel.Application")
    Else
        bExcelOpened = True
    End If
    On Error GoTo Error_Handler
    oExcel.ScreenUpdating = False
    oExcel.visible = False
    Set oExcelWrkBk = oExcel.Workbooks.Open(sOrigFile)
    oExcelWrkBk.SaveAs Replace(sOrigFile, sOrigFileExt, sNewFileExt), lNewFileFormat, , , , False
    XLS_ConvertFileFormat = True
    oExcelWrkBk.Close False
    If bExcelOpened = False Then
        oExcel.Quit
    Else
        oExcel.ScreenUpdating = True
        oExcel.visible = True
    End If
    If bDelOrigFile = True Then Recycle (sOrigFile)
Error_Handler_Exit:
    On Error Resume Next
    Set oExcelWrkBk = Nothing
    Set oExcel = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: XLS_ConvertFileFormat" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    oExcel.ScreenUpdating = True
    oExcel.visible = True
    Resume Error_Handler_Exit
End Function

Sub ExcelToPDF(FileFullPath As String, SeparateSheets As Boolean, CloseFile As Boolean)
    Dim wb As Workbook
    Set wb = Workbooks.Open(FileFullPath)
    Dim ws As Worksheet
    If SeparateSheets = False Then
        wb.ExportAsFixedFormat xlTypePDF, _
                               VBA.Replace(FileFullPath, Right(FileFullPath, Len(FileFullPath) - InStrRev(FileFullPath, ".") + 1), ".pdf")
        If CloseFile = True Then wb.Close False
    Else
        For Each ws In wb
            ws.ExportAsFixedFormat xlTypePDF, wb.Path & "\" & ws.Name & ".pdf"
        Next ws
    End If
    MsgBox "Process Completed"
End Sub

Rem @Subfolder Files>Unsorted
Function IsFileFolderURL(Path) As String
    '#INCLUDE HttpExists
    '#INCLUDE FileExists
    '#INCLUDE FolderExists
    Dim retval As String
    retval = "I"
    If (retval = "I") And FileExists(Path) Then retval = "F"
    If (retval = "I") And FolderExists(Path) Then retval = "D"
    If (retval = "I") And HttpExists(Path) Then retval = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    IsFileFolderURL = retval
End Function

Rem Folders
Public Function SelectFolder(Optional initFolder As String) As String
    '#INCLUDE FolderExists
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .title = "Select a folder"
        If FolderExists(initFolder) Then .initialFileName = initFolder
        .Show
        If .SelectedItems.count > 0 Then
            SelectFolder = .SelectedItems.item(1)
        Else
        End If
    End With
End Function

Sub FoldersCreate(FolderPath As String)
    '#INCLUDE FolderExists
    On Error Resume Next
    Dim individualFolders() As String
    Dim tempFolderPath As String
    Dim arrayElement As Variant
    individualFolders = Split(FolderPath, "\")
    For Each arrayElement In individualFolders
        tempFolderPath = tempFolderPath & arrayElement & "\"
        If FolderExists(tempFolderPath) = False Then
            MkDir tempFolderPath
        End If
    Next arrayElement
End Sub

Rem Files
Function GetFilePartPath(fileNameWithExtension, Optional IncludeSlash As Boolean) As String
    GetFilePartPath = left(fileNameWithExtension, InStrRev(fileNameWithExtension, "\") - 1 - IncludeSlash)
End Function

Public Function FFileDialog(Optional ByRef lDialogType As MsoFileDialogType = msoFileDialogFilePicker, _
                            Optional sTitle As String = "", _
                            Optional sInitFileName = "", _
                            Optional bMultiSelect As Boolean = False, _
                            Optional sFilter As String = "All Files,*.*") As String
    Dim out As String
    On Error GoTo Error_Handler
    Dim oFd                   As Object
    Dim vItems                As Variant
    Dim vFilter               As Variant
    Const msoFileDialogViewDetails = 2
    Set oFd = Application.FileDialog(lDialogType)
    With oFd
        If sTitle = "" Then
            Select Case lDialogType
                Case msoFileDialogFilePicker
                    .title = "Browse for File"
                Case msoFileDialogFolderPicker
                    .title = "Browse for Folder"
            End Select
        Else
            .title = sTitle
        End If
        If sInitFileName <> "" Then .initialFileName = sInitFileName
        .AllowMultiSelect = bMultiSelect
        .InitialView = msoFileDialogViewDetails
        If lDialogType <> msoFileDialogFolderPicker Then
            Call .Filters.clear
            For Each vFilter In Split(sFilter, "~")
                Call .Filters.Add(Split(vFilter, ",")(0), Split(vFilter, ",")(1))
            Next vFilter
        End If
        If .Show = True Then
            For Each vItems In .SelectedItems
                If out = "" Then
                    out = vItems
                Else
                    out = out & "," & vItems
                End If
            Next
        End If
    End With
    FFileDialog = out
Error_Handler_Exit:
    On Error Resume Next
    If Not oFd Is Nothing Then Set oFd = Nothing
    Exit Function
Error_Handler:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: fFileDialog" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    Resume Error_Handler_Exit
End Function

Function splitLineBreaks(ByVal str As String) As String
    str = Replace(str, vbCrLf, vbCr)
    str = Replace(str, vbLf, vbCr)
    splitLineBreaks = Split(str, vbCr)
End Function

Public Sub LoopAllFilesAndFolders(FolderPath As String)
    '#INCLUDE getFolder
    If Right(FolderPath, 1) <> "\" Then FolderPath = FolderPath & "\"
    Dim objFSO As Scripting.FileSystemObject
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFile As Scripting.file
    Dim objFolder As Scripting.Folder
    Set objTopFolder = objFSO.getFolder(FolderPath)
    For Each objFile In objFolder.Files
        Rem code here
    Next
    Dim objSubFolder As Scripting.Folder
    For Each objSubFolder In objFolder.SubFolders
        LoopAllFilesAndFolders objSubFolder.Path
    Next
End Sub

Function isFDU(Path) As String
    '#INCLUDE HttpExists
    '#INCLUDE FileExists
    '#INCLUDE FolderExists
    Dim retval
    retval = "I"
    If (retval = "I") And FileExists(Path) Then retval = "F"
    If (retval = "I") And FolderExists(Path) Then retval = "D"
    If (retval = "I") And HttpExists(Path) Then retval = "U"
    ' I => Invalid | F => File | D => Directory | U => Valid Url
    isFDU = retval
End Function

Public Function FileExists(ByVal fileName As String) As Boolean
    If InStr(1, fileName, "\") = 0 Then Exit Function
    If Right(fileName, 1) = "\" Then fileName = left(fileName, Len(fileName) - 1)
    FileExists = (Dir(fileName, vbArchive + vbHidden + vbReadOnly + vbSystem) <> "")
End Function

'Function FileExists(ByVal strFile As String, Optional bFindFolders As Boolean) As Boolean
'    'Purpose:   Return True if the file exists, even if it is hidden.
'    'Arguments: strFile: File name to look for. Current directory searched if no path included.
'    '           bFindFolders. If strFile is a folder, FileExists() returns False unless this argument is True.
'    'Note:      Does not look inside subdirectories for the file.
'    'Author:    Allen Browne. http://allenbrowne.com June, 2006.
'    Dim lngAttributes As Long
'
'    'Include read-only files, hidden files, system files.
'    lngAttributes = (vbReadOnly Or vbHidden Or vbSystem)
'    If bFindFolders Then
'        lngAttributes = (lngAttributes Or vbDirectory) 'Include folders as well.
'    Else
'        'Strip any trailing slash, so Dir does not look inside the folder.
'        Do While Right$(strFile, 1) = "\"
'            strFile = left$(strFile, Len(strFile) - 1)
'        Loop
'    End If
'    'If Dir() returns something, the file exists.
'    On Error Resume Next
'    FileExists = (Len(Dir(strFile, lngAttributes)) > 0)
'End Function
Function TrailingSlash(varIn As Variant) As String
    '#INCLUDE FileExists
    If Len(varIn) > 0 Then
        If Right(varIn, 1) = "\" Then
            TrailingSlash = varIn
        Else
            TrailingSlash = varIn & "\"
        End If
    End If
End Function

Function GetInputRange(rInput As Excel.Range, _
                       sPrompt As String, _
                       sTitle As String, _
                       Optional ByVal sDefault As String, _
                       Optional ByVal bActivate As Boolean, _
                       Optional X, _
                       Optional Y) As Boolean
    Dim bGotRng As Boolean
    Dim bEvents As Boolean
    Dim nAttempt As Long
    Dim sAddr As String
    Dim vReturn
    On Error Resume Next
    If Len(sDefault) = 0 Then
        If TypeName(Application.Selection) = "Range" Then
            sDefault = "=" & Application.Selection.Address
            If Len(sDefault) > 240 Then
                sDefault = "=" & Application.ActiveCell.Address
            End If
        ElseIf TypeName(Application.ActiveSheet) = "Chart" Then
            sDefault = " first select a Worksheet"
        Else
            sDefault = " Select Cell(s) or type address"
        End If
    End If
    Set rInput = Nothing
    For nAttempt = 1 To 3
        vReturn = False
        vReturn = Application.InputBox(sPrompt, sTitle, sDefault, X, Y, Type:=0)
        If False = vReturn Or Len(vReturn) = 0 Then
            Exit For
        Else
            sAddr = vReturn
            If left$(sAddr, 1) = "=" Then sAddr = Mid$(sAddr, 2, 256)
            If left$(sAddr, 1) = Chr(34) Then sAddr = Mid$(sAddr, 2, 255)
            If Right$(sAddr, 1) = Chr(34) Then sAddr = left$(sAddr, Len(sAddr) - 1)
            Set rInput = Application.Range(sAddr)
            If rInput Is Nothing Then
                sAddr = Application.ConvertFormula(sAddr, xlR1C1, xlA1)
                Set rInput = Application.Range(sAddr)
                bGotRng = Not rInput Is Nothing
            Else
                bGotRng = True
            End If
        End If
        If bGotRng Then
            If bActivate Then
                On Error GoTo errH
                bEvents = Application.EnableEvents
                Application.EnableEvents = False
                If Not Application.ActiveWorkbook Is rInput.parent.parent Then
                    rInput.parent.parent.Activate
                End If
                If Not Application.ActiveSheet Is rInput.parent Then
                    rInput.parent.Activate
                End If
                rInput.Select
            End If
            Exit For
        ElseIf nAttempt < 3 Then
            If MsgBox("Invalid reference, do you want to try again ?", _
                      vbOKCancel, sTitle) <> vbOK Then
                Exit For
            End If
        End If
    Next
cleanup:
    On Error Resume Next
    If bEvents Then
        Application.EnableEvents = True
    End If
    GetInputRange = bGotRng
    Exit Function
errH:
    Set rInput = Nothing
    bGotRng = False
    Resume cleanup
End Function

Function FolderExists(ByVal strPath As String) As Boolean
    On Error Resume Next
    FolderExists = ((GetAttr(strPath) And vbDirectory) = vbDirectory)
End Function

Sub ExtractCodeFromFiles()
    '#INCLUDE OleVbaRun
    '#INCLUDE GetFilePath
    Dim v
    v = GetFilePath(Array("xl*"), True)
    Dim el
    Dim coll As New Collection
    For Each el In v
        coll.Add el
    Next
    OleVbaRun coll
End Sub

Sub OleVbaRun(FilePathCollection As Collection)
    '#INCLUDE FollowLink
    '#INCLUDE FoldersCreate
    '#INCLUDE oleVBA
    '#INCLUDE getFilePartName
    '#INCLUDE TxtOverwrite
    Dim MainPath As String
    MainPath = Environ$("USERPROFILE") & "\Documents\vbArc\oleVba\"
    FoldersCreate MainPath
    Dim OutPath As String
    Dim fileName As String
    Dim element As Variant
    Dim output As String
    For Each element In FilePathCollection
        fileName = getFilePartName(CStr(element))
        output = oleVBA(element)
        OutPath = MainPath & "\" & fileName & ".txt"
        TxtOverwrite OutPath, output
    Next
    If FilePathCollection.count > 0 Then FollowLink MainPath
End Sub

Function oleVBA(Path As Variant) As String
    '#INCLUDE ShellText
    Dim Q As String
    Q = """"
    oleVBA = ShellText("cmd.exe /c olevba " & Q & Path & Q)
End Function

Function ShellText(sCmd As String) As String
    Dim oShell   As New WshShell        'requires ref to Windows Script Host Object Model
    ShellText = oShell.Exec(sCmd).StdOut.ReadAll
End Function

Function getFilePartName(fileNameWithExtension As String, Optional IncludeExtension As Boolean) As String
    If InStr(1, fileNameWithExtension, "\") > 0 Then
        getFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "\"))
    ElseIf InStr(1, fileNameWithExtension, "/") > 0 Then
        getFilePartName = Right(fileNameWithExtension, Len(fileNameWithExtension) - InStrRev(fileNameWithExtension, "/"))
    Else
        getFilePartName = fileNameWithExtension
    End If
    If IncludeExtension = False Then getFilePartName = left(getFilePartName, InStr(1, getFilePartName, ".") - 1)
End Function

Function PickExcelFile()
    Dim strFile As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.clear
        .Filters.Add "Excel Files", "*.xl*", 1
        .title = "Choose an Excel file"
        .AllowMultiSelect = False
        .initialFileName = Environ("USERprofile") & "\Desktop\"
        If .Show = True Then
            strFile = .SelectedItems(1)
            PickExcelFile = strFile
        End If
    End With
End Function

Function LoopThroughFiles(Folder, criteria) As Collection
    If Right(Folder, 1) <> "\" Then Folder = Folder & "\"
    Dim out As Collection: Set out = New Collection
    Dim strFile As String
    strFile = Dir(Folder & criteria)
    Do While Len(strFile) > 0
        out.Add strFile
        strFile = Dir
    Loop
    Set LoopThroughFiles = out
End Function

Function FileCreated(FilePath) As Date
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FilePath)
    FileCreated = f.datecreated
End Function

Function FileLastAccessed(FilePath) As Date
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FilePath)
    FileLastAccessed = f.DateLastAccessed
End Function

Function FileLastModified(FilePath) As Date
    Dim fs, f
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(FilePath)
    FileLastModified = f.DateLastModified
End Function

Public Function GetFilePath(Optional fileType As Variant, Optional multiSelect As Boolean) As Variant
    Dim blArray As Boolean
    Dim i As Long
    Dim strErrMsg As String, strTitle As String
    Dim varItem As Variant
    If Not IsMissing(fileType) Then
        blArray = IsArray(fileType)
        If Not blArray Then strErrMsg = "Please pass an array in the first parameter of this function!"
    End If
    If strErrMsg = vbNullString Then
        If multiSelect Then strTitle = "Choose one or more files" Else strTitle = "Choose file"
        With Application.FileDialog(msoFileDialogFilePicker)
            .initialFileName = Environ("USERprofile") & "\Desktop\"
            .AllowMultiSelect = multiSelect
            .Filters.clear
            If blArray Then .Filters.Add "File type", Replace("*." & Join(fileType, ", *."), "..", ".")
            .title = strTitle
            If .Show <> 0 Then
                ReDim arrResults(1 To .SelectedItems.count) As Variant
                If blArray Then
                    For Each varItem In .SelectedItems
                        i = i + 1
                        arrResults(i) = varItem
                    Next varItem
                Else
                    arrResults(1) = .SelectedItems(1)
                End If
                GetFilePath = arrResults
            End If
        End With
    Else
        MsgBox strErrMsg, vbCritical, "Error!"
    End If
End Function

Rem @Subfolder Files>TXTFiles
Sub TxtOverwrite(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim FileNumber As Integer
    FileNumber = FreeFile
    Open sFile For Output As #FileNumber
    Print #FileNumber, sText
    Close #FileNumber
Exit_Err_Handler:
    Exit Sub
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: TxtOverwrite" & vbCrLf & _
           "Error Description: " & err.Description, vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Sub

Function TXTtoArray(sFile$)
    Rem https://newbedev.com/vb-vba-import-csv-to-array-code-example
    Rem VBA function to open a CSV file in memory and parse it to a 2D array without ever touching a worksheet:
    '#INCLUDE OpenTextFile
    Dim c&, i&, j&, p&, d$, s$, rows&, cols&, a, r, v
    Const Q = """", QQ = Q & Q
    Const ENQ = ""
    Const ESC = ""
    Const COM = ","
    d = OpenTextFile$(sFile)
    If LenB(d) Then
        r = Split(Trim(d), vbCrLf)
        rows = UBound(r) + 1
        cols = UBound(Split(r(0), ",")) + 1
        ReDim v(1 To rows, 1 To cols)
        For i = 1 To rows
            s = r(i - 1)
            If LenB(s) Then
                If InStrB(s, QQ) Then s = Replace(s, QQ, ENQ)
                For p = 1 To Len(s)
                    Select Case Mid(s, p, 1)
                        Case Q:   c = c + 1
                        Case COM: If c Mod 2 Then Mid(s, p, 1) = ESC
                    End Select
                Next
                If InStrB(s, Q) Then s = Replace(s, Q, "")
                a = Split(s, COM)
                For j = 1 To cols
                    s = a(j - 1)
                    If InStrB(s, ESC) Then s = Replace(s, ESC, COM)
                    If InStrB(s, ENQ) Then s = Replace(s, ENQ, Q)
                    v(i, j) = s
                Next
            End If
        Next
        TXTtoArray = v
    End If
End Function

Rem insert string to txt file (not append, but on top)
Sub TxtPretend(FilePath As String, txt As String)
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    Dim s As String
    s = TxtRead(FilePath)
    TxtOverwrite FilePath, txt & Chr(10) & s
End Sub

Function TxtAppend(sFile As String, sText As String)
    On Error GoTo ERR_HANDLER
    Dim iFileNumber           As Integer
    iFileNumber = FreeFile
    Open sFile For Append As #iFileNumber
    Print #iFileNumber, sText
    Close #iFileNumber
Exit_Err_Handler:
    Exit Function
ERR_HANDLER:
    MsgBox "The following error has occurred" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: Txt_Append" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occurred!"
    GoTo Exit_Err_Handler
End Function

Function OpenTextFile$(f)
    With CreateObject("ADODB.Stream")
        .Charset = "utf-8"
        .Open
        .LoadFromFile f
        OpenTextFile = .ReadText
        .Close
    End With
End Function

Function TxtRead(sPath As Variant) As String
    Dim sTXT As String
    If Dir(sPath) = "" Then
        MsgBox "File was not found."
        Exit Function
    End If
    Open sPath For Input As #1
    Do Until EOF(1)
        Line Input #1, sTXT
        TxtRead = TxtRead & sTXT & vbLf
    Loop
    Close
    If Len(TxtRead) = 0 Then
        TxtRead = ""
    Else
        TxtRead = left(TxtRead, Len(TxtRead) - 1)
    End If
End Function

Sub testSplitProcTXT()
    '#INCLUDE SplitATextFileintoIndividualOnes
    SplitATextFileintoIndividualOnes "C:\Users\acer\Desktop\test\ArrayContainsInRowOrColumn2D.txt"
End Sub

Sub SplitATextFileintoIndividualOnes(FilePath As Variant)
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    Dim FName As String
    Dim Pth As String
    Dim txt As String
    Dim i As Long
    Pth = left(FilePath, InStrRev(FilePath, "\"))
    txt = TxtRead(FilePath)
    a = Split(TxtRead(FilePath), vbLf)
    Dim out As String
    For i = LBound(a) To UBound(a)
        If InStr(1, a(i), "Declare ") > 0 Then
            Do While Right(Trim(a(i)), 1) = "_"
                i = i + 1
            Loop
        End If
        out = IIf(out = "", a(i), out & a(i)) & vbNewLine
        If InStr(1, a(i), "Sub ") > 0 Then
            FName = Split(a(i), "Sub ")(1)
            FName = Trim(Split(FName, "(")(0)) & ".txt"
        ElseIf InStr(1, a(i), "Function ") > 0 Then
            FName = Split(a(i), "Function ")(1)
            FName = Trim(Split(FName, "(")(0)) & ".txt"
        End If
        If InStr(1, a(i), "End Sub") > 0 Or InStr(1, a(i), "End Function") > 0 Then
            TxtOverwrite Pth & FName, out
            out = ""
            FName = ""
        End If
    Next
    Set f = Nothing
    Set fs = Nothing
End Sub

Sub MergeFileText(FileCollection As Collection, NewFile As String, Optional criteria As String = "*.txt")
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    Dim s As String
    For Each item In FileCollection
        s = s & vbNewLine & TxtRead(FolderPath & item)
    Next
    TxtOverwrite NewFile, s
End Sub

Sub PretendListOfContainedProceduresInTXT(fileName As String)
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    '#INCLUDE ProceduresOfTXT
    Dim v As Variant: v = ProceduresOfTXT(fileName, True)
    If TypeName(v) = "Empty" Then Exit Sub
    Dim s As String: s = TxtRead(fileName)
    Dim line As String: line = String(30, "'")
    TxtOverwrite fileName, _
                 line & vbNewLine & _
                 "'Contains the following " & "#" & UBound(v) & " procedures " & vbNewLine & line & vbNewLine & "'" & _
                 Join(v, vbNewLine & "'") & vbNewLine & vbNewLine & s
End Sub

Function ProceduresOfTXT(FilePath As Variant, Optional NameOnly As Boolean) As Variant
    '#INCLUDE SortArray
    '#INCLUDE joinArrays
    '#INCLUDE TxtRead
    Dim var
    var = Split(TxtRead(CStr(FilePath)), Chr(10))
    Dim out
    out = joinArrays(Filter(var, "Sub "), Filter(var, "Function "))
    If TypeName(out) = "Empty" Then Exit Function
    out = Filter(out, "(", True)
    out = Filter(out, "Declare", False)
    out = Filter(out, Chr(34) & "Sub ", False)
    out = Filter(out, Chr(34) & "Function ", False)
    If NameOnly = True Then
        Dim i As Long
        For i = LBound(out) To UBound(out)
            out(i) = left(out(i), InStr(1, out(i), "(") - 1)
            out(i) = Replace(out(i), "Private ", "")
            out(i) = Replace(out(i), "Public ", "")
            out(i) = Replace(out(i), "Sub ", "")
            out(i) = Replace(out(i), "Function ", "")
        Next
    End If
    out = SortArray(out)
    ProceduresOfTXT = out
    Rem ProceduresOfTXT = Join(out, Chr(10))
End Function

Sub ListProceduresOfTXT(FilePaths As Variant, Optional NameOnly As Boolean)
    '#INCLUDE ArrayToRange1d
    '#INCLUDE ProceduresOfTXT
    Dim fileName As String
    Dim var
    Dim out As String
    Dim element
    Dim FileElement
    If TypeName(FilePaths) = "String" Then
        var = ProceduresOfTXT(FilePaths, NameOnly)
        For Each element In var
            out = IIf(out = "", element & "," & FilePaths, out & vbNewLine & element & "," & FilePaths)
        Next
    Else
        For Each FileElement In FilePaths
            var = ProceduresOfTXT(FileElement, NameOnly)
            For Each element In var
                out = IIf(out = "", element & "," & FileElement, out & vbNewLine & element & "," & FileElement)
            Next
        Next
    End If
    ArrayToRange1d Split(out, vbNewLine)
End Sub

Sub TxtRemoveBlankLines(FileFullPath As Variant)
    '#INCLUDE OpenTextFile
    Const ForReading = 1
    Const ForWriting = 2
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(FileFullPath, ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        If Len(Trim(strLine)) > 0 Then
            strNewContents = strNewContents & strLine & vbCrLf
        End If
    Loop
    objFile.Close
    Set objFile = objFSO.OpenTextFile(FileFullPath, ForWriting)
    objFile.Write strNewContents
    objFile.Close
End Sub

Sub TxtRemoveComments(FileFullPath As Variant)
    '#INCLUDE OpenTextFile
    Const ForReading = 1
    Const ForWriting = 2
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Set objFile = objFSO.OpenTextFile(FileFullPath, ForReading)
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        If left(Trim(strLine), 1) <> "'" Then
            strNewContents = strNewContents & strLine & vbCrLf
        End If
    Loop
    objFile.Close
    Set objFile = objFSO.OpenTextFile(FileFullPath, ForWriting)
    objFile.Write strNewContents
    objFile.Close
End Sub

Rem @Subfolder Files>ZIPFiles
Sub FilesAndOrFoldersInFolderOrZipDemo()
    '#INCLUDE dp
    '#INCLUDE FilesAndOrFoldersInFolderOrZip
    Dim out As New Collection
    FilesAndOrFoldersInFolderOrZip _
        FolderOrZipFilePath:="C:\Users\acer\Dropbox\SOFTWARE\EXCEL\00 Review", _
        LogFolders:=True, _
        LogFiles:=False, _
        ScanInSubfolders:=False, _
        out:=out
    dp out
End Sub

Function FilesAndOrFoldersInFolderOrZip(ByVal FolderOrZipFilePath As Variant, LogFolders As Boolean, LogFiles As Boolean, ScanInSubfolders As Boolean, out As Collection, Optional Filter As String = "*")
    Dim oSh As New Shell
    Dim ofi As Object
    For Each ofi In oSh.Namespace(FolderOrZipFilePath).items
        If ofi.IsFolder Then
            If LogFolders Then
                out.Add ofi.Path & "\"
            End If
            If ScanInSubfolders Then FilesAndOrFoldersInFolderOrZip ofi.Path, LogFolders, LogFiles, ScanInSubfolders, out, Filter
        Else
            If LogFiles Then
                If UCase(ofi.Name) Like UCase(Filter) Then
                    out.Add ofi.Path
                End If
            End If
        End If
    Next
    Set FilesAndOrFoldersInFolderOrZip = out
    Set oSh = Nothing
End Function

Public Sub UnzipToOwnFolder(ZippedFile As String, DeleteExistingFiles As Boolean, DeleteZip As Boolean)
    Rem for each cell in selection.cells: UnzipToOwnFolder cell.text,False,false :next
    '#INCLUDE RecycleSafe
    '#INCLUDE FoldersCreate
    '#INCLUDE FolderExists
    '#INCLUDE FilesAndOrFoldersInFolderOrZip
    Dim FileCollection As New Collection
    FilesAndOrFoldersInFolderOrZip ZippedFile, False, True, False, FileCollection
    Dim FolderCollection As New Collection
    FilesAndOrFoldersInFolderOrZip ZippedFile, True, False, False, FolderCollection
    Dim shell_app           As Object:     Set shell_app = CreateObject("Shell.Application")
    Rem   Dim FilesInZip          As Long:        FilesInZip = shell_app.Namespace(CVar(ZippedFile)).items.Count
    Dim LastSlash            As Long:       LastSlash = InStrRev(ZippedFile, "\")
    Dim Dot                      As Long:      Dot = InStrRev(ZippedFile, ".")
    Dim ParentFolder       As String:     ParentFolder = left(ZippedFile, LastSlash)
    Dim UnzipToFolder   As String
    If FolderCollection.count = 1 And FileCollection.count = 0 Then
        UnzipToFolder = ParentFolder
    ElseIf FolderCollection.count > 1 Or FileCollection.count > 0 Then
        UnzipToFolder = left(ZippedFile, Dot - 1) & "\"
        If DeleteExistingFiles Then
            If FolderExists(UnzipToFolder) Then RecycleSafe UnzipToFolder
        End If
        FoldersCreate UnzipToFolder
    End If
    shell_app.Namespace(CVar(UnzipToFolder)).CopyHere shell_app.Namespace(CVar(ZippedFile)).items
    If DeleteZip Then RecycleSafe ZippedFile
    Set shell_app = Nothing
End Sub

Public Sub UnzipAllInFolder(source_folder As String)
    '#INCLUDE create_temp_zip_folder
    '#INCLUDE Zip
    Dim current_zip_file As String
    current_zip_file = Dir(source_folder & "\*.zip")
    If Len(current_zip_file) = 0 Then
        MsgBox "No zip files found!", vbExclamation
        Exit Sub
    End If
    Dim zip_folder As String
    zip_folder = source_folder & "\unzipped"
    Dim error_message As String
    If Not create_temp_zip_folder(zip_folder, error_message) Then
        MsgBox error_message, vbCritical, "Error"
        Exit Sub
    End If
    Dim shell_app As Object
    Set shell_app = CreateObject("Shell.Application")
    Do While Len(current_zip_file) > 0
        shell_app.Namespace(CVar(zip_folder)).CopyHere shell_app.Namespace(source_folder & "\" & current_zip_file).items
        current_zip_file = Dir
    Loop
    Set shell_app = Nothing
End Sub

Function create_temp_zip_folder(ByVal zip_folder As String, ByRef error_message As String) As Boolean
    '#INCLUDE FolderExists
    On Error GoTo Error_Handler
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FolderExists(zip_folder) Then
        fso.DeleteFolder zip_folder, True
    End If
    fso.CreateFolder zip_folder
    create_temp_zip_folder = True
    Set fso = Nothing
    Exit Function
Error_Handler:
    error_message = "Error " & err.Number & ":" & vbCrLf & vbCrLf & err.Description
End Function

Public Function getFolder() As String
    With Application.FileDialog(msoFileDialogFolderPicker)
        .AllowMultiSelect = False
        .title = "Select a folder"
        .Show
        If .SelectedItems.count > 0 Then
            getFolder = .SelectedItems.item(1)
        Else
            'MsgBox "Folder is not selected."
        End If
    End With
End Function


