Attribute VB_Name = "U_ProjectManager"

Rem @Folder ProjectManager Declarations
Public pmWorkbook As Workbook
#If VBA7 Then
    Declare PtrSafe Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As LongPtr)
#Else
    Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte,  ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
#End If
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Rem @Folder WindowToPDF Declarations
Private Const VK_SNAPSHOT = 44
Private Const VK_LMENU = 164
Private Const KEYEVENTF_KEYUP = 2
Private Const KEYEVENTF_EXTENDEDKEY = 1
Rem @Folder CodePrinter Declarations
Public PrintFileName As String
Public Found1 As String
Public Found2 As String
Dim rng As Range
Public cell As Range
Public s As Shape
Public counter As Long

Rem @Folder ProjectManager
Rem @Subfolder CodePrinter
Public Function PrintProject(TargetWorkbook As Workbook)
    '#INCLUDE HasProject
    '#INCLUDE ProtectedVBProject
    '#INCLUDE PrinterTocAndCode
    '#INCLUDE PutCodeInPrinter
    '#INCLUDE LinkCodeBlocksWithShape
    '#INCLUDE AddPageBreaksToPrinter
    '#INCLUDE MergePrinterCells
    '#INCLUDE NumberLinesPrinter
    '#INCLUDE FormatPrinterTitles
    '#INCLUDE AddLogoToFirstPage
    '#INCLUDE SetupPrinterPage
    '#INCLUDE ColorizeBlockLinksByLevel
    '#INCLUDE FormatTextColors
    '#INCLUDE ResetPrinter
    '#INCLUDE PrintPDF
    '#INCLUDE AutofitMergedCells
    '#INCLUDE StartOptimizeCodeRun
    '#INCLUDE StopOptimizeCodeRun
    '#INCLUDE getLastRow
    Set pmWorkbook = TargetWorkbook
    If ProtectedVBProject(TargetWorkbook) = True Or HasProject(TargetWorkbook) = False Then
        MsgBox "Project Empty or Protected"
        Exit Function
    End If
    StartOptimizeCodeRun
    ResetPrinter
    PutCodeInPrinter PrinterTocAndCode
    NumberLinesPrinter
    FormatTextColors
    FormatPrinterTitles
    SetupPrinterPage
    AddLogoToFirstPage
    LinkCodeBlocksWithShape
    ColorizeBlockLinksByLevel
    MergePrinterCells
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    Dim TargetRange As Range
    Set TargetRange = ws.Range("P2:P" & getLastRow(ws))
    AutofitMergedCells TargetRange
    AddPageBreaksToPrinter
    PrintPDF
    StopOptimizeCodeRun
End Function

Function PrinterTocAndCode()
    '#INCLUDE GetCompText
    '#INCLUDE ComponentTypeToString
    '#INCLUDE CollectionToArray
    Dim workbookName As String
    workbookName = TargetWorkbook.Name
    Dim TargetWorkSheet As Worksheet
    Dim TargetWorkSheetName As String
    Dim Module As VBComponent
    Dim procedures As Collection
    Set procedures = New Collection
    Dim Procedure As String
    Dim tmpString As Variant
    Dim i As Long
    procedures.Add "Table of Contents:"
    Dim ModuleTypes
    ModuleTypes = Array(vbext_ct_Document, vbext_ct_ClassModule, vbext_ct_StdModule, vbext_ct_MSForm)
    Dim ModuleType
    For Each ModuleType In ModuleTypes
        For Each Module In TargetWorkbook.VBProject.VBComponents
            If Module.Type = ModuleType Then
                If ModuleType = vbext_ct_Document And Module.Name <> "ThisWorkbook" Then
                    For Each TargetWorkSheet In TargetWorkbook.Worksheets
                        If TargetWorkSheet.CodeName = Module.Name Then TargetWorkSheetName = TargetWorkSheet.Name
                    Next TargetWorkSheet
                    procedures.Add "(" & ComponentTypeToString(Module.Type) & ")" & " " & TargetWorkSheetName & " - " & Module.Name
                    TargetWorkSheetName = ""
                Else
                    procedures.Add "(" & ComponentTypeToString(Module.Type) & ")" & " " & Module.Name
                End If
            End If
        Next Module
    Next ModuleType
    Rem document
    For Each ModuleType In ModuleTypes
        For Each Module In TargetWorkbook.VBProject.VBComponents
            If Module.Type = ModuleType Then
                If ModuleType = vbext_ct_Document And Module.Name <> "ThisWorkbook" Then
                    For Each TargetWorkSheet In TargetWorkbook.Worksheets
                        If TargetWorkSheet.CodeName = Module.Name Then TargetWorkSheetName = TargetWorkSheet.Name
                    Next TargetWorkSheet
                    procedures.Add "--- " & TargetWorkSheetName & " - " & Module.Name & " ---"
                    TargetWorkSheetName = ""
                Else
                    procedures.Add "--- " & Module.Name & " ---"
                End If
                If Module.CodeModule.CountOfLines > 0 Then
                    tmpString = Split(GetCompText(Module), vbNewLine)
                    For i = LBound(tmpString) To UBound(tmpString)
                        procedures.Add tmpString(i)
                    Next i
                End If
            End If
        Next Module
    Next ModuleType
    ThisWorkbook.SHEETS("ProjectManagerPrinter").Cells.WrapText = False
    PrinterTocAndCode = CollectionToArray(procedures)
End Function

Sub PutCodeInPrinter(SourceCode As Variant)
    '#INCLUDE getLastRow
    Dim TargetSheet As Worksheet
    Set TargetSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    Dim i As Long
    Dim off As Long
    Dim rowNo As Long
    rowNo = getLastRow(TargetSheet) + 1
    With TargetSheet
        For i = LBound(SourceCode) To UBound(SourceCode)
            If left(Trim(SourceCode(i)), 1) = "'" Then
                .Cells(rowNo, 2).Value = "'"
            ElseIf Len(SourceCode(i)) - Len(LTrim(SourceCode(i))) = 0 Then
                .Cells(rowNo, 2).Value = Trim(SourceCode(i))
            Else
                off = Len(SourceCode(i)) - Len(LTrim(SourceCode(i)))
                .Cells(rowNo, 2).OFFSET(0, off / 4).Value = Trim(SourceCode(i))
            End If
            rowNo = rowNo + 1
        Next
    End With
End Sub

Function LinkCodeBlocksWithShape() As Boolean
    '#INCLUDE IsBlockStart
    '#INCLUDE openPair
    '#INCLUDE closePair
    '#INCLUDE getLastColumn
    '#INCLUDE AddShape
    Dim ShapeTypeNumber As Long
    Rem msoShapeLeftBrace   31  Left brace
    Rem msoShapeLeftBracket 29  Left bracket
    ShapeTypeNumber = 29
    Rem activesheet.shapes("bracket").adjustments.item(1)=10
    Dim CloseTXT As String
    Dim X As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    Dim trimCell As String
    Dim counter As Long
    Dim colNo As Long
    For colNo = 1 To getLastColumn(ws)
        If WorksheetFunction.CountA(ws.Columns(colNo)) > 0 Then
            For Each cell In ws.Columns(colNo).SpecialCells(xlCellTypeConstants)
                trimCell = Trim(cell.TEXT)
                If IsBlockStart(trimCell) Then
                    Select Case openPair(trimCell)
                        Case Is = "Case", "Else"
                        Case Is = "If"
                            If Right(trimCell, 4) <> "Then" Then GoTo Skip
                        Case Is = "#If"
                        Case Is = "skip"
                            GoTo Skip
                        Case Else
                    End Select
                    CloseTXT = closePair(trimCell)
                    counter = Len(cell) - Len(trimCell)
                    Found1 = cell.Address
                    On Error Resume Next
                    Dim foundMatch As Range
                    Set foundMatch = ws.Columns(colNo).Find(CloseTXT & "*", After:=cell, LookAt:=xlWhole)
                    On Error GoTo 0
                    If foundMatch Is Nothing Then GoTo Skip
                    Found2 = foundMatch.Address
                    ws.Shapes.AddShape ShapeTypeNumber, _
                                       ws.Range(Found1).left - 5, _
                                       ws.Range(Found1).top + (cell.Height / 2), _
                                       5, _
                                       ws.Range(Found1, Found2).Height - cell.Height
                End If
Skip:
            Next cell
        End If
    Next colNo
    LinkCodeBlocksWithShape = True
End Function

Function openPair(strLine As String) As String
    Dim nPos As Integer
    Dim strTemp As String
    strTemp = Trim(strLine)
    nPos = InStr(1, strTemp, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
        Case Is = "Private", "Public"
            strTemp = Trim(strLine)
            strTemp = Replace(strTemp, "Private ", "")
            strTemp = Replace(strTemp, "Public ", "")
            nPos = InStr(1, strTemp, " ") - 1
            If nPos < 0 Then nPos = Len(strTemp)
            strTemp = left$(strTemp, nPos)
            If strTemp = "Function" Then
                openPair = "Function"
            ElseIf strTemp = "Sub" Then
                openPair = "Sub"
            Else
                GoTo Skip
            End If
        Case Is = "With"
            openPair = "With"
        Case Is = "For"
            openPair = "For"
        Case Is = "Do"
            openPair = "Do"
        Case Is = "While"
            openPair = "While"
        Case Is = "Select"
            openPair = "Select"
        Case Is = "Case"
            openPair = "Case"
        Case Is = "Sub"
            openPair = "Sub"
        Case Is = "Function"
            openPair = "Function"
        Case Is = "Property"
            openPair = "Property"
        Case Is = "Enum"
            openPair = "Enum"
        Case Is = "Type"
            openPair = "Type"
        Case "If"
            openPair = "If"
        Case "ElseIf", "Else", "Else:"
            openPair = "Else"
        Case "#If"
            openPair = "#If"
        Case "#ElseIf", "#Else", "#Else:"
            openPair = "#Else"
        Case Else
Skip:
            openPair = "skip"
    End Select
End Function

Function closePair(strLine As String) As String
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
        Case Is = "Private", "Public"
            strTemp = Trim(strLine)
            strTemp = Replace(strTemp, "Private ", "")
            strTemp = Replace(strTemp, "Public ", "")
            nPos = InStr(1, strTemp, " ") - 1
            If nPos < 0 Then nPos = Len(strTemp)
            strTemp = left$(strTemp, nPos)
            If strTemp = "Function" Then
                closePair = "End Function"
            ElseIf strTemp = "Sub" Then
                closePair = "End Sub"
            Else
            End If
        Case Is = "With"
            closePair = "End With"
        Case Is = "For"
            closePair = "Next"
        Case Is = "Do", "While"
            closePair = "Loop"
        Case Is = "Select", "Case"
            closePair = "End Select"
        Case Is = "Sub"
            closePair = "End Sub"
        Case Is = "Function"
            closePair = "End Function"
        Case Is = "Property"
            closePair = "End Property"
        Case Is = "Enum"
            closePair = "End Enum"
        Case Is = "Type"
            closePair = "End Type"
        Case "If", "ElseIf", "Else", "Else:"
            closePair = "End If"
        Case "#If", "#ElseIf", "#Else", "#Else:"
            closePair = "#End If"
        Case Else
    End Select
End Function

Sub AddPageBreaksToPrinter()
    ThisWorkbook.SHEETS("ProjectManagerPrinter").ResetAllPageBreaks
    Dim rng As Range
    Set rng = Nothing
    Dim cell As Range
    With ThisWorkbook.SHEETS("ProjectManagerPrinter")
        For Each cell In .Range("B1:B" & .Range("B" & .rows.count).End(xlUp).row)
            If left(Trim(cell.Value), 3) = "---" Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        Next
        For Each cell In rng
            .HPageBreaks.Add Before:=.rows(cell.row)
            .rows(cell.row).PageBreak = xlPageBreakManual
        Next
    End With
End Sub

Sub MergePrinterCells()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    Dim cell As Range
    For Each cell In ws.Cells.SpecialCells(xlCellTypeConstants)
        ws.Range(cell, ws.Cells(cell.row, "P")).MERGE
    Next
    ws.Cells.WrapText = True
End Sub

Sub FormatColourFormatters()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("ProjectManagerTXTColour")
    uProjectManager.LBLcolourCode.ForeColor = ws.Range("J1").Value
    uProjectManager.LBLcolourKey.ForeColor = ws.Range("J3").Value
    uProjectManager.LBLcolourOdd.BackColor = ws.Range("J2").Value
    uProjectManager.LBLcolourComment.ForeColor = ws.Range("J4").Value
End Sub

Sub ColorPaletteDialog(rng As Range, Lbl As MSForms.label)
    If Application.Dialogs(xlDialogEditColor).Show(10, 0, 125, 125) = True Then
        lcolor = ActiveWorkbook.Colors(10)
        rng.Value = lcolor
        rng.OFFSET(0, 1).Interior.color = lcolor
        Lbl.ForeColor = lcolor
    End If
    ActiveWorkbook.ResetColors
End Sub

Sub NumberLinesPrinter()
    '#INCLUDE getLastRow
    Dim TargetWorkSheet As Worksheet
    Set TargetWorkSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    Dim cell As Range
    Dim lRow As Long
    lRow = getLastRow(TargetWorkSheet)
    Rem     Dim arr
    Rem     ReDim arr(1 To lRow)
    Rem     Dim i As Long
    Rem     For i = 1 To lRow
    Rem         arr(i) = i
    Rem     Next
    Rem     TargetWorkSheet.Range("a1").Value = WorksheetFunction.Transpose(arr)
    With TargetWorkSheet
        For Each cell In .Range("A1:A" & lRow)
            If cell.row Mod 2 = 0 Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        Next cell
    End With
    rng.EntireRow.Interior.color = _
                                 ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J2").Value
End Sub

Sub FormatPrinterTitles()
    Dim rng As Range
    Set rng = Nothing
    Dim cell As Range
    With ThisWorkbook.SHEETS("ProjectManagerPrinter")
        For Each cell In .Range("B1:B" & .Range("B" & .rows.count).End(xlUp).row)
            If left(Trim(cell.Value), 3) = "---" Then
                If rng Is Nothing Then
                    Set rng = cell
                Else
                    Set rng = Union(rng, cell)
                End If
            End If
        Next
    End With
    If rng Is Nothing Then Exit Sub
    rng.Font.Size = 18
    rng.Font.Bold = True
    rng.Font.color = vbBlack
End Sub

Sub AddLogoToFirstPage()
    Dim PrinterSheet As Worksheet
    Set PrinterSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    PrinterSheet.rows(1).EntireRow.Insert
    PrinterSheet.Range("A2:C2").Interior.ColorIndex = 0
    Dim TargetCell As Range
    Set TargetCell = PrinterSheet.Range("B1")
    With TargetCell
        Rem       .HorizontalAlignment = xlCenter
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlVAlignBottom
        .Value = AUTHOR_GITHUB & Space(16) & pmWorkbook.Name
        .Characters.Font.Size = 12
        .Characters.Font.Bold = True
        .Characters.Font.Underline = False
        .Characters.Font.ColorIndex = 10
        .Characters.Font.Name = "Comic Sans MS"
        .RowHeight = 330
    End With
    ThisWorkbook.SHEETS("README").Shapes("LOGO").Copy
    PrinterSheet.Paste TargetCell
    Dim shp As Shape
    Set shp = PrinterSheet.Shapes("LOGO")
    shp.left = TargetCell.left
    shp.top = TargetCell.top
    shp.Height = TargetCell.EntireRow.Height
End Sub

Sub SetupPrinterPage()
    Dim PrinterSheet As Worksheet
    Set PrinterSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    With PrinterSheet.PageSetup
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.25)
        .BottomMargin = Application.InchesToPoints(0.75)
        Dim fileName As String
        fileName = PrintFileName
        .LeftFooter = fileName
        .CenterFooter = "Page &P of &N"
        .RightFooter = "&D"
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With
End Sub

Sub ColorizeBlockLinksByLevel()
    '#INCLUDE RandomRGB
    Dim rnd As Long
    Dim n As Variant
    Dim i As Long
    Dim s As Shape
    Dim sNames
    Set sNames = CreateObject("System.Collections.ArrayList")
    For Each s In ThisWorkbook.SHEETS("ProjectManagerPrinter").Shapes
        If UCase(s.Name) <> "LOGO" Then
            s.Name = s.left
            If Not sNames.CONTAINS(s.Name) Then
                sNames.Add s.Name
            End If
        End If
    Next s
    For Each n In sNames
        rnd = RandomRGB
        For Each s In ThisWorkbook.SHEETS("ProjectManagerPrinter").Shapes
            If UCase(s.Name) <> "LOGO" Then
                If s.Name = n Then
                    With s.line
                        .ForeColor.RGB = rnd
                        .Weight = 1.5
                    End With
                End If
            End If
        Next s
    Next n
    Set sNames = Nothing
End Sub

Function RandomRGB()
    RandomRGB = RGB(Int(rnd() * 255), Int(rnd() * 255), Int(rnd() * 255))
End Function

Public Sub FormatTextColors()
    '#INCLUDE InStrExact
    Dim PrinterSheet As Worksheet
    Set PrinterSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    With PrinterSheet.Cells.Font
        .color = ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J1").Value
        .FontStyle = "Normal"
    End With
    Dim rng As Range
    Set rng = PrinterSheet.UsedRange.SpecialCells(xlCellTypeConstants)
    Dim cell As Range
    Dim NumChars As Long
    Dim StartChar As Long
    Dim cellChar As Long
    Dim EndWords As Long
    Dim keywords As Range
    On Error Resume Next
    For Each cell In rng
        If left(Trim(cell.Value), 1) = "'" Or left(Trim(cell.Value), 3) = "Rem" Then
            cell.Font.color = ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J4").Value
        Else
            cellChar = Len(cell)
            For Each keywords In ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("A1").CurrentRegion.OFFSET(1).RESIZE(, 1).SpecialCells(xlCellTypeConstants)
                If InStr(1, cell.TEXT, keywords) > 0 Then
                    StartChar = InStrExact(1, cell.TEXT, keywords.TEXT)
                    Do Until StartChar >= cellChar Or StartChar = 0
                        NumChars = Len(keywords.TEXT)
                        EndWords = StartChar + NumChars
                        If Mid(cell.TEXT, StartChar - 1, 1) = " " Or StartChar = 1 Then
                            If Mid(cell.TEXT, EndWords, 1) = " " Or EndWords >= cellChar Then
                                With cell.Characters(Start:=StartChar, Length:=NumChars).Font
                                    .FontStyle = "Bold"
                                    .color = ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J3").Value
                                End With
                            End If
                        End If
                        StartChar = InStr(EndWords, cell.TEXT, keywords.TEXT)
                    Loop
                End If
            Next
        End If
    Next
End Sub

Sub ResetPrinter(Optional keepText As Boolean = False)
    Dim PrinterSheet As Worksheet
    Set PrinterSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    With PrinterSheet
        .ResetAllPageBreaks
        .rows.VerticalAlignment = xlVAlignTop
        With .Cells
            If keepText = False Then
                .clear
            Else
                .ClearFormats
                .Font.ColorIndex = vbBlack
                .Font.Bold = False
            End If
            .Font.Name = "Consolas"
            .WrapText = False
            .UseStandardHeight = True
        End With
    End With
    For Each s In PrinterSheet.Shapes
        s.Delete
    Next
    Rem     If .PageSetup.Orientation = xlPortrait Then
    Rem         .Columns("B:B").ColumnWidth = 90
    Rem     Else
    Rem         .Columns("B:B").ColumnWidth = 120
    Rem     End If
End Sub

Sub PrintPDF()
    '#INCLUDE FoldersCreate
    Dim FilePath As String
    FilePath = Environ("USERprofile") & "\Documents\" & "vbArc\CodePrinter\"
    Dim fileName As String
    fileName = left(pmWorkbook.Name, InStr(1, pmWorkbook.Name, ".") - 1)
    Dim saveLocation As String
    saveLocation = FilePath
    If Dir(saveLocation, vbDirectory) = "" Then
        FoldersCreate saveLocation
    End If
    FilePath = saveLocation & fileName
    ThisWorkbook.Worksheets("ProjectManagerPrinter").visible = xlSheetVisible
    Dim TargetSheet As Worksheet
    Set TargetSheet = ThisWorkbook.SHEETS("ProjectManagerPrinter")
    TargetSheet.ExportAsFixedFormat xlTypePDF, FilePath
End Sub

Public Sub delay(seconds As Long)
    Dim endTime As Date
    endTime = DateAdd("s", seconds, Now())
    Do While Now() < endTime
        DoEvents
    Loop
End Sub

Sub UserformToPDF(wb As Workbook, Path As String)
    '#INCLUDE WindowToPDF
    '#INCLUDE PathMaker
    Application.VBE.MainWindow.visible = True
    Do While Application.VBE.MainWindow.visible = False
        Sleep 1000
    Loop
    Application.VBE.MainWindow.WindowState = vbext_ws_Maximize
    DoEvents
    Sleep 1000
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_MSForm Then
            vbComp.Activate
            DoEvents
            Sleep 3000
            Call WindowToPDF(PathMaker(Path, vbComp.Name, "pdf"))
        End If
    Next
End Sub

Sub ExportWorksheetsToPDF(wb As Workbook, expPath As String)
    '#INCLUDE PathMaker
    wb.Activate
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        Application.PrintCommunication = False
        With ws.PageSetup
            .FitToPagesWide = 1
            .FitToPagesTall = False
        End With
        Application.PrintCommunication = True
        If ws.UsedRange.count > 0 Then
            ws.ExportAsFixedFormat xlTypePDF, PathMaker(expPath, ws.Name, "pdf"), , True
        End If
    Next ws
End Sub

Rem @Folder WindowToPDF
Function WindowToPDF(pdf$, Optional Orientation As Integer = xlLandscape, _
                     Optional FitToPagesWide As Integer = 1) As Boolean
    Dim calc As Integer, ws As Worksheet
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        calc = .Calculation
        .Calculation = xlCalculationManual
    End With
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0
    keybd_event VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    keybd_event VK_LMENU, 0, KEYEVENTF_EXTENDEDKEY + KEYEVENTF_KEYUP, 0
    Set ws = Workbooks.Add(xlWBATWorksheet).Worksheets(1)
    Application.Wait (Now + TimeValue("0:00:2"))
    With ws
        .PasteSpecial Format:="Bitmap", link:=False, DisplayAsIcon:=False
        .Range("A1").Select
        .PageSetup.Orientation = Orientation
        .PageSetup.FitToPagesWide = FitToPagesWide
        .PageSetup.Zoom = False
        .ExportAsFixedFormat Type:=xlTypePDF, fileName:=pdf, _
                             quality:=xlQualityStandard, IncludeDocProperties:=True, _
                             IgnorePrintAreas:=False, OpenAfterPublish:=False
        .parent.Close False
    End With
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = calc
        .CutCopyMode = False
    End With
    WindowToPDF = Dir(pdf) <> ""
End Function

Function PathMaker(wbPath As String, fileName As String, fileExtention As String) As String
    If Right(wbPath, 1) <> "\" Then wbPath = wbPath & "\"
    PathMaker = wbPath & fileName & "." & fileExtention
    Do While InStr(1, PathMaker, "..") > 0
        PathMaker = Replace(PathMaker, "..", ".")
    Loop
End Function

Rem imports
Sub ImportComponents(wb As Workbook)
    '#INCLUDE ModuleExists
    '#INCLUDE IsArrayAllocated
    '#INCLUDE WorkbookIsOpen
    '#INCLUDE GetFilePartPath
    '#INCLUDE getFilePartName
    '#INCLUDE GetFilePath
    Dim varr
    Dim element
    Dim Proceed As Boolean, hasWorksheets As Boolean
    Proceed = True
    Dim compName As String
    varr = GetFilePath(Array("bas", "frm", "cls"), True)
    If Not IsArrayAllocated(varr) Then Exit Sub
    Dim vbProj As VBProject
    Set vbProj = wb.VBProject
    Dim coll As Collection
    Set coll = New Collection
    For Each element In varr
        compName = getFilePartName(CStr(element), False)
        Debug.Print compName
        If compName Like "DocClass*" Then
            compName = Right(compName, Len(compName) - 6)
            hasWorksheets = True
        End If
        If ModuleExists(compName, wb) = True Then
            Proceed = False
            coll.Add compName
        End If
    Next element
    If Proceed = False Then GoTo ErrorHandler
    Dim WasOpen As Boolean
    Dim wbSource As Workbook
    Dim wbSourceName As String
    Dim basePath As String
    basePath = GetFilePartPath(varr(1), True)
    If hasWorksheets = True Then
        wbSourceName = Dir(basePath & "*.xl*")
        If wbSourceName <> "" Then
            WasOpen = WorkbookIsOpen(wbSourceName)
            If WasOpen = False Then
                Set wbSource = Workbooks.Open(basePath & wbSourceName)
            Else
                Set wbSource = Workbooks(wbSourceName)
            End If
        End If
    End If
    For Each element In varr
        compName = getFilePartName(CStr(element), False)
        If Not compName Like "DocClass*" Then
            vbProj.VBComponents.Import element
        Else
            compName = Right(compName, Len(compName) - 9)
            If compName <> "ThisWorkbook" Then
                wbSource.SHEETS(compName).Copy Before:=wb.SHEETS(1)
            End If
        End If
    Next element
    GoTo ExitHandler
ErrorHandler:
    Dim str As String
    str = "The following components already exist. All import canceled."
    For Each element In coll
        str = str & vbNewLine & element
    Next element
    MsgBox str
    Exit Sub
ExitHandler:
    If WasOpen = False And WorkbookIsOpen(wbSourceName) Then wbSource.Close False
    Set vbProj = Nothing
    Set coll = Nothing
    Set wbSource = Nothing
    MsgBox "Import successful"
End Sub

Rem exports
Sub ExportProject( _
    wb As Workbook, _
    Optional ExportSheets As Boolean, _
    Optional ExportForms As Boolean, _
    Optional PrintCode As Boolean, _
    Optional bSeparateProcedures As Boolean, _
    Optional bExportComponents As Boolean, _
    Optional bWorkbookBackup As Boolean, _
    Optional bExportUnified As Boolean, _
    Optional bExportDeclarations As Boolean, _
    Optional bExportReferences As Boolean, _
    Optional bExportXML As Boolean)
    '#INCLUDE GetCompText
    '#INCLUDE ProcListCollection
    '#INCLUDE GetProcText
    '#INCLUDE GetSheetByCodeName
    '#INCLUDE PrintProject
    '#INCLUDE UserformToPDF
    '#INCLUDE ExportWorksheetsToPDF
    '#INCLUDE ExportReferencesToConfigFile
    '#INCLUDE ArrayToString
    '#INCLUDE FollowLink
    '#INCLUDE FoldersCreate
    '#INCLUDE TxtAppend
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE compare
    '#INCLUDE getDeclarations
    Dim workbookCleanName    As String: workbookCleanName = left(wb.Name, InStrRev(wb.Name, ".") - 1)
    Dim workbookExtension    As String: workbookExtension = Right(wb.Name, Len(wb.Name) - InStr(1, wb.Name, "."))
    Dim MainPath             As String: MainPath = Environ("USERprofile") & "\Documents\" & "vbArc\Code Library\"
    Dim exportPath           As String: exportPath = MainPath & workbookCleanName & "\"
    exportPath = exportPath & Format(Now, "YY-MM-DD HHNN") & "\"
    Rem create folders
    FoldersCreate exportPath
    Dim procColl As Collection, Procedure As Variant, vbComp As VBComponent, Extension As String
    If bExportComponents = True Then
        Rem Export Components
        For Each vbComp In wb.VBProject.VBComponents
            Select Case vbComp.Type
                Case vbext_ct_ClassModule, vbext_ct_Document: Extension = ".cls"
                Case vbext_ct_MSForm:       Extension = ".frm"
                Case vbext_ct_StdModule:    Extension = ".bas"
                Case Else:                  Extension = ".txt"
            End Select
            Rem if you import a docclass by this project,
            Rem it will open the original exported file and copy the sheet
            If vbComp.Type = vbext_ct_Document Then
                If vbComp.Name = "ThisWorkbook" Then
                    vbComp.Export exportPath & "DocClass " & vbComp.Name & Extension
                Else
                    vbComp.Export exportPath & "DocClass " & GetSheetByCodeName(wb, vbComp.Name).Name & Extension
                End If
            Else
                Rem export component
                vbComp.Export exportPath & vbComp.Name & Extension
            End If
        Next
    End If
    Rem export workbook backup
    If bWorkbookBackup = True Then
        wb.SaveCopyAs exportPath & wb.Name
    End If
    Rem export references
    If bExportReferences = True Then
        ExportReferencesToConfigFile pmWorkbook, exportPath
    End If
    Rem export declarations
    If bExportDeclarations = True Then
        Dim DeclarationArray As Variant
        DeclarationArray = CollectionsToArrayTable(getDeclarations(pmWorkbook))
        If TypeName(DeclarationArray) <> "Empty" Then
            TxtAppend exportPath & "Declarations.txt", ArrayToString(DeclarationArray)
        End If
    End If
    Rem export unified code to easily compare changes
    If bExportUnified Then
        Dim Code As String, tmp As String
        For Each vbComp In wb.VBProject.VBComponents
            tmp = "'" & vbComp.Name & vbTab & vbComp.Type & vbNewLine & vbNewLine & GetCompText(vbComp)
            Code = IIf(Code = "", tmp, Code & vbNewLine & vbNewLine & tmp)
        Next
        TxtAppend exportPath & "#UnifiedProject.txt", Code
    End If
    If bExportXML = True Then
        Rem Export ribbon xml (by JKP)
        Rem         Dim FullPath As String
        Rem         FullPath = wb.FullName
        Rem         wb.Close True
        Dim tmpFile As String
        tmpFile = ThisWorkbook.Path & "\temp_workbook_file" & Mid(wb.Name, InStr(1, wb.Name, "."))
        wb.SaveCopyAs tmpFile
        Dim c As New clsEditOpenXML
        Rem c.ExtractRibbonX FullPath, exportPath & "customUI.xml"
        c.ExtractRibbonX tmpFile, exportPath & "customUI.xml"
        Kill tmpFile
        Set c = Nothing
        Rem Workbooks.Open FullPath
    End If
    Rem Export procedures separately
    If bSeparateProcedures = True Then
        Dim ProcedurePath As String
        Dim ans As Long
        ans = MsgBox("If there are too many procedures the proccess will be slow. Proceed?", vbYesNo)
        If ans = vbYes Then
            For Each vbComp In wb.VBProject.VBComponents
                ProcedurePath = exportPath & vbComp.Name & " Procedures\"
                FoldersCreate ProcedurePath
                Set procColl = ProcListCollection(vbComp)
                Rem export component
                For Each Procedure In procColl
                    TxtAppend ProcedurePath & Procedure & ".txt", GetProcText(vbComp, CStr(Procedure))
                Next Procedure
            Next
        End If
    End If
    Rem Print to PDF, original feature (codeblocks linked, choose colour scheme)
    If PrintCode = True Then
        PrintFileName = wb.Name
        PrintProject wb
    End If
    Rem export Userform To PDF
    If ExportForms = True Then
        If wb.Name <> ThisWorkbook.Name Then UserformToPDF wb, exportPath
    End If
    Rem Export Worksheets To Image
    If ExportSheets = True Then
        Dim EXT As String: EXT = Right(wb.Name, Len(wb.Name) - InStr(1, wb.Name, "."))
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = False
        ExportWorksheetsToPDF wb, exportPath
        Sleep 1000
        If EXT = "xlam" Or EXT = "xla" Then wb.IsAddin = True
    End If
    Sleep 1000
    MsgBox "Export complete"
    Rem open export folder
    Rem FollowLink exportPath
End Sub

Rem by Todar
Public Sub ExportReferencesToConfigFile(TargetWorkbook As Workbook, RefPath As String)
    '#INCLUDE OpenTextFile
    Dim myProject As VBProject
    Set myProject = TargetWorkbook.VBProject
    Dim fso As New Scripting.FileSystemObject
    With fso.OpenTextFile(RefPath & "References.Txt", ForWriting, True)
        Dim library As Reference
        For Each library In myProject.REFERENCES
            .WriteLine library.Name & vbTab & library.GUID & vbTab & library.Major & vbTab & library.Minor
        Next
    End With
End Sub

Rem @TODO
Public Sub ImportReferencesFromConfigFile()
    '#INCLUDE OpenTextFile
    Dim fso As New Scripting.FileSystemObject
    With fso.OpenTextFile(exportPath & "References.Txt", ForReading, True)
        Dim line As Long
        Do While Not .AtEndOfStream
            Dim values As Variant
            values = Split(.ReadLine, vbTab)
            On Error Resume Next
            ThisWorkbook.VBProject.REFERENCES.AddFromGuid values(1), values(2), values(3)
        Loop
    End With
End Sub

Sub RefreshComponents(wkbSource As Workbook)
    '#INCLUDE ExportModules
    '#INCLUDE ImportModules
    If wkbSource.Name <> ThisWorkbook.Name Then
        ExportModules wkbSource
        ImportModules wkbSource
    Else
        MsgBox "Can't run this procedure on myself"
    End If
End Sub

Public Sub ExportModules(wkbSource As Workbook)
    '#INCLUDE FolderWithVBAProjectFiles
    Dim bExport As Boolean
    Dim szExportPath As String
    Dim szFileName As String
    Dim cmpComponent As VBIDE.VBComponent
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Export Folder not exist"
        Exit Sub
    End If
    On Error Resume Next
    Kill FolderWithVBAProjectFiles & "\*.*"
    On Error GoTo 0
    szExportPath = FolderWithVBAProjectFiles & "\"
    For Each cmpComponent In wkbSource.VBProject.VBComponents
        bExport = True
        szFileName = cmpComponent.Name
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                szFileName = szFileName & ".cls"
            Case vbext_ct_MSForm
                szFileName = szFileName & ".frm"
            Case vbext_ct_StdModule
                szFileName = szFileName & ".bas"
            Case vbext_ct_Document
                bExport = False
        End Select
        If bExport Then
            cmpComponent.Export szExportPath & szFileName
        End If
    Next cmpComponent
End Sub

Public Sub ImportModules(wkbTarget As Workbook)
    '#INCLUDE FolderWithVBAProjectFiles
    '#INCLUDE DeleteVBAModulesAndUserForms
    '#INCLUDE getFolder
    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.file
    Dim szImportPath As String
    Dim szFileName As String
    Dim cmpComponents As VBIDE.VBComponents
    If wkbTarget.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
               "Not possible to import in this workbook "
        Exit Sub
    End If
    If FolderWithVBAProjectFiles = "Error" Then
        MsgBox "Import Folder not exist"
        Exit Sub
    End If
    szImportPath = FolderWithVBAProjectFiles & "\"
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.getFolder(szImportPath).Files.count = 0 Then
        MsgBox "There are no files to import"
        Exit Sub
    End If
    Call DeleteVBAModulesAndUserForms(wkbTarget)
    Set cmpComponents = wkbTarget.VBProject.VBComponents
    For Each objFile In objFSO.getFolder(szImportPath).Files
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
                                                           (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
                                                           (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
    Next objFile
End Sub

Function FolderWithVBAProjectFiles() As String
    '#INCLUDE FolderExists
    Dim WshShell As Object
    Dim fso As Object
    Dim SpecialPath As String
    Set WshShell = CreateObject("WScript.Shell")
    Set fso = CreateObject("scripting.filesystemobject")
    SpecialPath = WshShell.SpecialFolders("MyDocuments")
    If Right(SpecialPath, 1) <> "\" Then
        SpecialPath = SpecialPath & "\"
    End If
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = False Then
        On Error Resume Next
        MkDir SpecialPath & "VBAProjectFiles"
        On Error GoTo 0
    End If
    If fso.FolderExists(SpecialPath & "VBAProjectFiles") = True Then
        FolderWithVBAProjectFiles = SpecialPath & "VBAProjectFiles"
    Else
        FolderWithVBAProjectFiles = "Error"
    End If
End Function

Sub DeleteVBAModulesAndUserForms(wkbSource As Workbook)
    Dim vbProj As VBIDE.VBProject
    Dim vbComp As VBIDE.VBComponent
    Set vbProj = wkbSource.VBProject
    For Each vbComp In vbProj.VBComponents
        If vbComp.Type = vbext_ct_Document Then
        Else
            vbProj.VBComponents.Remove vbComp
        End If
    Next vbComp
End Sub

Sub AutofitMergedCells(TargetRange As Range)
    '#INCLUDE StartOptimizeCodeRun
    '#INCLUDE StopOptimizeCodeRun
    Dim mw As Single
    Dim cM As Range
    Dim rng As Range
    Dim cw As Double
    Dim rwht As Double
    Dim ar As Variant
    Dim i As Integer
    StartOptimizeCodeRun
    Dim cell As Range
    Dim TimeStarted
    TimeStarted = Now()
    For Each cell In TargetRange
        If cell.MergeCells = True Then
            Set rng = Range(cell.MergeArea.Address)
            rng.MergeCells = False
            cw = rng.Cells(1).ColumnWidth
            mw = 0
            For Each cM In rng
                cM.WrapText = True
                mw = cM.ColumnWidth + mw
            Next
            mw = mw + rng.Cells.count * 0.66
            rng.Cells(1).ColumnWidth = mw
            rng.EntireRow.AutoFit
            rwht = rng.RowHeight
            rng.Cells(1).ColumnWidth = cw
            rng.MergeCells = True
            rng.RowHeight = rwht
        End If
        If Now > TimeStarted + TimeSerial(0, 1, 0) Then Stop
    Next
    StopOptimizeCodeRun
End Sub


