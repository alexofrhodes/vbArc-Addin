Attribute VB_Name = "F_FunctionsUnsorted"
Rem @Folder Unsorted Declarations
Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Enum myColors
    FormBackgroundDarkGray = 4208182        ' BACKGROUND DARK GRAY
    FormSidebarMediumGray = 5457992        ' TILE COLORS LIGHTER GRAY
    FormSidebarMouseOverLightGray = &H808080        ' lighter light gray
    FormSelectedGreen = 8435998        ' green tile
End Enum

Sub StartOptimizeCodeRun()
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
End Sub

Sub StopOptimizeCodeRun()
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = False
End Sub

Rem @Folder Unsorted
Function getLastRow(TargetSheet As Worksheet)
    '#INCLUDE LastCell
    Dim LastCell As Range
    On Error Resume Next
    Set LastCell = ActiveSheet.Cells.Find("*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If LastCell Is Nothing Then
        getLastRow = 1
    Else
        getLastRow = LastCell.row
    End If
End Function

Function getLastColumn(TargetSheet As Worksheet) As Long
    '#INCLUDE LastCell
    Dim LastCell As Range
    On Error Resume Next
    Set LastCell = ActiveSheet.Cells.Find("*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If LastCell Is Nothing Then
        getLastColumn = 1
    Else
        getLastColumn = LastCell.Column
    End If
End Function

Function GithubImageFormat(Optional GithubImageString As String) As String
    '#INCLUDE CLIP
    If GithubImageString = "" Then GithubImageString = CLIP
    Dim v
    v = Split(GithubImageString, Chr(10))
    Dim i As Long
    Dim s As String
    Dim tmp As String
    For i = LBound(v) To UBound(v)
        If left(Trim(v(i)), 1) = "!" Then
            v(i) = Trim(v(i))
            tmp = "<img src=" & Chr(34) & Mid(v(i), InStr(v(i), "(") + 1)
            tmp = left(tmp, InStrRev(tmp, ")") - 1)
            tmp = tmp & Chr(34) & " width=" & Chr(34) & "300" & Chr(34) & " height=" & Chr(34) & Chr(34) & ">"
            s = IIf(s = "", tmp, s & vbNewLine & tmp)
        Else
            s = IIf(s = "", v(i), s & vbNewLine & v(i))
        End If
    Next
    GithubImageFormat = s
    CLIP s
End Function

Sub OpenValidationComboboxOnClick(ByVal Target As Range)
    If TypeName(Selection) <> "Range" Then Exit Sub
    Dim lngValType As Long
    On Error Resume Next
    lngValType = Target.Validation.Type
    On Error GoTo 0
    If lngValType = 3 Then SendKeys "%{DOWN}", True
End Sub

Sub BackupActiveWorkbook()
    '#INCLUDE BackupWorkbook
    BackupWorkbook ActiveWorkbook
End Sub

Sub BackupActiveCodepaneWorkbook()
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE BackupWorkbook
    BackupWorkbook ActiveCodepaneWorkbook
End Sub

Sub BackupWorkbook(TargetWorkbook As Workbook)
    '#INCLUDE GetProjectText
    '#INCLUDE FollowLink
    '#INCLUDE FoldersCreate
    '#INCLUDE TxtOverwrite
    workbookName = left(TargetWorkbook.Name, InStr(1, TargetWorkbook.Name, ".") - 1)
    Dim exportPath As String
    exportPath = Environ("USERprofile") & "\My Documents\vbArc\Backups\" & workbookName
    FoldersCreate exportPath
    TargetWorkbook.SaveCopyAs _
        fileName:=exportPath & "\" & _
                   Format(Now, "YY-MM-DD HHNNSS") & " " & _
                   TargetWorkbook.Name
    TxtOverwrite exportPath & "\" & Format(Now, "YY-MM-DD HHNNSS") & " " & workbookName & ".txt", _
                                                                   GetProjectText(TargetWorkbook)
    FollowLink exportPath
End Sub

Function LargestLength(Optional myObj) As Long
    LargestLength = 0
    Dim element As Variant
    If IsMissing(myObj) Then
        If TypeName(Selection) = "Range" Then
            Set myObj = Selection
        Else
            Exit Function
        End If
    End If
    Select Case TypeName(myObj)
        Case Is = "String"
            LargestLength = Len(myObj)
        Case "Collection"
            For Each element In myObj
                If Len(element) > LargestLength Then LargestLength = Len(element)
            Next element
        Case "Variant", "Variant()", "String()"
            For element = LBound(myObj) To UBound(myObj)
                If Len(myObj(element)) > LargestLength Then LargestLength = Len(myObj(element))
            Next
        Case Else
    End Select
End Function

Sub ListContainedProceduresInTXT(fileName As String)
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

Function ExcludeRange(ByVal rngMain As Range, rngExc As Range) As Range
    Dim rngTemp     As Range
    Dim rng         As Range
    Set rngTemp = rngMain
    Set rngMain = Nothing
    For Each rng In rngTemp
        If Intersect(rng, rngExc) Is Nothing Then
            If rngMain Is Nothing Then
                Set rngMain = rng
            Else
                Set rngMain = Union(rngMain, rng)
            End If
        End If
    Next
    Set ExcludeRange = rngMain
End Function

Function CellRow(cell As Range) As Range
    Dim ws As Worksheet
    Set ws = cell.parent
    Dim r: r = cell.row
    Dim c As Long: c = cell.CurrentRegion.Column
    Set CellRow = ws.Range(ws.Cells(r, c), ws.Cells(r, c + cell.CurrentRegion.Columns.count - 1))
End Function

Function RangeToString(ByVal myRange As Range, Optional delim As String = ",") As String
    RangeToString = ""
    If Not myRange Is Nothing Then
        Dim myCell As Range
        For Each myCell In myRange
            RangeToString = RangeToString & delim & myCell.Value
        Next myCell
        RangeToString = Right(RangeToString, Len(RangeToString) - Len(delim))
    End If
End Function

Function ContainsIllegalCharacter(strIn As String) As Boolean
    Dim strSpecialChars As String: strSpecialChars = "~""#%@&*:<>?!{|}/\[]" & Chr(10) & Chr(13)
    Dim strOut As String: strOut = strIn
    Dim i As Long
    For i = 1 To Len(strSpecialChars)
        strOut = Replace(strOut, Mid$(strSpecialChars, i, 1), "")
        If Len(strIn) <> Len(strOut) Then ContainsIllegalCharacter = True: Exit Function
    Next
End Function

Function Translate(SourceText As String, SourceLanguage As String, TargetLanguage As String) As String
    Dim URL As String, objHTTP As Object, objHTML As Object, allDivs As Object, div As Variant
    txtTgt = "N/A"
    URL = "https://translate.google.com/m?hl=" & SourceLanguage & _
          "&sl=" & SourceLanguage & _
          "&tl=" & TargetLanguage & _
          "&ie=UTF-8&prev=_m&q=" & SourceText
    Set objHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    objHTTP.Open "GET", URL, False
    objHTTP.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    objHTTP.send ""
    Set objHTML = CreateObject("htmlfile")
    With objHTML
        .Open
        .Write objHTTP.responseText
        .Close
    End With
    Rem Browse through the HTML virtual file to find the proper <div>
    Rem Currently, the <div> with the translation can be identified by its class name "result-container"
    Rem It
    Set allDivs = objHTML.getElementsByTagName("div")
    For Each div In allDivs
        If div.ClassName = "result-container" Then
            Rem Found
            Translate = div.innerText
            Exit For
        End If
    Next
    Set allDivs = Nothing
    Set objHTTP = Nothing
    Set objHTML = Nothing
End Function

Function AddShape() As Shape
    Dim shp As Shape
    Set shp = ActiveSheet.Shapes.AddShape _
              (msoShapeRoundedRectangle, 1, 1, 500, 10)
    With shp.ThreeD
        .BevelTopType = msoBevelCircle
        .BevelTopInset = 6
        .BevelTopDepth = 6
    End With
    With shp.Fill
        .visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With shp.line
        .visible = msoTrue
        .ForeColor.ObjectThemeColor = msoThemeColorBackground1
        .ForeColor.TintAndShade = 0
        .ForeColor.Brightness = 0
        .Transparency = 0
    End With
    Set AddShape = shp
End Function

Public Function CLIP(Optional StoreText As String) As String
    Dim X As Variant
    X = StoreText
    With CreateObject("htmlfile")
        With .parentWindow.clipboardData
            Select Case True
                Case Len(StoreText)
                    .SetData "text", X
                Case Else
                    CLIP = .GetData("text")
            End Select
        End With
    End With
End Function

Function OutlookCheck() As Boolean
    Dim xOLApp As Object
    Set xOLApp = CreateObject("Outlook.Application")
    If Not xOLApp Is Nothing Then
        OutlookCheck = True
        Set xOLApp = Nothing
        Exit Function
    End If
    OutlookCheck = False
End Function

Function InputboxString(Optional sTitle As String = "Select String", Optional sPrompt As String = "Select String", Optional DefaultValue = "") As String
    Dim stringVariable As String
    stringVariable = Application.InputBox( _
                     title:=sTitle, _
                     Prompt:=sPrompt, _
                     Type:=2, _
                     Default:=DefaultValue)
    InputboxString = CStr(stringVariable)
End Function

Function LastCell(rng As Range, Optional booCol As Boolean, Optional onlyAfterFirstCell As Boolean) As Range
    Dim ws As Worksheet
    Set ws = rng.parent
    Dim cell As Range
    If booCol = False Then
        Set cell = ws.Cells(rows.count, rng.Column).End(xlUp)
        If cell.MergeCells Then Set cell = Cells(cell.row + cell.rows.count - 1, cell.Column)
    Else
        Set cell = ws.Cells(rng.row, Columns.count).End(xlToLeft)
        If cell.MergeCells Then Set cell = Cells(cell.row, cell.Column + cell.Columns.count - 1)
    End If
    If onlyAfterFirstCell = True Then
        If booCol = False Then
            Do While cell.row <= rng.row
                Set cell = cell.OFFSET(1, 0)
            Loop
        Else
            Do While cell.Column <= rng.Column
                Set cell = cell.OFFSET(0, 1)
            Loop
        End If
    End If
    Set LastCell = cell
End Function

Public Sub MsgPOP(message As String, Optional CloseAfterSeconds As Long = 1)
    CreateObject("WScript.Shell").PopUp message, CloseAfterSeconds
End Sub

Public Function PadRight(ByVal str As String, ByVal Length As Long, Optional Character As String = " ", Optional removeExcess As Boolean)
    If Len(str) < Length Then
        PadRight = str & String$(Length - Len(str), Character)
    ElseIf removeExcess = True Then
        PadRight = left$(str, Length)
    Else
        PadRight = str
    End If
End Function

Function RangeFindNth(rng As Range, strText As String, occurence As Integer) As Range
    Dim c As Range
    Dim counter As Integer
    For Each c In rng
        If InStr(1, c, strText) > 0 Then counter = counter + 1
        If counter = occurence Then
            Set RangeFindNth = c
            Exit Function
        End If
    Next c
End Function

Function WorkbookIsOpen(ByVal sWbkName As String) As Boolean
    WorkbookIsOpen = False
    On Error Resume Next
    WorkbookIsOpen = Len(Workbooks(sWbkName).Name) <> 0
    On Error GoTo 0
End Function

Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet
    If wb Is Nothing Then
        Set wb = ThisWorkbook
    End If
    On Error Resume Next
    Set sht = wb.SHEETS(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function

Function CreateOrSetSheet(SheetName As String, TargetWorkbook As Workbook) As Worksheet
    '#INCLUDE WorksheetExists
    Dim NewSheet As Worksheet
    If WorksheetExists(SheetName, TargetWorkbook) = True Then
        Set CreateOrSetSheet = TargetWorkbook.SHEETS(SheetName)
    Else
        Set CreateOrSetSheet = TargetWorkbook.SHEETS.Add
        CreateOrSetSheet.Name = SheetName
    End If
End Function

Sub FollowLink(FolderPath As String)
    If Right(FolderPath, 1) = "\" Then FolderPath = left(FolderPath, Len(FolderPath) - 1)
    On Error Resume Next
    Dim oShell As Object
    Dim Wnd As Object
    Set oShell = CreateObject("Shell.Application")
    For Each Wnd In oShell.Windows
        If Wnd.Name = "File Explorer" Then
            If Wnd.document.Folder.Self.Path = FolderPath Then Exit Sub
        End If
    Next Wnd
    Application.ThisWorkbook.FollowHyperlink Address:=FolderPath, NewWindow:=True
End Sub

Sub SortColumns(rng As Range)
    '#INCLUDE ArrayToRange2D
    '#INCLUDE CreateOrSetSheet
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("TEMP SORT", ThisWorkbook)
    Dim arr As Variant
    arr = rng
    ArrayToRange2D WorksheetFunction.Transpose(arr), ws.Range("A1")
    ws.UsedRange.Sort (ws.UsedRange.Cells(1, 1))
    rng.ClearContents
    rng = WorksheetFunction.Transpose(ws.UsedRange)
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub


