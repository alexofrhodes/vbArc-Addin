Attribute VB_Name = "F_VbeFormat"

Rem @Folder Format Declarations
Public Enum ProcedureScope
    PRIVATE_SCOPE = 1
    Public_SCOPE = 2
    FRIEND_SCOPE = 3
    DEFAULT_SCOPE = 4
End Enum

Public Enum LineSplits
    LineSplitRemove = 0
    LineSplitKeep = 1
    LineSplitConvert = 2
End Enum

Public Type ProcInfo
    procName As String
    ProcKind As VBIDE.vbext_ProcKind
    ProcStartLine As Long
    ProcBodyLine As Long
    ProcCountLines As Long
    ProcedureScope As ProcedureScope
    ProcDeclaration As String
End Type

Rem @Folder FormatVBATools Declarations
Private Const vbTab2 = vbTab & vbTab
Private Const vbTab4 = vbTab2 & vbTab2
Private Const ctFormat = "dd-mm-yyyy hh:nn"

Rem @Folder Format
Sub FormatVBA7()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    '#INCLUDE CLIP
    '#INCLUDE collectionToString
    '#INCLUDE SortCollection
    Dim selectedText
    selectedText = ActiveModule.CodeModule.Lines(CodePaneSelectionStartLine, CodePaneSelectionEndLine - CodePaneSelectionStartLine + 1)
    selectedText = Split(selectedText, vbNewLine)
    Dim IsVba7 As String
    Dim NotVba7 As String
    Dim colIsVBA7 As New Collection
    Dim colNotVBA7 As New Collection
    Dim i As Long
    For i = LBound(selectedText) To UBound(selectedText)
        If InStr(1, selectedText(i), "PtrSafe", vbTextCompare) Then
            IsVba7 = selectedText(i)
            NotVba7 = Replace(selectedText(i), "Declare ptrsafe ", "Declare ", , , vbTextCompare)
        Else
            IsVba7 = Replace(selectedText(i), "Declare ", "Declare PtrSafe ")
            NotVba7 = selectedText(i)
        End If
        colIsVBA7.Add IsVba7
        colNotVBA7.Add NotVba7
    Next
    Set colIsVBA7 = SortCollection(colIsVBA7)
    Set colNotVBA7 = SortCollection(colNotVBA7)
    Dim out As String
    out = "#If VBA7 then" & vbNewLine & _
          collectionToString(colIsVBA7, vbNewLine) & vbNewLine & _
          "#Else" & vbNewLine & _
          collectionToString(colNotVBA7, vbNewLine) & vbNewLine & _
          "#End If"
    CLIP out
    MsgBox "copied to clipboard"
End Sub

Sub AlignVbeComments()
    '#INCLUDE AlignCodepaneLineElements
    AlignCodepaneLineElements "'"
End Sub

Sub AlignDimAs()
    '#INCLUDE AlignCodepaneLineElements
    AlignCodepaneLineElements "As"
End Sub

Sub AlignCodepaneLineElements(AlignString As String, Optional AlignAtColumn As Long)
    '#INCLUDE Inject
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE CodePaneSelectionSet
    '#INCLUDE ActiveModule
    Dim firstRow As Long
    firstRow = CodePaneSelectionStartLine
    Dim LastRow As Long
    LastRow = CodePaneSelectionEndLine
    Dim elementOriginalColumn As Long
    Dim LineText As String
    Dim i As Long
    Dim rightMostColumn As Long
    For i = firstRow To LastRow
        LineText = ActiveModule.CodeModule.Lines(i, 1)
        elementOriginalColumn = InStrRev(LineText, AlignString)
        If elementOriginalColumn > rightMostColumn Then rightMostColumn = elementOriginalColumn
    Next
    If AlignAtColumn = 0 Then AlignAtColumn = rightMostColumn
    Dim numberOfSpacesToInsert
    For i = firstRow To LastRow
        LineText = ActiveModule.CodeModule.Lines(i, 1)
        elementOriginalColumn = InStr(1, LineText, AlignString)
        If elementOriginalColumn > 0 Then
            numberOfSpacesToInsert = AlignAtColumn - elementOriginalColumn
            If numberOfSpacesToInsert > 0 Then
                elementOriginalColumn = InStrRev(LineText, AlignString)
                CodePaneSelectionSet i, elementOriginalColumn, i, elementOriginalColumn
                Inject Space(numberOfSpacesToInsert)
            End If
        End If
    Next
End Sub

Sub InsertTimerToProcedure(Optional TargetWorkbook As Workbook, Optional Procedure As String)
    '#INCLUDE InsertStringToProcedureBody
    '#INCLUDE InsertStringToProcedureEnd
    '#INCLUDE StartTimer
    '#INCLUDE EndTimer
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then Procedure = ActiveProcedure
    Dim ProcedureText As String
    ProcedureText = GetProcText(ModuleOfProcedure(TargetWorkbook, Procedure), Procedure)
    If InStr(1, ProcedureText, "StartTimer") > 0 Then Exit Sub
    InsertStringToProcedureBody TargetWorkbook, CStr(Procedure), "StartTimer"
    Sleep 100
    InsertStringToProcedureEnd TargetWorkbook, CStr(Procedure), "EndTimer"
End Sub

Sub InsertStringToProcedureStart(TargetWorkbook As Workbook, Procedure As String, AddThis As String, Optional SkipIfExists As Boolean = True)
    Rem as comment before Procedure declaration
    '#INCLUDE ProcedureStartLine
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    Dim Module As VBComponent: Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    If SkipIfExists = True Then
        If InStr(1, GetProcText(Module, Procedure), AddThis) > 0 Then Exit Sub
    End If
    Module.CodeModule.InsertLines ProcedureStartLine(Module, Procedure), vbNewLine & "'" & AddThis
End Sub

Sub InsertStringToProcedureBody(TargetWorkbook As Workbook, Procedure As String, AddThis As String, Optional SkipIfExists As Boolean = True)
    '#INCLUDE ProcedureFirstLine
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    Dim Module As VBComponent: Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    If SkipIfExists = True Then
        If InStr(1, GetProcText(Module, CStr(Procedure)), AddThis) > 0 Then Exit Sub
    End If
    Module.CodeModule.InsertLines ProcedureFirstLine(Module, Procedure), AddThis
End Sub

Sub InsertStringToProcedureEnd(TargetWorkbook As Workbook, Procedure As String, AddThis As String, Optional SkipIfExists As Boolean = True)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    Dim Module As VBComponent: Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    If SkipIfExists = True Then
        If InStr(1, GetProcText(Module, CStr(Procedure)), AddThis) > 0 Then Exit Sub
    End If
    Module.CodeModule.InsertLines ProcedureEndLine(Module, Procedure), AddThis
End Sub

Sub CaseProperModulesOfWorkbook(Optional TargetWorkbook As Workbook)
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Name <> "ThisWorkbook" Then
            Module.Name = UCase(left(Module.Name, 1)) & Mid(Module.Name, 2)
        End If
    Next
End Sub

Sub MODULEINFO(Optional Module As VBComponent)
    '#INCLUDE ProcListCollection
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    s = s & vbNewLine
    s = s & vbNewLine & "'" & "Module Name: " & Module.Name & "'"
    s = s & vbNewLine & "'" & "Procedures Count: " & ProcListCollection(Module).count
    s = s & vbNewLine & "'" & "Lines: " & Module.CodeModule.CountOfLines & " of which " & Module.CodeModule.CountOfDeclarationLines & " are declaration or comments at the top"
    s = s & vbNewLine
    Module.CodeModule.InsertLines 1, s
End Sub

Sub InjectDevInfo()
    '#INCLUDE DevInfo
    '#INCLUDE Inject
    '#INCLUDE CodepaneSelection
    If Len(CodepaneSelection) = 0 Then Inject DevInfo
End Sub

Function DevInfo() As String
    '#INCLUDE DpHeader
    '#INCLUDE CLIP
    Dim i As Long: i = 14
    Dim Character As String: Character = "_"
    DevInfo = DpHeader(Array( _
                       "AUTHOR     " & AUTHOR_NAME, _
                       "EMAIL      " & AUTHOR_EMAIL, _
                       "GITHUB     " & AUTHOR_GITHUB, _
                       "YOUTUBE    " & AUTHOR_YOUTUBE, _
                       "VK         " & AUTHOR_VK) _
                       , , , True, True)
    CLIP DevInfo
End Function

Sub CreateShapeButtonToRunProcedure(Optional Procedure As String = "")
    '#INCLUDE ActiveProcedure
    '#INCLUDE AddShape
    If TypeName(Selection) <> "Range" Then
        MsgBox "Select cell to contain the shape"
        Exit Sub
    End If
    If Procedure = "" Then Procedure = ActiveProcedure
    With AddShape
        .OnAction = Procedure
        .Name = "ProcButton_" & Procedure
        .TextFrame2.TextRange.TEXT = Procedure
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.AutoSize = msoAutoSizeShapeToFitText
        .left = Selection.left
        .top = Selection.top
    End With
End Sub

Public Sub RemoveBlankLinesFromModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long, s As String
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            s = .Lines(n, 1)
            If Trim(s) = vbNullString Then
                If n > 1 Then
                    If InStr(1, Module.CodeModule.Lines(n - 1, 1), "End Sub") = 0 _
                                                                                Or InStr(1, Module.CodeModule.Lines(n - 1, 1), "End Function") = 0 Then
                        Module.CodeModule.DeleteLines n
                    End If
                End If
            ElseIf left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub

Public Sub RemoveBlankLinesFromActiveProcedure()
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim n As Long, s As String
    Dim Module As VBIDE.VBComponent: Set Module = ActiveModule
    Dim procName As String: procName = ActiveProcedure
    For n = ProcedureEndLine(Module, procName) To ProcedureStartLine(Module, procName) Step -1
        s = Module.CodeModule.Lines(n, 1)
        If Trim(s) = vbNullString Then
            If n > 1 Then
                If InStr(1, Module.CodeModule.Lines(n - 1, 1), "End Sub") = 0 Or _
                                                                            InStr(1, Module.CodeModule.Lines(n - 1, 1), "End Function") = 0 Then
                    Module.CodeModule.DeleteLines n
                End If
            End If
        ElseIf left(Trim(s), 1) = "'" Then
        Else
        End If
    Next
End Sub

Public Sub RemoveBlankLinesFromWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE RemoveBlankLinesFromModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        RemoveBlankLinesFromModule Module
    Next
End Sub

Public Sub CaseLower()
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & LCase(Code) & _
                    PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Sub CaseProper()
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & WorksheetFunction.Proper(Code) & _
                                       PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Sub CaseUpper()
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & UCase(Code) & _
                    PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Sub RemoveEmptyModules(Optional TargetWorkbook As Workbook)
    '#INCLUDE ProcListCollection
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            If ProcListCollection(Module).count = 0 And Module.CodeModule.CountOfLines < 3 Then TargetWorkbook.VBProject.VBComponents.Remove Module
        End If
    Next
End Sub

Sub MoveProcedureToOtherModule(ProcedureName As String, FromWorkbook As Workbook, TargetModule As VBComponent)
    '#INCLUDE ModuleOfProcedure
    Dim Module As VBComponent
    Set Module = ModuleOfProcedure(FromWorkbook, ProcedureName)
    With Module.CodeModule
        c00 = .Lines(.ProcStartLine(ProcedureName, 0), .ProcCountLines(ProcedureName, 0))
        .DeleteLines .ProcStartLine(ProcedureName, 0), .ProcCountLines(ProcedureName, 0)
    End With
    TargetModule.CodeModule.AddFromString c00
End Sub

Sub MoveModuleTextToOtherModule(FromModule As VBComponent, TargetModule As VBComponent)
    Dim ModuleDeclarations As String
    Dim ModuleCode As String
    Dim counter As Long
    If FromModule.CodeModule.CountOfDeclarationLines > 0 Then
        ModuleDeclarations = "Rem @Folder " & FromModule.Name
        For counter = 1 To FromModule.CodeModule.CountOfDeclarationLines
            ModuleDeclarations = ModuleDeclarations & vbNewLine & FromModule.CodeModule.Lines(counter, 1)
        Next
    End If
    If FromModule.CodeModule.CountOfLines - FromModule.CodeModule.CountOfDeclarationLines > 0 Then
        For counter = FromModule.CodeModule.CountOfDeclarationLines + 1 To FromModule.CodeModule.CountOfLines
            ModuleCode = ModuleCode & vbNewLine & FromModule.CodeModule.Lines(counter, 1)
        Next
    End If
    With TargetModule.CodeModule
        .InsertLines 1, ModuleDeclarations
        .InsertLines .CountOfLines + 1, ModuleCode
    End With
    With FromModule.CodeModule
        If .CountOfLines > 0 Then
            For counter = 1 To .CountOfLines
                .ReplaceLine counter, "' " & .Lines(counter, 1)
            Next
        End If
    End With
End Sub

Sub MergeModules( _
    FromWorkbook As Workbook, _
    TargetModule As VBComponent, _
    Optional OnlyTheseModules As Variant)
    '#INCLUDE MoveModuleTextToOtherModule
    Dim Module As VBComponent
    If OnlyTheseModules Is Nothing Then
        For Each Module In FromWorkbook.VBProject.VBComponents
            If Module.Type = vbext_ct_StdModule Then
                If Module.Name <> TargetModule.Name Then MoveModuleTextToOtherModule Module, TargetModule
            End If
        Next
    Else
        For Each Module In OnlyTheseModules
            MoveModuleTextToOtherModule Module, TargetModule
        Next
    End If
End Sub

Sub MoveProceduresFromAllUserformsToModules(Optional TargetWorkbook As Workbook)
    '#INCLUDE MoveProceduresFromUserformToModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_MSForm Then
            MoveProceduresFromUserformToModule Module
        End If
    Next
End Sub

Sub MoveProceduresFromUserformToModule(Optional form As VBComponent)
    '#INCLUDE ProcListCollection
    '#INCLUDE ActiveModule
    '#INCLUDE CreateOrSetModule
    '#INCLUDE GetProcText
    '#INCLUDE WorkbookOfModule
    If form Is Nothing Then Set form = ActiveModule
    Rem Procedures without underscore "_" can be moved to a module
    Rem possible error if the procedures rely on const or enums or variables contained in form codemodule
    If form.Type <> vbext_ct_MSForm Then Exit Sub
    Dim TargetModule As VBComponent
    Set TargetModule = CreateOrSetModule("m" & form.Name, vbext_ct_StdModule, WorkbookOfModule(form))
    Dim strProc As String, StartLine As Long, totalLines As Long
    Dim Procedure As Variant
    Dim procedures As New Collection
    Set procedures = ProcListCollection(form)
    For Each Procedure In procedures
        If InStr(1, Procedure, "_") > 0 Then
        Else
            strProc = GetProcText(form, CStr(Procedure))
            StartLine = form.CodeModule.ProcStartLine(CStr(Procedure), vbext_pk_Proc)
            totalLines = form.CodeModule.ProcCountLines(CStr(Procedure), vbext_pk_Proc)
            TargetModule.CodeModule.AddFromString strProc
            form.CodeModule.DeleteLines StartLine, totalLines
        End If
    Next Procedure
End Sub

Sub RemoveLinesFromWorkbook(Optional TargetWorkbook As Workbook, Optional ContainsThis As String)
    '#INCLUDE RemoveLinesFromModule
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE InputboxString
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If ContainsThis = "" Then ContainsThis = InputboxString("Delete lines from " & TargetWorkbook.Name, _
                                                            "Delete lines containing what text?")
    If ContainsThis = "" Then Exit Sub
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        RemoveLinesFromModule Module, ContainsThis
    Next
End Sub

Public Sub RemoveLinesFromModule(Optional Module As VBComponent, Optional ContainsThis As String)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    With Module.CodeModule
        For j = .CountOfLines To 1 Step -1
            LineText = Trim(.Lines(j, 1))
            If InStr(1, LineText, ContainsThis, vbTextCompare) > 0 Then
                .DeleteLines j, 1
            End If
        Next
    End With
End Sub

Public Sub RemoveCommentsFromModule(Optional Module As VBComponent, Optional RemoveRem As Boolean)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n               As Long
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim exitString      As String
    Dim QUOTES          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    With Module.CodeModule
        For j = .CountOfLines To 1 Step -1
            LineText = Trim(.Lines(j, 1))
            If LineText = "ExitString = " & _
               """" & "Ignore Comments In This Module" & """" Then
                Exit For
            End If
            StartPos = 1
retry:
            n = InStr(StartPos, LineText, "'")
            Q = InStr(StartPos, LineText, """")
            QUOTES = 0
            If Q < n Then
                For l = 1 To n
                    If Mid(LineText, l, 1) = """" Then
                        QUOTES = QUOTES + 1
                    End If
                Next l
            End If
            If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
                StartPos = n + 1
                GoTo retry:
            Else
                Select Case n
                    Case Is = 0
                    Case Is = 1
                        .DeleteLines j, 1
                    Case Is > 1
                        .ReplaceLine j, left(LineText, n - 1)
                End Select
                If RemoveRem Then
                    If left(LineText, 4) = "Rem " Then .ReplaceLine j, " "
                End If
            End If
        Next j
    End With
    exitString = "Ignore Comments In This Module"
End Sub

Public Sub RemoveCommentsFromActiveProcedure(Optional RemoveRem As Boolean)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim n               As Long
    Dim i               As Long
    Dim j               As Long
    Dim k               As Long
    Dim l               As Long
    Dim LineText        As String
    Dim exitString      As String
    Dim QUOTES          As Long
    Dim Q               As Long
    Dim StartPos        As Long
    Dim StartLine As Long
    StartLine = ProcedureStartLine(Module, ActiveProcedure)
    Dim EndLine As Long
    EndLine = ProcedureEndLine(Module, ActiveProcedure)
    With Module.CodeModule
        For j = EndLine To StartLine Step -1
            LineText = Trim(.Lines(j, 1))
            If LineText = "ExitString = " & _
               """" & "Ignore Comments In This Module" & """" Then
                Exit For
            End If
            StartPos = 1
retry:
            n = InStr(StartPos, LineText, "'")
            Q = InStr(StartPos, LineText, """")
            QUOTES = 0
            If Q < n Then
                For l = 1 To n
                    If Mid(LineText, l, 1) = """" Then
                        QUOTES = QUOTES + 1
                    End If
                Next l
            End If
            If QUOTES = Application.WorksheetFunction.Odd(QUOTES) Then
                StartPos = n + 1
                GoTo retry:
            Else
                Select Case n
                    Case Is = 0
                    Case Is = 1
                        .DeleteLines j, 1
                    Case Is > 1
                        .ReplaceLine j, left(LineText, n - 1)
                End Select
                If RemoveRem Then
                    If left(LineText, 4) = "Rem " Then .ReplaceLine j, " "
                End If
            End If
        Next j
    End With
    exitString = "Ignore Comments In This Module"
End Sub

Public Sub RemoveCommentsFromWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE RemoveCommentsFromModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBIDE.VBComponent
    For Each Module In ActiveCodepaneWorkbook.VBProject.VBComponents
        RemoveCommentsFromModule Module
    Next
End Sub

Public Sub ReplaceQuoteWithRemInProcedure(Optional Module As VBComponent, Optional procName As String)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    If Module Is Nothing Then Set Module = ActiveModule
    If procName = "" Then procName = ActiveProcedure
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = ProcedureEndLine(Module, ActiveProcedure) To ProcedureStartLine(Module, ActiveProcedure) Step -1
            s = .Lines(n, 1)
            If left(Trim(s), 1) = "'" Then
                .ReplaceLine n, Replace(s, "'", "Rem ", , 1)
            End If
        Next n
    End With
End Sub

Public Sub ReplaceQuoteWithRemInModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If left(Trim(s), 1) = "'" Then
                .ReplaceLine n, Replace(s, "'", "Rem ", , 1)
            End If
        Next n
    End With
End Sub

Public Sub ReplaceQuoteWithRemInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE ReplaceQuoteWithRemInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        ReplaceQuoteWithRemInModule vbComp
    Next
End Sub

Public Sub DisableDebugPrintInModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If left(Trim(s), 5) = "Debug" Then
                .ReplaceLine n, "'" & s
            End If
        Next n
    End With
End Sub

Public Sub DisableDebugPrintInProcedure(Optional Module As VBComponent, Optional procName As String)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    If Module Is Nothing Then Set Module = ActiveModule
    If procName = "" Then procName = ActiveProcedure
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = ProcedureEndLine(Module, ActiveProcedure) To ProcedureStartLine(Module, ActiveProcedure) Step -1
            s = .Lines(n, 1)
            If left(Trim(s), 5) = "Debug" Then
                .ReplaceLine n, "'" & s
            End If
        Next n
    End With
End Sub

Public Sub DisableDebugPrintInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE DisableDebugPrintInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        DisableDebugPrintInModule Module
    Next
End Sub

Public Sub EnableDebugPrintInModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If left(Trim(s), 6) = "'Debug" Then
                s = Replace(s, "'", "", , 1)
                .ReplaceLine n, s
            ElseIf left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub

Public Sub EnableDebugPrintInProcedure(Optional Module As VBComponent, Optional procName As String)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    If Module Is Nothing Then Set Module = ActiveModule
    If procName = "" Then procName = ActiveProcedure
    Dim n As Long
    Dim s As String
    With Module.CodeModule
        For n = ProcedureEndLine(Module, ActiveProcedure) To ProcedureStartLine(Module, procName) Step -1
            s = .Lines(n, 1)
            If left(Trim(s), 6) = "'Debug" Then
                s = Replace(s, "'", "", , 1)
                .ReplaceLine n, s
            ElseIf left(Trim(s), 1) = "'" Then
            Else
            End If
        Next n
    End With
End Sub

Public Sub EnableDebugPrintInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE EnableDebugPrintInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        EnableDebugPrintInModule Module
    Next
End Sub

Public Sub DisableStopInModule(Optional Module As VBComponent)
    '#INCLUDE InStrExact
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long
    Dim s As String
    Dim keyword As String
    keyword = "Stop"
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If InStrExact(1, s, keyword) > 0 Then
                .ReplaceLine n, "'" & s
            End If
        Next n
    End With
End Sub

Public Sub DisableStopInProcedure(Optional Module As VBComponent, Optional procName As String)
    '#INCLUDE InStrExact
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    If Module Is Nothing Then Set Module = ActiveModule
    If procName = "" Then procName = ActiveProcedure
    Dim n As Long
    Dim s As String
    Dim keyword As String
    keyword = "Stop"
    With Module.CodeModule
        For n = ProcedureEndLine(Module, ActiveProcedure) To ProcedureStartLine(Module, ActiveProcedure) Step -1
            s = .Lines(n, 1)
            If InStrExact(1, s, keyword) > 0 Then
                .ReplaceLine n, "'" & s
            End If
        Next n
    End With
End Sub

Public Sub DisableStopInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE DisableStopInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        DisableStopInModule Module
    Next
End Sub

Public Sub EnableStopInModule(Optional Module As VBComponent)
    '#INCLUDE InStrExact
    '#INCLUDE ActiveModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim n As Long
    Dim s As String
    Dim keyword As String
    keyword = "Stop"
    With Module.CodeModule
        For n = .CountOfLines To 1 Step -1
            If .CountOfLines = 0 Then Exit For
            s = .Lines(n, 1)
            If InStrExact(1, s, keyword) > 0 Then
                s = Replace(s, "'", "", , 1)
                .ReplaceLine n, s
            End If
        Next n
    End With
End Sub

Public Sub EnableStopInProcedure(Optional Module As VBComponent, Optional procName As String)
    '#INCLUDE InStrExact
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    If Module Is Nothing Then Set Module = ActiveModule
    If procName = "" Then procName = ActiveProcedure
    Dim n As Long
    Dim s As String
    Dim keyword As String
    keyword = "Stop"
    With Module.CodeModule
        For n = ProcedureEndLine(Module, ActiveProcedure) To ProcedureStartLine(Module, procName) Step -1
            s = .Lines(n, 1)
            If InStrExact(1, s, keyword) > 0 Then
                s = Replace(s, "'", "", , 1)
                .ReplaceLine n, s
            End If
        Next n
    End With
End Sub

Public Sub EnableStopInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE EnableStopInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        EnableStopInModule Module
    Next
End Sub

Sub AssignEnumValues(Optional ToThePower As Boolean = True)
    '#INCLUDE CodepaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String: Code = CodepaneSelection
    Dim arr: arr = Split(Code, vbNewLine)
    Code = ""
    Dim out As String
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), "=") > 0 Then arr(i) = Split(arr(i), "=")(0)
        arr(i) = Space(4) & Trim(arr(i))
    Next
    If ToThePower = True Then
        For i = 0 To UBound(arr)
            out = arr(i) & "= 2 ^ " & i
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        Next
    Else
        For i = 0 To UBound(arr)
            out = arr(i) & "= " & i + 1
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        Next
    End If
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Rem Sub ReplaceCodePaneSelection(TargetString As String)
Rem     Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
Rem     Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
Rem End Sub
Sub EncapsulateQuotes()
    '#INCLUDE EncapsulateCodepaneSelection
    EncapsulateCodepaneSelection Chr(34), Chr(34)
End Sub

Sub EncapsulateParenthesis()
    '#INCLUDE EncapsulateCodepaneSelection
    EncapsulateCodepaneSelection "(", ")"
End Sub

Public Sub EncapsulateCodepaneSelection(Before As String, After As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Before & Code & After & _
        PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Sub EncapsulateMultipleCommaSeparatedWithQuotes()
    '#INCLUDE EncapsulateMultipleInCodepaneSelection
    EncapsulateMultipleInCodepaneSelection Chr(34), Chr(34), ","
End Sub

Sub EncapsulateMultipleCommaSeparatedWithParenthesis()
    '#INCLUDE EncapsulateMultipleInCodepaneSelection
    EncapsulateMultipleInCodepaneSelection "(", ")", ","
End Sub

Sub EncapsulateMultipleLinesWithParenthesis()
    '#INCLUDE EncapsulateMultipleInCodepaneSelection
    EncapsulateMultipleInCodepaneSelection "(", ")", ","
End Sub

Sub EncapsulateMultipleLinesWithQuotes()
    '#INCLUDE EncapsulateMultipleInCodepaneSelection
    EncapsulateMultipleInCodepaneSelection Chr(34), Chr(34), vbNewLine
End Sub

Public Sub EncapsulateMultipleInCodepaneSelection(LeftCapsule As String, RightCapsule As String, Splitter As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Dim arr
    arr = Split(Code, Splitter)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = LeftCapsule & arr(counter) & RightCapsule
    Next
    Code = Join(arr, Splitter)
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Code & _
        PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Sub UnFoldLine(Optional Splitter As String = "_" & vbNewLine, Optional joiner As String = " ")
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Join(Split(Code, Splitter), joiner) & _
                                            PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Sub FoldLine(Optional Splitter As String = ",", Optional joiner As String = ", _" & vbNewLine)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Join(Split(Code, Splitter), joiner) & _
                                            PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Function FirstDigit(ByVal strData As String) As Integer
    Dim RE As Object
    Dim REMatches As Object
    Set RE = CreateObject("vbscript.regexp")
    RE.Pattern = "[0-9]"
    Set REMatches = RE.Execute(strData)
    FirstDigit = REMatches(0).FirstIndex + 1
End Function

Sub FlipNewline()
    '#INCLUDE FLIP
    FLIP vbNewLine
End Sub

Sub FlipEqualSign()
    '#INCLUDE FLIP
    FLIP "="
End Sub

Sub FlipComma()
    '#INCLUDE FLIP
    FLIP ","
End Sub

Public Sub FLIP(delim As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) & _
                                                                                   Split(Code, delim)(1) & delim & Split(Code, delim)(0) & _
                                                                                   PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Sub FlipMultipleEqualSignOnMultipleLines()
    '#INCLUDE FlipMultiple
    FlipMultiple "=", vbNewLine
End Sub

Public Sub FlipMultiple(flipper As String, Optional Splitter)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    '#INCLUDE ArrayRemoveEmptyElemets
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Dim arr As Variant
    arr = Split(Code, Splitter)
    arr = ArrayRemoveEmptyElemets(arr)
    Dim counter As Long
    For counter = LBound(arr) To UBound(arr) - IIf(Right(UBound(arr), Len(Splitter)) = Splitter, Len(Splitter), 0)
        arr(counter) = Split(arr(counter), flipper)(1) & flipper & Split(arr(counter), flipper)(0)
    Next
    Code = Join(arr, Splitter)
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Code & _
        PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Sub FlipRotateCommas()
    '#INCLUDE FlipRotate
    FlipRotate ","
End Sub

Sub FlipRotateLines()
    '#INCLUDE FlipRotate
    FlipRotate vbNewLine
End Sub

Public Sub FlipRotate(delim As String)
    Rem Rotate multiple  eg. a,b,c,d -> b,c,d,a
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    '#INCLUDE RotateArray
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = Join(RotateArray(Split(Code, delim)), delim)
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) & _
                                                                                   Code & _
                                                                                   PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Sub RemLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim i As Long
    For i = blockStart To blockEnd
        Module.CodeModule.ReplaceLine i, "Rem " & Module.CodeModule.Lines(i, 1)
    Next
End Sub

Sub UnRemLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim i As Long
    For i = blockStart To blockEnd
        With Module.CodeModule
            If left(Trim(.Lines(i, 1)), 4) = "Rem " Then
                .ReplaceLine i, Replace(.Lines(i, 1), "Rem ", "", , 1)
            End If
        End With
    Next
End Sub

Sub CommentLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim i As Long
    For i = blockStart To blockEnd
        Module.CodeModule.ReplaceLine i, "'" & Module.CodeModule.Lines(i, 1)
    Next
End Sub

Sub UnCommentLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim i As Long
    For i = blockStart To blockEnd
        With Module.CodeModule
            If left(Trim(.Lines(i, 1)), 1) = "'" Then
                .ReplaceLine i, Replace(.Lines(i, 1), "'", "", , 1)
            End If
        End With
    Next
End Sub

Sub LineDuplicate()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE ProcedureFirstLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim RowNumber As Long
    RowNumber = CodePaneSelectionStartLine
    Dim activeLine As Long
    activeLine = RowNumber
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim lineString As String
    Do While Len(Trim(ActiveModule.CodeModule.Lines(RowNumber, 1))) = 0 And RowNumber - 1 > ProcedureFirstLine(Module, ActiveProcedure)
        RowNumber = RowNumber - 1
    Loop
    If Len(Trim(ActiveModule.CodeModule.Lines(RowNumber, 1))) > 0 Then
        lineString = Module.CodeModule.Lines(RowNumber, 1)
    End If
    Module.CodeModule.InsertLines activeLine, lineString
End Sub

Sub LineIncrement()
    '#INCLUDE IncreaseAllNumbersInString
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE ProcedureFirstLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim RowNumber As Long
    RowNumber = CodePaneSelectionStartLine
    Dim activeLine As Long
    activeLine = RowNumber
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim lineString As String
    Do While Len(Trim(ActiveModule.CodeModule.Lines(RowNumber, 1))) = 0 And RowNumber - 1 >= ProcedureFirstLine(Module, ActiveProcedure)
        RowNumber = RowNumber - 1
    Loop
    If Len(Trim(ActiveModule.CodeModule.Lines(RowNumber, 1))) > 0 Then
        lineString = Module.CodeModule.Lines(RowNumber, 1)
    End If
    lineString = IncreaseAllNumbersInString(lineString)
    Module.CodeModule.InsertLines activeLine, lineString
End Sub

Sub CutLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    '#INCLUDE CLIP
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim blockCountOfLines As Long
    blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String
    blockString = Module.CodeModule.Lines(blockStart, blockCountOfLines)
    CLIP blockString
    Module.CodeModule.DeleteLines blockStart, blockCountOfLines
End Sub

Sub CopyLines()
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ActiveModule
    '#INCLUDE CLIP
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim blockCountOfLines As Long
    blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String
    blockString = Module.CodeModule.Lines(blockStart, blockCountOfLines)
    CLIP blockString
End Sub

Sub LinesMoveUp()
    '#INCLUDE LinesMove
    LinesMove moveUp:=True
End Sub

Sub LinesMoveDown()
    '#INCLUDE LinesMove
    LinesMove moveUp:=False
End Sub

Sub LinesMove(moveUp As Boolean)
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE CodePaneSelectionEndLine
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim blockStart As Long
    blockStart = CodePaneSelectionStartLine
    Dim blockEnd As Long
    blockEnd = CodePaneSelectionEndLine
    Dim blockCountOfLines As Long
    blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String
    blockString = Module.CodeModule.Lines(blockStart, blockCountOfLines)
    Dim insertBlockAtLine As Long
    Select Case moveUp
        Case True
            insertBlockAtLine = blockStart - 1
            Select Case insertBlockAtLine
                Case 1, ProcedureStartLine(Module, ActiveProcedure)
                    Exit Sub
            End Select
        Case False
            insertBlockAtLine = blockStart + 1
            Select Case insertBlockAtLine + blockCountOfLines - 1
                Case Module.CodeModule.CountOfLines, ProcedureEndLine(Module, ActiveProcedure)
                    Exit Sub
            End Select
    End Select
    Module.CodeModule.DeleteLines blockStart, blockCountOfLines
    Module.CodeModule.InsertLines insertBlockAtLine, blockString
    Module.CodeModule.CodePane.SetSelection insertBlockAtLine, 1, insertBlockAtLine + blockCountOfLines - 1, 300
End Sub

Sub MoveProcedureUp()
    '#INCLUDE MoveProcedure
    MoveProcedure moveUp:=True
End Sub

Sub MoveProcedureDown()
    '#INCLUDE MoveProcedure
    MoveProcedure moveUp:=False
End Sub

Sub MoveProcedure(moveUp As Boolean)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureStartLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim Procedure As String
    Procedure = ActiveProcedure
    Dim blockStart As Long
    blockStart = ProcedureStartLine(Module, Procedure)
    Dim blockEnd As Long
    blockEnd = ProcedureEndLine(Module, Procedure)
    Dim blockCountOfLines As Long
    blockCountOfLines = blockEnd - blockStart + 1
    Dim blockString As String
    blockString = Module.CodeModule.Lines(blockStart, blockCountOfLines)
    Dim NextProcedure As String
    Dim insertBlockAtLine As Long
    Dim i As Long
    Dim counter As Long
    On Error Resume Next
    Select Case moveUp
        Case Is = True
            For i = blockStart - 1 To 1 Step -1
                NextProcedure = Module.CodeModule.ProcOfLine(i, vbext_pk_Proc)
                If NextProcedure <> vbNullString Then Exit For
            Next
            insertBlockAtLine = ProcedureStartLine(Module, NextProcedure)
        Case Is = False
            For i = blockEnd + 1 To Module.CodeModule.CountOfLines
                NextProcedure = Module.CodeModule.ProcOfLine(i, vbext_pk_Proc)
                If NextProcedure <> vbNullString Then counter = counter + 1
                i = ProcedureEndLine(Module, NextProcedure) + 1
                If counter = 1 Then NextProcedure = vbNullString
                If counter = 2 Then Exit For
            Next
            If counter = 1 Then Exit Sub
            insertBlockAtLine = ProcedureStartLine(Module, NextProcedure)
    End Select
    On Error GoTo 0
    If moveUp = True Then
        Module.CodeModule.DeleteLines blockStart, blockCountOfLines
        Module.CodeModule.InsertLines insertBlockAtLine, blockString
    Else
        Module.CodeModule.InsertLines insertBlockAtLine, blockString
        Module.CodeModule.DeleteLines blockStart, blockCountOfLines
    End If
End Sub

Function IncreaseAllNumbersInString(str As String)
    Dim output As String
    Dim counter As Long
    counter = Len(str)
    Dim i As Long
    For i = 1 To Len(str)
        counter = i
        If IsNumeric(Mid(str, i, 1)) Then
            Do
                output = output & Mid(str, counter, 1)
                counter = counter + 1
            Loop While IsNumeric(Mid(str, counter, 1))
            i = counter - 1
            IncreaseAllNumbersInString = IncreaseAllNumbersInString & val(output + 1)
        Else
            output = output & Mid(str, i, 1)
            IncreaseAllNumbersInString = IncreaseAllNumbersInString & output
        End If
        output = ""
    Next
End Function

Public Function GetArguments( _
       Optional Procedure As String, _
       Optional TargetWorkbook As Workbook) _
        As String
    '#INCLUDE GetProcedureDeclaration
    '#INCLUDE CodepaneSelection
    '#INCLUDE ModReplaceMulti
    '#INCLUDE dp
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE CLIP
    Dim s As String
    s = CodepaneSelection
    If Procedure = "" Then
        If Len(s) > 0 Then
            Procedure = s
        Else
            Procedure = ActiveProcedure
        End If
    End If
    Dim Module As VBComponent
    If TargetWorkbook Is Nothing Then
        Set Module = ActiveModule
    Else
        Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    End If
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim str As Variant, element As Long, line As String
    Dim firstPart As String, secondPart As String, output As String
    str = GetProcedureDeclaration(Module, Procedure, ProcKind)
    If IsEmpty(str) Then GetArgs = "": Exit Function
    output = Procedure & "( _"
    Dim Indentation As String: Indentation = String(Len(output), " ")
    str = Right(str, Len(str) - InStr(1, str, "("))
    str = left(str, InStrRev(str, ")") - 1)
    If InStr(1, str, Chr(34) & "," & Chr(34)) > 0 Then str = Replace(str, Chr(34) & "," & Chr(34), Chr(34) & "###" & Chr(34))
    str = Split(str, ",")
    For i = LBound(str) To UBound(str)
        str(i) = Replace(str(i), Chr(34) & "###" & Chr(34), Chr(34) & "," & Chr(34))
    Next
    If UBound(str) = -1 Then Exit Function
    For element = LBound(str) To UBound(str)
        line = ModReplaceMulti(vbTextCompare, Trim(str(element)), "", "Optional ", "As ", "ByVal ", "ByRef", "ParamArray ", "_")
        firstPart = Split(line, " ")(0): secondPart = Split(line, " ")(1)
        output = output & vbNewLine & Indentation & firstPart & ":= " & "as" & secondPart & IIf(element <> UBound(str), ", _", ")")
    Next
    CLIP output
    dp output
    GetArguments = output
End Function

Public Function GetProcedureDeclaration(Module As VBComponent, _
                                        procName As String, _
                                        Optional LineSplitBehavior As LineSplits = LineSplitRemove)
    '#INCLUDE SingleSpace
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim s As String
    Dim Declaration As String
    On Error Resume Next
    LineNum = Module.CodeModule.ProcBodyLine(procName, ProcKind)
    If err.Number <> 0 Then
        Exit Function
    End If
    s = Module.CodeModule.Lines(LineNum, 1)
    Do While Right(s, 1) = "_"
        Select Case True
            Case LineSplitBehavior = LineSplitConvert
                s = left(s, Len(s) - 1) & vbNewLine
            Case LineSplitBehavior = LineSplitKeep
                s = s & vbNewLine
            Case LineSplitBehavior = LineSplitRemove
                s = left(s, Len(s) - 1) & " "
        End Select
        Declaration = Declaration & s
        LineNum = LineNum + 1
        s = Module.CodeModule.Lines(LineNum, 1)
    Loop
    Declaration = SingleSpace(Declaration & s)
    GetProcedureDeclaration = Declaration
End Function

Public Sub IndentModule(Optional vbComp As VBComponent)
    '#INCLUDE ActiveModule
    '#INCLUDE IsBlockEnd
    '#INCLUDE IsBlockStart
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    If vbComp.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    For nLine = 1 To vbComp.CodeModule.CountOfLines
        strNewLine = vbComp.CodeModule.Lines(nLine, 1)
        strNewLine = LTrim$(strNewLine)
        If IsBlockEnd(strNewLine) Then nIndent = nIndent - 1
        If nIndent < 0 Then nIndent = 0
        If strNewLine <> "" Then vbComp.CodeModule.ReplaceLine nLine, Space$(nIndent * 4) & strNewLine
        If IsBlockStart(strNewLine) Then nIndent = nIndent + 1
    Next nLine
End Sub

Public Sub IndentProcedure(Optional TargetWorkbook As Workbook, Optional ProcedureName As String)
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ProcedureType
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE IsBlockEnd
    '#INCLUDE IsBlockStart
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE WorkbookOfModule
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    Dim Module As VBComponent
    If ProcedureName = ActiveProcedure Then
        Set Module = ActiveModule
    Else
        Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    End If
    On Error GoTo eh
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim FirstLine As Long: FirstLine = Module.CodeModule.ProcStartLine(ProcedureName, vbext_pk_Proc)
    Dim EndLine As Long: EndLine = ProcedureEndLine(Module, ProcedureName)
    Rem = Module.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Proc)
    Rem = ProcedureType(WorkbookOfModule(Module), ProcedureName)
    Dim nIndent As Integer
    Dim nLine As Long
    Dim strNewLine As String
    For nLine = FirstLine To EndLine
        strNewLine = Module.CodeModule.Lines(nLine, 1)
        strNewLine = LTrim$(strNewLine)
        If IsBlockEnd(strNewLine) Then nIndent = nIndent - 1
        If nIndent < 0 Then nIndent = 0
        If strNewLine <> "" Then Module.CodeModule.ReplaceLine nLine, Space$(nIndent * 4) & strNewLine
        If IsBlockStart(strNewLine) Then nIndent = nIndent + 1
    Next nLine
eh:
End Sub

Public Sub IndentWorkbook(Optional wb As Workbook)
    '#INCLUDE IndentModule
    '#INCLUDE dp
    '#INCLUDE ActiveCodepaneWorkbook
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    dp wb.Name
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        IndentModule vbComp
    Next
End Sub

Function AddVariable(VariableName As String, Optional VariableType As String = "variant") As Boolean
    '#INCLUDE Inject
    If Len(VariableName) < 5 Then Debug.Print VariableName & " is not descriptive enough"
    Dim output As String
    VariableName = WorksheetFunction.Proper(VariableName)
    VariableName = Replace(VariableName, " ", "_")
    Select Case LCase(VariableType)
        Case "string"
            output = "Dim " & VariableName & " As String"
        Case "integer"
            output = "Dim " & VariableName & " As Integer"
        Case "variant"
            output = "Dim " & VariableName & " As Variant"
        Case "collection"
            output = "Dim " & VariableName & " As Collection"
            output = output & vbNewLine & "Set Output = New Collection"
        Case "boolean"
            output = "Dim " & VariableName & " As Boolean"
        Case Else
            AddVariable = False
            Exit Function
    End Select
    Inject output
    AddVariable = True
End Function

Public Sub Inject(str As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    If Len(Code) > 0 Then Exit Sub
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & str & _
        PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    On Error Resume Next
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Function CodepaneSelection() As String
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    If EndLine - StartLine = 0 Then
        CodepaneSelection = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(StartLine, 1), StartColumn, EndColumn - StartColumn)
        Exit Function
    End If
    Dim str As String
    Dim i As Long
    For i = StartLine To EndLine
        If str = "" Then
            str = Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), StartColumn)
        ElseIf i < EndLine Then
            str = str & vbNewLine & Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), 1)
        Else
            str = str & vbNewLine & Mid(Application.VBE.ActiveCodePane.CodeModule.Lines(i, 1), 1, EndColumn - 1)
        End If
    Next
    CodepaneSelection = str
End Function

Public Function PartAfterCodePaneSelection(StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    Dim str As String
    str = Application.VBE.ActiveCodePane.CodeModule.Lines(EndLine, 1)
    str = Mid(str, EndColumn)
    PartAfterCodePaneSelection = str
End Function

Public Function PartBeforeCodePaneSelection(StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    Dim str As String
    str = Application.VBE.ActiveCodePane.CodeModule.Lines(StartLine, 1)
    str = Mid(str, 1, StartColumn - 1)
    PartBeforeCodePaneSelection = str
End Function

Sub ListProceduresInModule(Optional vbComp As VBComponent)
    '#INCLUDE ProcListArray
    '#INCLUDE ActiveModule
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    Dim procedures As Variant
    procedures = ProcListArray(vbComp)
    Dim txt As String
    txt = "'" & Join(procedures, vbNewLine & "'") & vbNewLine
    vbComp.CodeModule.InsertLines 1, txt
End Sub

Sub ListProceduresInAllModulesOfWorkbook(Optional wb As Workbook)
    '#INCLUDE ListProceduresInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        ListProceduresInModule vbComp
    Next
End Sub

Sub MakeProceduresInModulePrivate(Optional vbComp As VBComponent)
    '#INCLUDE ActiveModule
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    Dim n As Long
    Dim s As String
    Dim lineString As String
    With vbComp.CodeModule
        If .CountOfLines = 0 Then Exit Sub
        For n = .CountOfLines To 1 Step -1
            s = .Lines(n, 1)
            If Trim(s) Like "Public Sub *" Then
                .ReplaceLine n, Strings.Replace(s, "Public Sub ", "Private Sub ", , 1)
            ElseIf Trim(s) Like "Public Function *" Then
                .ReplaceLine n, Strings.Replace(s, "Public Function ", "Private Function ", , 1)
            ElseIf Trim(s) Like "Function *" Then
                .ReplaceLine n, Strings.Replace(s, "Function ", "Private Function ", , 1)
            ElseIf Trim(s) Like "Sub *" Then
                .ReplaceLine n, Strings.Replace(s, "Sub ", "Private Sub ", , 1)
            ElseIf Trim(s) Like "Public Const *" Then
                .ReplaceLine n, Strings.Replace(s, "Public Const ", "Private Const ", , 1)
            ElseIf Trim(s) Like "Public *" Then
                .ReplaceLine n, Strings.Replace(s, "Public ", "Private ", , 1)
            End If
        Next n
    End With
End Sub

Sub MakeProceduresInWorkbookPrivate(Optional wb As Workbook)
    '#INCLUDE MakeProceduresInModulePrivate
    '#INCLUDE ActiveCodepaneWorkbook
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        MakeProceduresInModulePrivate vbComp
    Next
End Sub

Sub MakeProceduresInModulePublic(Optional vbComp As VBComponent)
    '#INCLUDE ActiveModule
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    Dim n As Long
    Dim s As String
    Dim lineString As String
    With vbComp.CodeModule
        If .CountOfLines = 0 Then Exit Sub
        For n = .CountOfLines To 1 Step -1
            s = .Lines(n, 1)
            If Trim(s) Like "Private Sub *" Then
                .ReplaceLine n, Strings.Replace(s, "Private Sub ", "Public Sub ", , 1)
            ElseIf Trim(s) Like "Private Function *" Then
                .ReplaceLine n, Strings.Replace(s, "Private Function ", "Public Function ", , 1)
            ElseIf Trim(s) Like "Function *" Then
                .ReplaceLine n, Strings.Replace(s, "Function ", "Public Function ", , 1)
            ElseIf Trim(s) Like "Sub *" Then
                .ReplaceLine n, Strings.Replace(s, "Sub ", "Public Sub ", , 1)
            ElseIf Trim(s) Like "Private *" Then
                .ReplaceLine n, Strings.Replace(s, "Private ", "Public ", , 1)
            ElseIf Trim(s) Like "Declare *" Then
                .ReplaceLine n, Strings.Replace(s, "Declare ", "Public Declare ", , 1)
            ElseIf Trim(s) Like "Private Declare *" Then
                .ReplaceLine n, Strings.Replace(s, "Private Declare", "Public Declare ", , 1)
            ElseIf Trim(s) Like "Const *" Then
                .ReplaceLine n, Strings.Replace(s, "Const ", "Public Const ", , 1)
            ElseIf Trim(s) Like "Private Const *" Then
                .ReplaceLine n, Strings.Replace(s, "Private Const ", "Public Const ", , 1)
            End If
        Next n
    End With
End Sub

Sub MakeProceduresInWorkbookPublic(Optional TargetWorkbook As Workbook)
    '#INCLUDE MakeProceduresInModulePublic
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In TargetWorkbook.VBProject.VBComponents
        MakeProceduresInModulePublic vbComp
    Next
End Sub

Public Function ModReplaceMulti( _
       ByVal compare As VbCompareMethod, _
       ByVal str As String, _
       toStr As String, _
       ParamArray replacements() As Variant) _
        As String
    Rem ModReplaceMulti vbTextCompare, "a b c d", "X",array("a","c")
    Rem returns: "X b X d"
    '#INCLUDE compare
    Dim element As Variant
    For Each element In replacements
        str = Replace(str, element, toStr, , , compare)
    Next
    ModReplaceMulti = str
End Function

Public Sub AddLineNumbersToModule(Optional vbComp As VBComponent)
    '#INCLUDE AddLineNumbersToProcedure
    '#INCLUDE ProceduresOfModule
    '#INCLUDE ActiveModule
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    Dim element
    For Each element In ProceduresOfModule(vbComp)
        AddLineNumbersToProcedure vbComp, CStr(element)
    Next
End Sub

Public Sub RemoveLineNumbersFromModule(Optional vbComp As VBComponent)
    '#INCLUDE RemoveLineNumbersFromProcedure
    '#INCLUDE ProceduresOfModule
    '#INCLUDE ActiveModule
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    Dim element
    For Each element In ProceduresOfModule(vbComp)
        RemoveLineNumbersFromProcedure vbComp, CStr(element)
    Next
End Sub

Public Sub AddLineNumbersToProcedure( _
       Optional vbComp As VBComponent, _
       Optional ProcedureName As String)
    Rem number & ":" or number & vbtab
    '#INCLUDE IsCodepaneLineNumberAble
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    Dim txt
    Dim varr
    Dim i As Long
    Dim a As Long
    a = 1
    varr = Split(GetProcText(vbComp, ProcedureName), vbNewLine)
    For i = LBound(varr) To UBound(varr)
        If txt = "" Then
            If IsCodepaneLineNumberAble(varr(i)) Then
                txt = a & ":" & varr(i)
                a = a + 1
            Else
                txt = varr(i)
            End If
        Else
            If IsCodepaneLineNumberAble(varr(i)) And Right(Trim(varr(i - 1)), 1) <> "_" Then
                txt = txt & vbNewLine & a & ":" & varr(i)
                a = a + 1
            Else
                txt = txt & vbNewLine & varr(i)
            End If
        End If
    Next i
    StartLine = vbComp.CodeModule.ProcStartLine(ProcedureName, vbext_pk_Proc)
    NumLines = vbComp.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Proc)
    vbComp.CodeModule.DeleteLines StartLine:=StartLine, count:=NumLines
    vbComp.CodeModule.InsertLines StartLine, txt
End Sub

Public Sub RemoveLineNumbersFromProcedure(Optional Module As VBComponent, Optional ProcedureName As String)
    '#INCLUDE FirstDigit
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    If Module Is Nothing Then Set Module = ActiveModule
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    Dim StartLine As Long
    Dim NumLines As Long
    Dim txt
    Dim varr
    varr = Split(GetProcText(Module, ProcedureName), vbNewLine)
    Dim i As Long
    For i = LBound(varr) To UBound(varr)
        If txt = "" Then
            If Not IsNumeric(left(Trim(varr(i)), 1)) Then
                txt = varr(i)
            Else
                txt = left(varr(i), FirstDigit(varr(i)) - 1) & Right(varr(i), Len(varr(i)) - InStr(1, varr(i), ":") - 1)
            End If
        Else
            If Not IsNumeric(left(Trim(varr(i)), 1)) Then
                txt = txt & vbNewLine & varr(i)
            Else
                varr(i) = varr(i) & " "
                txt = txt & vbNewLine & left(varr(i), FirstDigit(varr(i)) - 1) & Right(varr(i), Len(varr(i)) - InStr(1, varr(i), ":") - 1)
            End If
        End If
    Next i
    StartLine = Module.CodeModule.ProcStartLine(ProcedureName, vbext_pk_Proc)
    NumLines = Module.CodeModule.ProcCountLines(ProcedureName, vbext_pk_Proc)
    Module.CodeModule.DeleteLines StartLine:=StartLine, count:=NumLines
    Module.CodeModule.InsertLines StartLine, txt
End Sub

Public Function IsCodepaneLineNumberAble(ByVal str As String) As Boolean
    Dim test As String
    test = Trim(str)
    If Len(test) = 0 Then Exit Function
    If Right(test, 1) = ":" Then Exit Function
    If IsNumeric(left(test, 1)) Then Exit Function
    If test Like "'*" Then Exit Function
    If test Like "Rem*" Then Exit Function
    If test Like "Dim*" Then Exit Function
    If test Like "Sub*" Then Exit Function
    If test Like "Public*" Then Exit Function
    If test Like "Private*" Then Exit Function
    If test Like "Function*" Then Exit Function
    If test Like "End Sub*" Then Exit Function
    If test Like "End Function*" Then Exit Function
    If test Like "Debug*" Then Exit Function
    IsCodepaneLineNumberAble = True
End Function

Public Function ProcedureInfo(Module As VBComponent, ProcedureName As String, ProcKind As VBIDE.vbext_ProcKind) As ProcInfo
    '#INCLUDE GetProcedureDeclaration
    Dim PInfo As ProcInfo
    Dim BodyLine As Long
    Dim Declaration As String
    Dim FirstLine As String
    BodyLine = Module.CodeModule.ProcStartLine(ProcedureName, ProcKind)
    If BodyLine > 0 Then
        PInfo.procName = ProcedureName
        PInfo.ProcKind = ProcKind
        PInfo.ProcBodyLine = Module.CodeModule.ProcBodyLine(ProcedureName, ProcKind)
        PInfo.ProcCountLines = Module.CodeModule.ProcCountLines(ProcedureName, ProcKind)
        PInfo.ProcStartLine = Module.CodeModule.ProcStartLine(ProcedureName, ProcKind)
        FirstLine = Module.CodeModule.Lines(PInfo.ProcBodyLine, 1)
        If StrComp(left(FirstLine, Len("Public")), "Public", vbBinaryCompare) = 0 Then
            PInfo.ProcedureScope = Public_SCOPE
        ElseIf StrComp(left(FirstLine, Len("Private")), "Private", vbBinaryCompare) = 0 Then
            PInfo.ProcedureScope = PRIVATE_SCOPE
        ElseIf StrComp(left(FirstLine, Len("Friend")), "Friend", vbBinaryCompare) = 0 Then
            PInfo.ProcedureScope = FRIEND_SCOPE
        Else
            PInfo.ProcedureScope = DEFAULT_SCOPE
        End If
        PInfo.ProcDeclaration = GetProcedureDeclaration(Module, ProcedureName, LineSplitKeep)
    End If
    ProcedureInfo = PInfo
End Function

Function ProceduresOfModule(Module As VBComponent) As Collection
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim coll As New Collection
    Dim procName As String
    With Module.CodeModule
        LineNum = .CountOfDeclarationLines + 1
        Do Until LineNum >= .CountOfLines
            procName = .ProcOfLine(LineNum, ProcKind)
            coll.Add procName
            LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
        Loop
    End With
    Set ProceduresOfModule = coll
End Function

Public Sub SeparateProceduresAndFunctionsInModule(Optional vbComp As VBComponent)
    '#INCLUDE GetProcedureDeclaration
    '#INCLUDE ProcListArray
    '#INCLUDE ActiveModule
    '#INCLUDE GetProcText
    '#INCLUDE SortArray
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    If vbComp.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim StartLine As Long
    StartLine = vbComp.CodeModule.CountOfDeclarationLines
    Dim totalLines As Long
    totalLines = vbComp.CodeModule.CountOfLines - vbComp.CodeModule.CountOfDeclarationLines
    Dim TheSubs As String, TheFunctions As String, TheOther As String
    Dim sProcedureDeclaration As String
    Dim sProcedureText As String
    Dim sProcedureName As String
    Dim i As Long
    Dim varr
    varr = ProcListArray(vbComp)
    SortArray varr
    For i = LBound(varr) To UBound(varr)
        sProcedureName = CStr(varr(i))
        sProcedureDeclaration = GetProcedureDeclaration(vbComp, sProcedureName, 0)
        sProcedureText = GetProcText(vbComp, sProcedureName)
        If InStr(1, sProcedureDeclaration, "Sub " & sProcedureName) > 0 Then
            TheSubs = IIf(TheSubs = "", sProcedureText, TheSubs & vbNewLine & sProcedureText)
        ElseIf InStr(1, sProcedureDeclaration, "Function " & sProcedureName) > 0 Then
            TheFunctions = IIf(TheFunctions = "", sProcedureText, TheFunctions & vbNewLine & sProcedureText)
        End If
    Next i
    vbComp.CodeModule.DeleteLines IIf(StartLine <> 0, StartLine, StartLine + 1), totalLines
    vbComp.CodeModule.AddFromString TheFunctions & vbNewLine
    vbComp.CodeModule.AddFromString TheSubs & vbNewLine
End Sub

Public Sub SeparateProceduresAndFunctionsInWorkbook(Optional wb As Workbook)
    '#INCLUDE SeparateProceduresAndFunctionsInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then SeparateProceduresAndFunctionsInModule vbComp
    Next
End Sub

Public Sub SeparatePublicAndPrivateInWorkbook(Optional wb As Workbook)
    '#INCLUDE SeparatePublicAndPrivateInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If wb Is Nothing Then Set wb = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then SeparatePublicAndPrivateInModule vbComp
    Next
End Sub

Public Sub SeparatePublicAndPrivateInModule(Optional vbComp As VBComponent)
    '#INCLUDE GetProcedureDeclaration
    '#INCLUDE ProcListArray
    '#INCLUDE ActiveModule
    '#INCLUDE GetProcText
    '#INCLUDE SortArray
    If vbComp Is Nothing Then Set vbComp = ActiveModule
    If vbComp.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim StartLine As Long
    StartLine = vbComp.CodeModule.CountOfDeclarationLines
    Dim totalLines As Long
    totalLines = vbComp.CodeModule.CountOfLines - vbComp.CodeModule.CountOfDeclarationLines
    Dim ThePublic As String, ThePrivate As String, TheOther As String
    Dim sProcedureDeclaration As String
    Dim sProcedureText As String
    Dim sProcedureName As String
    Dim i As Long
    Dim varr
    varr = ProcListArray(vbComp)
    SortArray varr
    For i = LBound(varr) To UBound(varr)
        sProcedureName = CStr(varr(i))
        sProcedureDeclaration = GetProcedureDeclaration(vbComp, sProcedureName, 0)
        sProcedureText = GetProcText(vbComp, sProcedureName)
        If InStr(1, sProcedureDeclaration, "Public ") > 0 Then
            ThePublic = IIf(ThePublic = "", sProcedureText, ThePublic & vbNewLine & sProcedureText)
        Else
            ThePrivate = IIf(ThePrivate = "", sProcedureText, ThePrivate & vbNewLine & sProcedureText)
        End If
    Next i
    vbComp.CodeModule.DeleteLines IIf(StartLine <> 0, StartLine, StartLine + 1), totalLines
    vbComp.CodeModule.AddFromString ThePrivate & vbNewLine
    vbComp.CodeModule.AddFromString ThePublic & vbNewLine
End Sub

Public Sub SetToNothing(Optional ProcedureName As String, Optional TargetWorkbook As Workbook)
    '#INCLUDE ProcedureFirstLine
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE ModuleOfProcedure
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    Dim Module As VBComponent
    If TargetWorkbook Is Nothing Then
        Set Module = ActiveModule
    Else
        Set Module = ModuleOfProcedure(TargetWorkbook, ProcedureName)
    End If
    Dim FirstLine As Long, LastLine As Long, LineNumber As Long
    Dim strLine As String, Append As String, Terminate As String
    FirstLine = ProcedureFirstLine(Module, ProcedureName)
    LastLine = ProcedureEndLine(Module, ProcedureName)
    For LineNumber = FirstLine To LastLine
        strLine = Trim(Module.CodeModule.Lines(LineNumber, 1))
        If strLine Like "Set * = *" Or strLine Like "Dim*As New*" Then
            Terminate = Split(strLine, " ")(1)
            Append = Append & vbNewLine & "Set " & Terminate & " = Nothing"
        End If
    Next
    If Append <> "" Then Module.CodeModule.InsertLines LastLine, Append
End Sub

Public Sub ShowProcedureInfo(Module As VBComponent, procName As String)
    '#INCLUDE ProcedureInfo
    Dim ProcKind As VBIDE.vbext_ProcKind: ProcKind = vbext_pk_Proc
    Dim PInfo As ProcInfo: PInfo = ProcedureInfo(Module, procName, ProcKind)
    Debug.Print "ProcName: " & PInfo.procName
    Debug.Print "ProcKind: " & CStr(PInfo.ProcKind)
    Debug.Print "ProcStartLine: " & CStr(PInfo.ProcStartLine)
    Debug.Print "ProcBodyLine: " & CStr(PInfo.ProcBodyLine)
    Debug.Print "ProcCountLines: " & CStr(PInfo.ProcCountLines)
    Debug.Print "ProcedureScope: " & CStr(PInfo.ProcedureScope)
    Debug.Print "ProcDeclaration: " & PInfo.ProcDeclaration
End Sub

Public Function SingleSpace(ByVal TEXT As String) As String
    Dim pos As String
    pos = InStr(1, TEXT, Space(2), vbBinaryCompare)
    Do Until pos = 0
        TEXT = Replace(TEXT, Space(2), Space(1))
        pos = InStr(1, TEXT, Space(2), vbBinaryCompare)
    Loop
    SingleSpace = TEXT
End Function

Public Sub SortArrayBetween(vArray As Variant, inLow As Long, inHi As Long)
    Dim tmpSwap As Variant
    Dim tmpLow  As Long:    tmpLow = inLow
    Dim tmpHi   As Long:      tmpHi = inHi
    Dim pivot   As Variant:    pivot = vArray((inLow + inHi) \ 2)
    While (tmpLow <= tmpHi)
        While (vArray(tmpLow) < pivot And tmpLow < inHi)
            tmpLow = tmpLow + 1
        Wend
        While (pivot < vArray(tmpHi) And tmpHi > inLow)
            tmpHi = tmpHi - 1
        Wend
        If (tmpLow <= tmpHi) Then
            tmpSwap = vArray(tmpLow)
            vArray(tmpLow) = vArray(tmpHi)
            vArray(tmpHi) = tmpSwap
            tmpLow = tmpLow + 1
            tmpHi = tmpHi - 1
        End If
    Wend
    If (inLow < tmpHi) Then SortArrayBetween vArray, inLow, tmpHi
    If (tmpLow < inHi) Then SortArrayBetween vArray, tmpLow, inHi
End Sub

Public Sub SortProceduresInModule(Optional Module As VBComponent)
    '#INCLUDE ProcListArray
    '#INCLUDE ActiveModule
    '#INCLUDE GetProcText
    '#INCLUDE SortArray
    If Module Is Nothing Then Set Module = ActiveModule
    If Module.CodeModule.CountOfLines = 0 Then Exit Sub
    Dim varr: varr = ProcListArray(Module)
    Dim StartLine As Long: StartLine = Module.CodeModule.ProcStartLine(varr(0), vbext_pk_Proc)
    Dim totalLines As Long: totalLines = Module.CodeModule.CountOfLines - Module.CodeModule.CountOfDeclarationLines
    varr = SortArray(varr)
    Dim ReplacedProcedures As String
    Dim i As Long
    For i = LBound(varr) To UBound(varr)
        If ReplacedProcedures = "" Then
            ReplacedProcedures = GetProcText(Module, CStr(varr(i)))
        Else
            ReplacedProcedures = ReplacedProcedures & vbNewLine & _
                                 GetProcText(Module, CStr(varr(i)))
        End If
    Next i
    Module.CodeModule.DeleteLines StartLine, totalLines
    Module.CodeModule.AddFromString ReplacedProcedures
End Sub

Public Sub SortProceduresWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE SortProceduresInModule
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        SortProceduresInModule Module
    Next
End Sub

Sub SortOnComma()
    '#INCLUDE SortSelection
    SortSelection ","
End Sub

Sub SortOnNewline()
    '#INCLUDE SortSelection
    SortSelection vbNewLine
End Sub

Public Sub SortSelection(delimeter As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    '#INCLUDE SortSelectionArray
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Dim arr
    arr = Split(Code, delimeter)
    SortSelectionArray arr
    Code = Join(arr, delimeter)
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Code & _
        PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Sub SortSelectionArray(ByRef tempArray As Variant)
    Dim MaxVal As Variant
    Dim MaxIndex As Integer
    Dim i As Integer, j As Integer
    For i = UBound(tempArray) To 0 Step -1
        MaxVal = tempArray(i)
        MaxIndex = i
        For j = 0 To i
            If tempArray(j) > MaxVal Then
                MaxVal = tempArray(j)
                MaxIndex = j
            End If
        Next j
        If MaxIndex < i Then
            tempArray(MaxIndex) = tempArray(i)
            tempArray(i) = MaxVal
        End If
    Next i
End Sub

Public Sub CodePaneSelectionSubstitute(OldValue As String, NewValue As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE PartAfterCodePaneSelection
    '#INCLUDE PartBeforeCodePaneSelection
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    Dim Code As String
    Code = CodepaneSelection
    Code = PartBeforeCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn) _
      & Replace(Code, OldValue, NewValue, , , vbTextCompare) & _
                                                             PartAfterCodePaneSelection(StartLine, StartColumn, EndLine, EndColumn)
    Application.VBE.ActiveCodePane.CodeModule.DeleteLines StartLine, EndLine - StartLine + 1
    Application.VBE.ActiveCodePane.CodeModule.InsertLines StartLine, Code
End Sub

Public Function CodePaneSelectionStartLine() As Long
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    CodePaneSelectionStartLine = StartLine
End Function

Public Function CodePaneSelectionStartColumn() As Long
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    CodePaneSelectionStartColumn = StartColumn
End Function

Public Function CodePaneSelectionEndLine() As Long
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    CodePaneSelectionEndLine = EndLine
End Function

Public Function CodePaneSelectionEndColumn() As Long
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    CodePaneSelectionEndColumn = EndColumn
End Function

Public Function CodepaneSelectionRowsCount() As Long
    Dim StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long
    Application.VBE.ActiveCodePane.GetSelection StartLine, StartColumn, EndLine, EndColumn
    CodepaneSelectionRowsCount = EndLine - StartLine + 1
End Function

Sub CodePaneSelectionSet(StartLine As Long, StartColumn As Long, EndLine As Long, EndColumn As Long)
    Application.VBE.ActiveCodePane.SetSelection StartLine, StartColumn, EndLine, EndColumn
End Sub

Rem @Folder FormatVBATools
Sub sysAddHeader()
    '#INCLUDE GetProcedureDeclaration
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE GetCurrentProcInfo
    '#INCLUDE AddStringParameterFromProcedureDeclaration
    '#INCLUDE TypeProcedureComment
    '#INCLUDE ActiveModule
    Dim txtName As String
    txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    Dim txtContacts As String
    txtContacts = AUTHOR_EMAIL
    If txtContacts <> vbNullString Then txtContacts = "'* Contacts   :" & vbTab & txtContacts & vbCrLf
    Dim txtCopyright As String
    txtCopyright = AUTHOR_COPYRIGHT
    If txtCopyright <> vbNullString Then txtCopyright = "'* Copyright  :" & vbTab & txtCopyright & vbCrLf
    Dim txtOther As String
    txtOther = AUTHOR_OTHERTEXT
    If txtOther <> vbNullString Then txtOther = "'* Other      :" & vbTab & txtOther & vbCrLf
    Dim txtMedia As String
    txtMedia = "'* " & vbLf & AUTHOR_MEDIA
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim CurentCodePane As CodePane
    Set CurentCodePane = Module.CodeModule.CodePane
    Dim nLine  As Long
    Dim i      As Byte
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim sProc  As String
    Dim sTemp  As String
    Dim sTime  As String
    Dim sType  As String
    Dim sProcDeclartion As String
    Dim sProcArguments As String
    On Error Resume Next
    With CurentCodePane
        GetCurrentProcInfo nLine, sProc, CurentCodePane
        sTemp = Replace(String(90, "*"), "**", "* ")
        sTime = Format(Now, ctFormat)
        If sProc = "" Or CodePaneSelectionStartLine = 1 Then
            sType = "* Module     :"
            sProc = .CodeModule.Name
            nLine = 1
        Else
            txtMedia = ""
            For i = 0 To 4
                ProcKind = i
                sProcDeclartion = GetProcedureDeclaration(Module, sProc, ProcKind)
                If sProcDeclartion <> vbNullString Then Exit For
            Next
            sProcArguments = AddStringParameterFromProcedureDeclaration(sProcDeclartion)
            sType = TypeProcedureComment(sProcDeclartion)
        End If
        sTemp = vbLf & "'" & sTemp & vbCrLf & _
                "'" & sType & vbTab & sProc & vbCrLf & _
                "'* Created    :" & vbTab & sTime & vbTab & vbCrLf & _
                "'* Author     :" & vbTab & txtName & vbCrLf & _
                txtContacts & _
                txtCopyright & _
                txtOther & _
                txtMedia & _
                sProcArguments & _
                "'" & sTemp
        .CodeModule.InsertLines nLine, sTemp & vbNewLine
    End With
End Sub

Sub sysAddModified()
    Rem Author VBATools
    '#INCLUDE GetCurrentProcInfo
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    Dim CurentCodePane As CodePane
    Set CurentCodePane = Module.CodeModule.CodePane
    Dim nLine  As Long
    Dim sProc  As String
    Dim sTime  As String
    Dim sSecondLine As String
    Dim sUser As String
    Const sUPDATE As String = "'* Updated    :"
    Const sFersLine As String = "'* Modified   :" & vbTab & "Date and Time" & vbTab2 & "Author" & vbTab4 & "Description" & vbCrLf
    On Error Resume Next
    With CurentCodePane
        GetCurrentProcInfo nLine, sProc, CurentCodePane
        sTime = Format(Now, ctFormat)
        sUser = "Alex"
        If sUser = vbNullString Then sUser = Environ("UserName")
        sSecondLine = sUPDATE & vbTab & sTime & vbTab & sUser & vbTab2
        If Not .CodeModule.Lines(nLine - 2, 1) Like sUPDATE & "*" Then
            sSecondLine = sFersLine & sSecondLine
        End If
        .CodeModule.InsertLines nLine - 1, sSecondLine
    End With
End Sub

Private Sub GetCurrentProcInfo(ByRef nLine As Long, ByRef sProc As String, ByRef CurentCodePane As CodePane)
    Dim t      As Long
    With CurentCodePane
        .GetSelection nLine, t, t, t
        sProc = .CodeModule.ProcOfLine(nLine, vbext_pk_Proc)
        If sProc = "" Then
            Do While .CodeModule.Find("'*", nLine, 1, .CodeModule.CountOfDeclarationLines, 2)
                nLine = nLine + 1
                If nLine > .CodeModule.CountOfDeclarationLines Then Exit Do
            Loop
        Else
            nLine = .CodeModule.ProcBodyLine(sProc, vbext_pk_Proc)
        End If
    End With
End Sub

Private Function AddStringParameterFromProcedureDeclaration(ByVal sPocDeclartion As String) As String
    Dim sDeclaration As String
    sDeclaration = Right$(sPocDeclartion, Len(sPocDeclartion) - InStr(1, sPocDeclartion, "("))
    sDeclaration = left$(sDeclaration, InStr(1, sDeclaration, ")") - 1)
    If sDeclaration = vbNullString Then Exit Function
    Dim arStr() As String
    Dim sTemp  As String
    Dim i      As Byte
    Dim iMaxLen As Byte
    Dim iTempLen As Byte
    arStr = Split(sDeclaration, ",")
    iMaxLen = 0
    For i = 0 To UBound(arStr)
        iTempLen = Len(Trim$(arStr(i)))
        If iMaxLen < iTempLen Then iMaxLen = iTempLen
    Next i
    sDeclaration = "'*" & vbLf & "'* Argument(s):" & String$(iMaxLen - Len(Trim$("'* Argument(s):")), " ") & vbTab2 & "Description" & vbCrLf & "'*" & vbCrLf
    For i = 0 To UBound(arStr)
        sTemp = "'* " & Trim$(arStr(i)) & String$(iMaxLen - Len(Trim$(arStr(i))), " ") & " :"
        sDeclaration = sDeclaration & sTemp & vbCrLf
    Next i
    AddStringParameterFromProcedureDeclaration = sDeclaration & "'* " & vbCrLf
End Function

Private Function TypeProcedureComment(ByRef StrDeclarationProcedure As String) As String
    If StrDeclarationProcedure Like "*Sub*" Then
        TypeProcedureComment = "* Sub        :"
    ElseIf StrDeclarationProcedure Like "*Function*" Then
        TypeProcedureComment = "* Function   :"
    ElseIf StrDeclarationProcedure Like "*Property Set*" Then
        TypeProcedureComment = "* Prop Set   :"
    ElseIf StrDeclarationProcedure Like "*Property Get*" Then
        TypeProcedureComment = "* Prop Get   :"
    ElseIf StrDeclarationProcedure Like "*Property Let*" Then
        TypeProcedureComment = "* Prop Let   :"
    Else
        TypeProcedureComment = "* Un Type    :"
    End If
End Function


