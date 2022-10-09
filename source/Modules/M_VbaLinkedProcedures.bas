Attribute VB_Name = "M_VbaLinkedProcedures"
Rem @Folder STACK
Function getWhichProceduresCallTargetProcedure(TargetWorkbook As Workbook, Optional ProcedureName As String)
    Rem @TODO
    Rem This may be faster than the previous method i used in Stack
    '#INCLUDE ProceduresOfModule
    '#INCLUDE InStrExact
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    Dim Module As VBIDE.VBComponent
    Dim Procedure As Variant
    Dim i As Long
    Dim output As String
    Dim matchCollection As New Collection
    Dim ProcedureText As String
    For Each Module In TargetWorkbook.VBProject.VBComponents
        For Each Procedure In ProceduresOfModule(Module)
            ProcedureText = GetProcText(Module, Procedure)
            If InStr(1, ProcedureText, ProcedureName, vbTextCompare) Then
                If InStrExact(1, ProcedureText, ProcedureName, False) > 0 Then
                    output = IIf(output = "", Procedure, output & vbNewLine & Procedure)
                End If
            End If
        Next
    Next
    getWhichProceduresCallTargetProcedure = output
End Function

Sub ListAllProcedureImportsInWorkbook(Optional TargetWorkbook As Workbook)
    Rem CommentsRemoveWorkbook thisworkbook
    '#INCLUDE AddListOfLinkedProceduresToProcedure
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim procedures As Collection: Set procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure
    For Each Procedure In procedures
        AddListOfLinkedProceduresToProcedure CStr(Procedure), procedures, TargetWorkbook
    Next
End Sub

Sub ListAllProcedureImportsInModule(Optional Module As VBComponent)
    '#INCLUDE ProceduresOfModule
    '#INCLUDE AddListOfLinkedProceduresToProcedure
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveModule
    '#INCLUDE WorkbookOfModule
    If Module Is Nothing Then Set Module = ActiveModule
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = WorkbookOfModule(Module)
    Dim procedures As Collection: Set procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure
    For Each Procedure In ProceduresOfModule(Module)
        AddListOfLinkedProceduresToProcedure CStr(Procedure), procedures, TargetWorkbook
    Next
End Sub

Sub ExportAllProcedures(Optional TargetWorkbook As Workbook)
    Rem CommentsRemoveWorkbook thisworkbook
    '#INCLUDE ExportProcedure
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim procedures As Collection: Set procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure
    For Each Procedure In procedures
        ExportProcedure CStr(Procedure), TargetWorkbook
    Next
End Sub

Sub ExportProcedure( _
    Optional ProcedureName As String, _
    Optional FromWorkbook As Workbook)
    Rem @star
    '#INCLUDE CodepaneSelection
    '#INCLUDE AddListOfLinkedProceduresToProcedure
    '#INCLUDE ExportTargetProcedure
    '#INCLUDE LinkedProcs
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ProcedureExists
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE ListContainedProceduresInTXT
    '#INCLUDE FollowLink
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    '#INCLUDE CollectionContains
    If ProcedureName = "" Then
        If Len(CodepaneSelection) = 0 Then
            ProcedureName = ActiveProcedure
        Else
            ProcedureName = CodepaneSelection
        End If
    End If
    If FromWorkbook Is Nothing Then Set FromWorkbook = ActiveCodepaneWorkbook
    If ProcedureExists(ProcedureName, FromWorkbook) = False Then
        MsgBox ProcedureName & " not found in workbook " & FromWorkbook.Name
        Exit Sub
    End If
    AddListOfLinkedProceduresToProcedure CStr(ProcedureName), ProceduresOfWorkbook(FromWorkbook), FromWorkbook
    Dim ExportedProcedures As New Collection
    Dim Proccessed As New Collection
    On Error Resume Next
    ExportTargetProcedure ProcedureName
    ExportedProcedures.Add CStr(ProcedureName), CStr(ProcedureName)
    For Each Procedure In LinkedProcs(ProcedureName, FromWorkbook)
        AddListOfLinkedProceduresToProcedure CStr(Procedure), ProceduresOfWorkbook(FromWorkbook), FromWorkbook
        ExportTargetProcedure CStr(Procedure), FromWorkbook
        ExportedProcedures.Add CStr(Procedure), CStr(Procedure)
    Next
    Dim ProceduresCount As Long
    ProceduresCount = ExportedProcedures.count
retry:
    For Each Procedure In ExportedProcedures
        For Each element In LinkedProcs(Procedure, FromWorkbook)
            If Not CollectionContains(ExportedProcedures, , element) Then
                AddListOfLinkedProceduresToProcedure CStr(element), ProceduresOfWorkbook(FromWorkbook), FromWorkbook
                ExportedProcedures.Add CStr(element), CStr(element)
            End If
        Next
    Next
    If ExportedProcedures.count > ProceduresCount Then
        ProceduresCount = ExportedProcedures.count
        GoTo retry
    End If
    On Error GoTo 0
    If ExportedProcedures.count > 1 Then
        Dim MergedName As String
        Dim procFile As String
        MergedName = "Merged_" & ProcedureName
        Dim MergedString As String
        For Each Procedure In ExportedProcedures
            procFile = SNIP_FOLDER & CStr(Procedure) & ".txt"
            MergedString = IIf(MergedString = "", TxtRead(procFile), MergedString & vbNewLine & TxtRead(procFile))
        Next
        TxtOverwrite SNIP_FOLDER & MergedName & ".txt", MergedString
        ListContainedProceduresInTXT SNIP_FOLDER & MergedName & ".txt"
    End If
    FollowLink SNIP_FOLDER
End Sub

Sub AddListOfLinkedProceduresToProcedure(Optional ProcedureName As String, Optional procedures As Collection, Optional FromWorkbook As Workbook)
    '#INCLUDE RegexTest
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ProcedureFirstLine
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If FromWorkbook Is Nothing Then Set FromWorkbook = ActiveCodepaneWorkbook
    If procedures Is Nothing Then Set procedures = ProceduresOfWorkbook(FromWorkbook)
    Dim ListOfImports As String
    Dim Module As VBComponent:  Set Module = ModuleOfProcedure(FromWorkbook, ProcedureName)
    Dim ProcedureText As String:    ProcedureText = GetProcText(Module, ProcedureName)
    Dim Procedure As Variant
    For Each Procedure In procedures
        If UCase(CStr(Procedure)) <> UCase(CStr(ProcedureName)) Then
            Rem         If InStr(1, ProcedureText, CStr(PROCEDURE)) > 0 Then
            If RegexTest(ProcedureText, "\W" & Procedure & "[.(\W]", , True) = True Then
                If InStr(1, ProcedureText, "#INCLUDE " & Procedure) = 0 And InStr(1, ListOfImports, "#INCLUDE " & Procedure) = 0 Then
                    If ListOfImports = "" Then
                        ListOfImports = "'#INCLUDE " & Procedure
                    Else
                        ListOfImports = ListOfImports & vbNewLine & "'#INCLUDE " & Procedure
                    End If
                End If
            End If
        End If
    Next
    If ListOfImports <> "" Then Module.CodeModule.InsertLines ProcedureFirstLine(Module, ProcedureName), ListOfImports
End Sub

Public Function RegexTest( _
       ByVal string1 As String, _
       ByVal stringPattern As String, _
       Optional ByVal globalFlag As Boolean, _
       Optional ByVal ignoreCaseFlag As Boolean, _
       Optional ByVal multilineFlag As Boolean) _
        As Boolean
    Dim REGEX As Object
    Set REGEX = CreateObject("VBScript.RegExp")
    With REGEX
        .Global = globalFlag
        .IgnoreCase = ignoreCaseFlag
        .MultiLine = multilineFlag
        .Pattern = stringPattern
    End With
    RegexTest = REGEX.test(string1)
End Function

Sub ImportProcedure( _
    Optional Procedure As String, _
    Optional TargetWorkbook As Workbook, _
    Optional Module As VBComponent, _
    Optional Overwrite As Boolean)
    Rem @todo file picker?
    '#INCLUDE CodepaneSelection
    '#INCLUDE ImportImports
    '#INCLUDE UpdateProcedureCode
    '#INCLUDE ProcedureExists
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE CreateOrSetModule
    '#INCLUDE CheckPath
    '#INCLUDE TXTReadFromUrl
    '#INCLUDE TxtOverwrite
    '#INCLUDE TxtRead
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    If Procedure = "" Then Procedure = CodepaneSelection
    If Procedure = "" Or InStr(1, Procedure, " ") > 0 Then Exit Sub
    Dim ProcedurePath As String
    ProcedurePath = SNIP_FOLDER & Procedure & ".txt"
    If CheckPath(ProcedurePath) = "I" Then
        On Error Resume Next
        Dim DownloadedProcedure As String
        DownloadedProcedure = TXTReadFromUrl("https://github.com/alexofrhodes/vbArc-Snippets/Procedures/raw/main/" & Procedure & ".txt")
        On Error GoTo 0
        If Len(DownloadedProcedure) > 0 Then
            TxtOverwrite SNIP_FOLDER & Procedure & ".txt", DownloadedProcedure
        Else
            MsgBox "File not found neither localy nor online"
            Exit Sub
        End If
    End If
    If ProcedureExists(Procedure, TargetWorkbook) = True Then
        If Overwrite = True Then UpdateProcedureCode Procedure, TxtRead(ProcedurePath), TargetWorkbook
    Else
        If Module Is Nothing Then Set Module = CreateOrSetModule("vbArcImports", vbext_ct_StdModule, TargetWorkbook)
        Module.CodeModule.AddFromFile ProcedurePath
    End If
    ImportImports ProcedurePath, TargetWorkbook, Module, Overwrite
End Sub

Sub ImportImports( _
    Optional Procedure As String, _
    Optional TargetWorkbook As Workbook, _
    Optional Module As VBComponent, _
    Optional Overwrite As Boolean)
    '#INCLUDE ImportProcedure
    '#INCLUDE TxtRead
    Dim var
    Dim importfile As String
    var = Split(TxtRead(Procedure), vbLf)
    Dim TextLine As Variant
    For Each TextLine In var
        TextLine = Trim(TextLine)
        If left(TextLine, 9) = "'#INCLUDE" Then
            importfile = Split(TextLine, " ")(1)
            ImportProcedure importfile, TargetWorkbook, Module, Overwrite
        End If
    Next
End Sub

Sub UpdateAllProcedures(Optional TargetWorkbook As Workbook)
    '#INCLUDE ProceduresOfModule
    '#INCLUDE ImportProcedure
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim procedures As New Collection
    Dim Procedure
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.Type = vbext_ct_StdModule Then
            Set procedures = ProceduresOfModule(Module)
            For Each Procedure In procedures
                If UCase(CStr(Procedure)) <> UCase("UpdateAllProcedures") Then ImportProcedure CStr(Procedure), TargetWorkbook, , True
            Next
        End If
    Next
End Sub

Public Sub UpdateProcedureCode( _
       Procedure As Variant, _
       Code As String, _
       TargetWorkbook As Workbook)
    '#INCLUDE ModuleOfProcedure
    Dim StartLine As Integer
    Dim NumLines As Integer
    Dim Module As VBComponent
    Set Module = ModuleOfProcedure(TargetWorkbook, Procedure)
    With Module.CodeModule
        StartLine = .ProcStartLine(Procedure, vbext_pk_Proc)
        NumLines = .ProcCountLines(Procedure, vbext_pk_Proc)
        .DeleteLines StartLine, NumLines
        .InsertLines StartLine, Code
    End With
End Sub

Function getAllMissingDependencies(Optional TargetWorkbook As Workbook) As Boolean
    '#INCLUDE getMissingMissingDependenciesOfProcedure
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim procedures As New Collection
    Set procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Procedure As Variant
    For Each Procedure In procedures
        getMissingMissingDependenciesOfProcedure CStr(Procedure), TargetWorkbook
    Next
End Function

Function getMissingMissingDependenciesOfProcedure( _
         Optional Procedure As String, _
         Optional TargetWorkbook As Workbook)
    '#INCLUDE ImportProcedure
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE CollectionContains
    If Procedure = "" Then Procedure = ActiveProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim procedures As New Collection
    Set procedures = ProceduresOfWorkbook(TargetWorkbook)
    Dim Code As String
    Code = GetProcText(ModuleOfProcedure(TargetWorkbook, Procedure), Procedure)
    Dim CodeLines As Variant
    CodeLines = Split(Code, vbNewLine)
    Dim CodeLine As Variant
    Dim RequiredProcedure As String
    Dim Log As String
    If InStr(1, Code, "'#INCLUDE", vbTextCompare) > 0 Then
        For Each CodeLine In CodeLines
            If left(CodeLine, 9) = "'#INCLUDE" Then
                RequiredProcedure = Split(CodeLine, " ")(1)
                If Not CollectionContains(procedures, , RequiredProcedure) Then
                    ImportProcedure RequiredProcedure, TargetWorkbook, , True
                End If
            End If
        Next
    End If
    ProcedureDependenciesExist = (Log = "")
End Function

Sub ArrayTrim(arr As Variant)
    Dim i As Long
    For i = LBound(arr) To UBound(arr)
        If TypeName(arr(i)) = "String" Then arr(i) = Trim(arr(i))
    Next
End Sub

Sub ExportTargetProcedure(Optional ProcedureName As String, _
                          Optional FromWorkbook As Workbook)
    '#INCLUDE LinkedProcs
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE FileExists
    '#INCLUDE FileLastModified
    '#INCLUDE TxtOverwrite
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If FromWorkbook Is Nothing Then Set FromWorkbook = ActiveCodepaneWorkbook
    Dim procedures As Collection: Set procedures = LinkedProcs(ProcedureName, FromWorkbook)
    On Error Resume Next
    procedures.Add ProcedureName, ProcedureName
    On Error GoTo 0
    Dim Procedure As Variant
    Dim FileFullName As String
    Dim lastMod As Date
    Dim timeDif As Long
    For Each Procedure In procedures
        FileFullName = SNIP_FOLDER & Procedure & ".txt"
        If FileExists(FileFullName) = False Then
            Debug.Print IIf(FileExists(FileFullName) = False, "NEW ", "OVERWROTE ") & Procedure
            TxtOverwrite FileFullName, GetProcText(ModuleOfProcedure(FromWorkbook, CStr(Procedure)), CStr(Procedure))
        Else
            On Error Resume Next
            lastMod = FileLastModified(FileFullName)
            timeDif = DateDiff("s", lastMod, Now())
            On Error GoTo 0
            If timeDif > 60 Then
                Debug.Print IIf(FileExists(FileFullName) = False, "NEW ", "OVERWROTE ") & Procedure
                TxtOverwrite FileFullName, GetProcText(ModuleOfProcedure(FromWorkbook, CStr(Procedure)), CStr(Procedure))
            End If
        End If
    Next
End Sub

Function LinkedProcs( _
         ProcedureName As Variant, _
         FromWorkbook As Workbook) As Collection
    Rem dp LinkedProcs("FindIfGetRow",thisworkbook)
    '#INCLUDE GetCallsOfProcedure
    '#INCLUDE dp
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE CollectionContains
    '#INCLUDE FindIfGetRow
    Dim AllProcedures As Collection:       Set AllProcedures = ProceduresOfWorkbook(FromWorkbook)
    Dim Proccessed As Collection:          Set Proccessed = New Collection
    Dim CalledProcedures As Collection:    Set CalledProcedures = New Collection
    GetCallsOfProcedure ModuleOfProcedure(FromWorkbook, ProcedureName), ProcedureName, AllProcedures, CalledProcedures
    Dim CalledProceduresCount As Long:    CalledProceduresCount = CalledProcedures.count
    Dim Procedure As Variant
    Dim Module As VBComponent
REPEAT:
    For Each Procedure In CalledProcedures
        If Not CollectionContains(Proccessed, , CStr(Procedure)) Then
            Proccessed.Add Procedure, CStr(Procedure)
            Set Module = ModuleOfProcedure(FromWorkbook, ProcedureName)
            GetCallsOfProcedure Module, CStr(Procedure), AllProcedures, CalledProcedures
        End If
    Next
    If CalledProcedures.count > CalledProceduresCount Then
        CalledProceduresCount = CalledProcedures.count
        GoTo REPEAT
    End If
    Set LinkedProcs = CalledProcedures
End Function

Sub GetCallsOfProcedure( _
    Module As VBComponent, _
    ProcedureName As Variant, _
    AllProcedures As Collection, _
    ByRef OutputCollection As Collection)
    '#INCLUDE ArrayTrim
    '#INCLUDE InStrExact
    '#INCLUDE GetProcText
    Dim Code As String: Code = GetProcText(Module, ProcedureName)
    Dim CodeLines As Variant
    Dim Procedure As Variant
    Dim CodeLine As Variant
    For Each Procedure In AllProcedures
        If CStr(Procedure) <> ProcedureName Then
            If InStr(1, Code, CStr(Procedure)) > 0 Then
                CodeLines = Split(Code, vbNewLine)
                ArrayTrim CodeLines
                For Each CodeLine In CodeLines
                    If InStrExact(1, CStr(CodeLine), CStr(Procedure), True) > 0 Then
                        On Error Resume Next
                        OutputCollection.Add CStr(Procedure), CStr(Procedure)
                        On Error GoTo 0
                        Exit For
                    End If
                Next
            End If
        End If
    Next Procedure
End Sub

Function InStrExact(Start As Long, SourceText As String, WordToFind As String, _
                    Optional CaseSensitive As Boolean = False, _
                    Optional AllowAccentedCharacters As Boolean = False) As Long
    Dim X As Long, Str1 As String, Str2 As String, Pattern As String
    Const UpperAccentsOnly As String = "ÇÉÑ"
    Const UpperAndLowerAccents As String = "ÇÉÑçéñ"
    If CaseSensitive Then
        Str1 = SourceText
        Str2 = WordToFind
        Pattern = "[!A-Za-z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAndLowerAccents)
    Else
        Str1 = UCase(SourceText)
        Str2 = UCase(WordToFind)
        Pattern = "[!A-Z0-9]"
        If AllowAccentedCharacters Then Pattern = Replace(Pattern, "!", "!" & UpperAccentsOnly)
    End If
    For X = Start To Len(Str1) - Len(Str2) + 1
        If Mid(" " & Str1 & " ", X, Len(Str2) + 2) Like Pattern & Str2 & Pattern _
                                                   And Not Mid(Str1, X) Like Str2 & "'[" & Mid(Pattern, 3) & "*" Then
            InStrExact = X
            Exit Function
        End If
    Next
End Function

Sub IsDuplicateProceduresInWorkbook(Optional TargetWorkbook As Workbook)
    '#INCLUDE dp
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE DuplicatesInArray
    '#INCLUDE CollectionToArray
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim var
    var = Split(DuplicatesInArray(CollectionToArray(ProceduresOfWorkbook(TargetWorkbook))), ",")
    dp var
    MsgBox "Found " & UBound(var) & " duplicate procedures. Result output in immediate window."
End Sub


