Attribute VB_Name = "F_VBE"
Rem @Folder VBE
Function HasProject(wb As Workbook) As Boolean
    Dim WbProjComp As Object
    On Error Resume Next
    Set WbProjComp = wb.VBProject.VBComponents
    If Not WbProjComp Is Nothing Then HasProject = True
End Function

Public Function ActiveProjName() As String
    ActiveProjName = Mid(Application.VBE.ActiveVBProject.fileName, InStrRev(Application.VBE.ActiveVBProject.fileName, "\") + 1)
End Function

Function ConfirmInstallAddin() As Boolean
    Dim eai As Excel.AddIn
    Dim fso As Object
    Dim oXL As Object
    Dim response As Integer
    Dim thisAddInDate As Date
    Dim thisFileLen As Long
    Dim existingAddInName As String
    Dim existingAddinDate As Date
    Dim existingFileLen As Long
    Dim ai As AddIn
    Dim msg As String
    Dim toInstall As Integer
    Dim copiedWbName As String
    Dim desiredAddInName As String: desiredAddInName = PROJECT_NAME & ".xlam"
    Dim deleteOld As Boolean: deleteOld = True
    On Error GoTo ErrorHandler
    thisAddInDate = FileDateTime(ThisWorkbook.FullName)
    thisFileLen = FileLen(ThisWorkbook.FullName)
    existingAddInName = ""
    For Each ai In Application.AddIns
        If ai.title = PROJECT_NAME Then
            existingAddInName = ai.FullName
            Exit For
        End If
    Next ai
    If existingAddInName <> "" Then
        existingAddinDate = FileDateTime(existingAddInName)
        existingFileLen = FileLen(existingAddInName)
        If thisAddInDate > existingAddinDate And thisFileLen <> existingFileLen Then
            msg = "Do you want to update " & desiredAddInName & " ?"
        ElseIf thisAddInDate <= existingAddinDate And thisFileLen <> existingFileLen Then
            msg = "Do you want to update " & desiredAddInName & " ?" & vbNewLine & _
                  "The file you opened is not newer than the installed file."
        Else
            Exit Function
        End If
    Else
        msg = "Do you want to install " & desiredAddInName & " ?"
        existingAddInName = Application.UserLibraryPath & desiredAddInName
    End If
    toInstall = MsgBox(msg, vbYesNo)
    If toInstall = vbYes Then ConfirmInstallAddin = True
ErrorHandler:
    MsgBox "Error #" & _
           err.Number & _
           vbCrLf & _
           "Please, let the Author know.", vbInformation
End Function

Public Sub AddinCreate()
    '#INCLUDE ConfirmInstallAddin
    If Not ConfirmInstallAddin Then Exit Sub
    Dim AddFolder As String
    On Error GoTo InstallationAdd_Err
    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
    If Dir(AddFolder, vbDirectory) = vbNullString Then
        Call MsgBox("Unfortunately, the program cannot install the add-in on this computer.")
        Exit Sub
    End If
    Dim addinsPath As String
    addinsPath = AddFolder
    Dim partName As String
    partName = Right(ThisWorkbook.FullName, Len(ThisWorkbook.FullName) - InStrRev(ThisWorkbook.FullName, "\"))
    partName = left(partName, InStr(1, partName, ".") - 1)
    If Dir(addinsPath & partName & ".xlam") <> "" Then AddIns(partName).Installed = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    If Workbooks.count = 1 Then Workbooks.Add
    ThisWorkbook.SaveAs addinsPath & partName & ".xlam", FileFormat:=xlOpenXMLAddIn
    AddIns.Add fileName:=addinsPath & partName & ".xlam"
    AddIns(partName).Installed = True
NormalExit:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    Call MsgBox("The program is installed successfully!", vbInformation, _
                "Installing the add-in:" & partName)
    ThisWorkbook.Close False
    Exit Sub
InstallationAdd_Err:
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    If err.Number = 1004 Then
        MsgBox "To install the add-in, please close this file and run it again.", _
               64, "Installation"
    Else
        MsgBox err.Description & vbCrLf & " Addin installation failed "
    End If
End Sub

Function IsMissingEndIf(TargetWorkbook As Workbook) As Boolean
    '#INCLUDE dp
    Dim s As String
    s = "  If something then" & vbNewLine
    s = s & "  End If 'comm" & vbNewLine
    Dim var, countIf, coundEndIf
    var = Filter(Split(s, vbNewLine), "If")
    countIf = Filter(var, "End If", False)
    countIf = Filter(countIf, "ElseIf", False)
    dp "countif " & UBound(countIf)
    dp countIf
    countEndIf = Filter(var, "End If")
    dp "countendif " & UBound(countEndIf)
    dp countEndIf
    IsMissingEndIf = (UBound(countIf) = UBound(countEndIf))
End Function

Public Function RegexCountMatches(TEXT As String, Pattern As String) As Long
    Dim RE As New RegExp
    RE.Pattern = Pattern
    RE.Global = True
    RE.IgnoreCase = True
    RE.MultiLine = False
    Dim Matches As matchCollection
    Set Matches = RE.Execute(TEXT)
    RegexCountMatches = Matches.count
End Function

Sub VbeSetFont()
    Application.SendKeys "%TO+{TAB}{RIGHT}{TAB}{TAB}{TAB}{TAB}{TAB}{TAB}"
    Application.SendKeys "Consolas {(}Greek{)}"
    Application.SendKeys "{ENTER}"
End Sub

Sub ClearComponent(vbComp As VBComponent)
    vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
End Sub

Function GetCompText(vbComp As VBComponent) As String
    Dim CodeMod As CodeModule
    Set CodeMod = vbComp.CodeModule
    If CodeMod.CountOfLines = 0 Then GetCompText = "": Exit Function
    GetCompText = CodeMod.Lines(1, CodeMod.CountOfLines)
End Function

Function getModuleName(Module As VBComponent) As String
    '#INCLUDE GetSheetByCodeName
    '#INCLUDE WorkbookOfModule
    If Module.Type = vbext_ct_Document Then
        If Module.Name = "ThisWorkbook" Then
            getModuleName = Module.Name
        Else
            getModuleName = GetSheetByCodeName(WorkbookOfModule(Module), Module.Name).Name
        End If
    Else
        getModuleName = Module.Name
    End If
End Function

Sub GoToModule(Module As VBComponent)
    Application.VBE.MainWindow.visible = True
    Module.CodeModule.CodePane.Window.SetFocus
    Module.CodeModule.CodePane.SetSelection 1, 1, 1, 1
End Sub

Sub EnumToCase()
    Rem point inside enum before calling this from immediate window or vbe menu button
    '#INCLUDE ActiveEnumName
    '#INCLUDE ActiveEnumStartLine
    '#INCLUDE ActiveEnumEndLine
    '#INCLUDE ActiveModule
    '#INCLUDE CLIP
    Dim enumName As String
    enumName = ActiveEnumName
    Dim arr
    arr = Split(ActiveModule.CodeModule.Lines(ActiveEnumStartLine + 1, ActiveEnumEndLine - ActiveEnumStartLine - 1), vbNewLine)
    Dim out As String
    out = "Select case Variable "
    Dim Code As String
    Code = out
    Dim i As Long
    For i = 0 To UBound(arr)
        If InStr(1, arr(i), "=") > 0 Then arr(i) = Split(arr(i), "=")(0)
        arr(i) = Trim(arr(i))
    Next
    For i = 0 To UBound(arr)
        If arr(i) <> "" Then
            out = "    Case is = " & enumName & "." & arr(i) & vbNewLine
            Code = IIf(Code = "", out, Code & vbNewLine & out)
        End If
    Next
    Code = Code & vbNewLine & "End Select"
    Debug.Print Code
    CLIP Code
End Sub

Function ActiveEnumName() As String
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE ActiveModule
    Dim i As Long
    Dim enumName As String
    Dim line As String
    Dim cp As CodeModule
    Set cp = ActiveModule.CodeModule
    For i = CodePaneSelectionStartLine To 1 Step -1
        line = cp.Lines(i, 1)
        If InStr(1, line, "Enum ") > 0 Then
            enumName = Trim(Split(line, " ")(1))
            ActiveEnumName = enumName
            Exit Function
        End If
    Next
End Function

Function ActiveEnumStartLine() As Long
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE ActiveModule
    Dim i As Long
    Dim enumName As String
    Dim line As String
    Dim cp As CodeModule
    Set cp = ActiveModule.CodeModule
    For i = CodePaneSelectionStartLine To 1 Step -1
        line = cp.Lines(i, 1)
        If InStr(1, line, "Enum ") > 0 Then
            ActiveEnumStartLine = i
            Exit Function
        End If
    Next
End Function

Function ActiveEnumEndLine() As Long
    '#INCLUDE CodePaneSelectionStartLine
    '#INCLUDE ActiveModule
    Dim i As Long
    Dim enumName As String
    Dim line As String
    Dim cp As CodeModule
    Set cp = ActiveModule.CodeModule
    For i = CodePaneSelectionStartLine To cp.CountOfLines
        line = cp.Lines(i, 1)
        If InStr(1, line, "End Enum") > 0 Then
            ActiveEnumEndLine = i
            Exit Function
        End If
    Next
End Function

Sub SideBySide(Optional Module1 As VBComponent, _
               Optional Module2 As VBComponent)
    '#INCLUDE ActiveModule
    If Module1 Is Nothing Then Set Module1 = ActiveModule
    With Module1.CodeModule.CodePane.Window
        .Width = 800
        .left = 1
        .top = 1
        .Height = 1000
        .visible = True
        .WindowState = vbext_ws_Normal
        .SetFocus
    End With
    If Module1.Type = vbext_ct_MSForm Then
        With Module1.DesignerWindow
            .Width = 800
            .left = 800
            .top = 1
            .Height = 1000
            .visible = True
            .WindowState = vbext_ws_Normal
            Module1.DesignerWindow.SetFocus
        End With
    ElseIf Not Module2 Is Nothing Then
        With Module2.CodeModule.CodePane.Window
            .Width = 800
            .left = 800
            .top = 1
            .Height = 1000
            .visible = True
            .WindowState = vbext_ws_Normal
            .SetFocus
        End With
    End If
End Sub

Rem  Procedure : DuplicateUserForm
Rem  Author    : Daniel Pineault, CARDA Consultants Inc.
Rem  Website   : http://www.cardaconsultants.com
Rem  Purpose   : Duplicate an existing Userform
Rem  Copyright : The following is release as Attribution-ShareAlike 4.0 International
Rem              (CC BY-SA 4.0) - https://creativecommons.org/licenses/by-sa/4.0/
Rem  Reqrem d Refs: None required
Rem
Rem  Input Variables:
Rem  ~~~~~~~~~~~~~~~~
Rem  sUsrFrmName           : (string) Name of the Userform to create a copy of
Rem  sNewUsrFrmName        : (string) Name to be given to the new copy
Rem  bActivateNewUserFrm   : (boolean) True => Activate/Set the focus of the VBE on the
Rem                                            New Userformonce created
Rem                                    False => Leave the VBE display unchanged
Rem
Rem  Usage:
Rem  ~~~~~~
Rem  Call DuplicateUserForm "UserForm1", "ChartOptions"
Rem    Returns -> True/False; True = Successfully duplication, False = Something went wrong
Rem
Rem  Revision History:
Rem  Rev       Date(yyyy-mm-dd)        Description
Rem  **************************************************************************************
Rem  1         2021-09-25              Initial Public Release
Public Function DuplicateUserForm(Optional sUsrFrmName As String, _
                                  Optional sNewUsrFrmName As String, _
                                  Optional bActivateNewUserFrm As Boolean = True) As Boolean
    '#INCLUDE ActiveModule
    If sUsrFrmName = "" Then
        If ActiveModule.Type <> vbext_ct_MSForm Then
            MsgBox "No Form name passed and active module not userform"
            Exit Function
        Else
            sUsrFrmName = ActiveModule.Name
        End If
    End If
    If sNewUsrFrmName = "" Then sNewUsrFrmName = sUsrFrmName & "_Copy"
    On Error GoTo Error_Handler
    Dim sNewUsrFrmFileName    As String
    sNewUsrFrmFileName = Environ("Temp") & "\" & sNewUsrFrmName & ".frm"
    ThisWorkbook.VBProject.VBComponents(sUsrFrmName).Name = sNewUsrFrmName
    ThisWorkbook.VBProject.VBComponents(sNewUsrFrmName).Export sNewUsrFrmFileName
    ThisWorkbook.VBProject.VBComponents(sNewUsrFrmName).Name = sUsrFrmName
    ThisWorkbook.VBProject.VBComponents.Import sNewUsrFrmFileName
    If Len(Dir(sNewUsrFrmFileName)) > 0 Then Kill Replace(sNewUsrFrmFileName, ".frm", ".*")
    If bActivateNewUserFrm = True Then ThisWorkbook.VBProject.VBComponents(sNewUsrFrmName).Activate
    DuplicateUserForm = True
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    MsgBox "The following error has occured" & vbCrLf & vbCrLf & _
           "Error Number: " & err.Number & vbCrLf & _
           "Error Source: DuplicateUserForm" & vbCrLf & _
           "Error Description: " & err.Description & _
           Switch(Erl = 0, "", Erl <> 0, vbCrLf & "Line No: " & Erl) _
           , vbOKOnly + vbCritical, "An Error has Occured!"
    Resume Error_Handler_Exit
End Function

Public Sub CopyProcedureFromThisWorkbook(Optional ProcedureName As String, Optional TargetWorkbook As Workbook)
    Rem for each element in Array("ActiveModule","ProcListArray","GetProcText"):CopyProcedureFromThisWorkbook cstr(element):next
    '#INCLUDE ProcListArray
    '#INCLUDE ActiveModule
    '#INCLUDE ActiveProcedure
    '#INCLUDE CreateOrSetModule
    '#INCLUDE GetModuleText
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveWorkbook
    If TargetWorkbook.Name = ThisWorkbook.Name Then Exit Sub
    Dim Module As VBComponent:         Set Module = CreateOrSetModule("vbArc", vbext_ct_StdModule, ActiveWorkbook)
    Dim ProcedureText As String:       ProcedureText = GetProcText(ModuleOfProcedure(ThisWorkbook, ProcedureName), ProcedureName)
    ProcedureName = IIf(InStr(1, ProcedureText, "Function " & ProcedureName) > 0, "Function ", "Sub ") & ProcedureName
    If InStr(1, GetModuleText(Module), ProcedureName, vbTextCompare) = 0 Then Module.CodeModule.AddFromString (ProcedureText)
End Sub

Public Sub CopyProcedures(ProcedureName As Variant, FromWorkbook As Workbook, TargetWorkbook As Workbook, Optional Overwrite As Boolean)
    '#INCLUDE UpdateProcedureCode
    '#INCLUDE LinkedProcs
    '#INCLUDE ProcedureExists
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE ActiveProcedure
    '#INCLUDE CreateOrSetModule
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    Dim Module As VBComponent
    Dim Code As String
    If ProcedureName = "" Then ProcedureName = ActiveProcedure
    If FromWorkbook Is Nothing Then Set FromWorkbook = ActiveCodepaneWorkbook
    Dim procedures As Collection: Set procedures = LinkedProcs(ProcedureName, FromWorkbook)
    On Error Resume Next
    procedures.Add ProcedureName, ProcedureName
    On Error GoTo 0
    Dim Procedure As Variant
    Dim FileFullName As String
    For Each Procedure In procedures
        Code = GetProcText(ModuleOfProcedure(FromWorkbook, CStr(Procedure)), CStr(Procedure))
        If ProcedureExists(CStr(Procedure), TargetWorkbook) = False Then
            Set Module = CreateOrSetModule("vbArcImports", vbext_ct_StdModule, TargetWorkbook)
            Module.CodeModule.AddFromString (Code)
        Else
            If Overwrite = True Then UpdateProcedureCode Procedure, Code, TargetWorkbook
        End If
    Next
End Sub

Function ProceduresOfWorkbook( _
         TargetWorkbook As Workbook, _
         Optional ExcludeDocument As Boolean = True, _
         Optional ExcludeClass As Boolean = True, _
         Optional ExcludeForm As Boolean = True) As Collection
    Dim Module As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long
    Dim coll As New Collection
    Dim ProcedureName As String
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If ExcludeClass = True Then
            If Module.Type = vbext_ct_ClassModule Then GoTo Skip
        End If
        If ExcludeDocument = True Then
            If Module.Type = vbext_ct_Document Then GoTo Skip
        End If
        If ExcludeForm = True Then
            If Module.Type = vbext_ct_MSForm Then GoTo Skip
        End If
        With Module.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                ProcedureName = .ProcOfLine(LineNum, ProcKind)
                coll.Add ProcedureName
                LineNum = .ProcStartLine(ProcedureName, ProcKind) + .ProcCountLines(ProcedureName, ProcKind) + 1
            Loop
        End With
Skip:
    Next Module
    Set ProceduresOfWorkbook = coll
End Function

Function ProcedureExists( _
         ProcedureName As Variant, _
         FromWorkbook As Workbook) _
        As Boolean
    '#INCLUDE ProceduresOfWorkbook
    Dim AllProcedures As Collection: Set AllProcedures = ProceduresOfWorkbook(FromWorkbook)
    Dim Procedure As Variant
    For Each Procedure In AllProcedures
        If UCase(CStr(Procedure)) = UCase(ProcedureName) Then
            ProcedureExists = True
            Exit Function
        End If
    Next
End Function

Function ProcedureFirstLine(Module As VBComponent, procName As String) As Long
    '#INCLUDE InStrExact
    Dim n As Long
    Dim s As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    For n = Module.CodeModule.ProcBodyLine(procName, ProcKind) + IIf(n = 0, 1, 0) To Module.CodeModule.CountOfLines
        s = Trim(Module.CodeModule.Lines(n, 1))
        If s = vbNullString Then
            Exit For
        ElseIf left(s, 1) = "'" Then
        ElseIf left(s, 3) = "Rem" Then
        ElseIf Right(Trim(Module.CodeModule.Lines(n - 1, 1)), 1) = "_" Then
        ElseIf Right(s, 1) = "_" Then
        ElseIf InStrExact(1, s, "Sub ") Then
        ElseIf InStrExact(1, s, "Function ") Then
        Else
            Exit For
        End If
    Next n
    ProcedureFirstLine = n
End Function

Public Function IsEditorInSync() As Boolean
    With Application.VBE
        IsEditorInSync = .ActiveVBProject Is _
                         .ActiveCodePane.CodeModule.parent.Collection.parent
    End With
End Function

Sub SyncVBAEditor()
    With Application.VBE
        If Not .ActiveCodePane Is Nothing Then
            Set .ActiveVBProject = .ActiveCodePane.CodeModule.parent.Collection.parent
        End If
    End With
End Sub

Public Function vbModule(ModuleName As String, Optional TargetWorkbook As Workbook) As VBComponent
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    On Error Resume Next
    Set vbModule = TargetWorkbook.VBProject.VBComponents(ModuleName)
End Function

Function ProcListArray(Module As VBComponent) As Variant
    Dim out As String
    Dim LineNum As Long
    Dim NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    LineNum = Module.CodeModule.CountOfDeclarationLines + 1
    Do Until LineNum >= Module.CodeModule.CountOfLines
        procName = Module.CodeModule.ProcOfLine(LineNum, ProcKind)
        If out = vbNullString Then
            out = procName
        Else
            out = out & "," & procName
        End If
        LineNum = Module.CodeModule.ProcStartLine(procName, ProcKind) + Module.CodeModule.ProcCountLines(procName, ProcKind) + 1
    Loop
    ProcListArray = Split(out, ",")
End Function

Function ProcListCollection(Module As VBComponent) As Collection
    Dim coll As Collection: Set coll = New Collection
    Dim LineNum As Long, NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    LineNum = Module.CodeModule.CountOfDeclarationLines + 1
    Do Until LineNum >= Module.CodeModule.CountOfLines
        procName = Module.CodeModule.ProcOfLine(LineNum, ProcKind)
        If InStr(1, procName, "_") = 0 Then coll.Add procName
        LineNum = Module.CodeModule.ProcStartLine(procName, ProcKind) + Module.CodeModule.ProcCountLines(procName, ProcKind) + 1
    Loop
    Set ProcListCollection = coll
End Function

Public Function ProcedureEndLine(Module As VBComponent, procName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim startAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
    startAt = Module.CodeModule.ProcStartLine(procName, ProcKind)
    EndAt = Module.CodeModule.ProcStartLine(procName, ProcKind) + Module.CodeModule.ProcCountLines(procName, ProcKind) - 1
    CountOf = Module.CodeModule.ProcCountLines(procName, ProcKind)
    ProcedureEndLine = EndAt
End Function

Public Function ProcedureStartLine(Module As VBComponent, ProcedureName As String) As Long
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim startAt As Long
    Dim EndAt As Long
    Dim CountOf As Long
    startAt = Module.CodeModule.ProcStartLine(ProcedureName, ProcKind)
    EndAt = Module.CodeModule.ProcStartLine(ProcedureName, ProcKind) + Module.CodeModule.ProcCountLines(ProcedureName, ProcKind) - 1
    CountOf = Module.CodeModule.ProcCountLines(ProcedureName, ProcKind)
    ProcedureStartLine = startAt
End Function

Sub DebugPrintVbeWindows()
    Dim wd As Window
    Dim i As Long
    For i = 1 To Application.VBE.Windows.count
        With Application.VBE.Windows(i)
            Debug.Print .Type & vbTab & .Caption
        End With
    Next
End Sub

Function vbeWindowProjectExplorer() As VBIDE.Window
    Dim i As Long
    For i = 1 To Application.VBE.Windows.count
        If Application.VBE.Windows(i).Caption Like "Project - *" Then
            Set vbeWindowProjectExplorer = Application.VBE.Windows(i)
        End If
    Next
End Function

Function vbeWindowProperties() As VBIDE.Window
    Dim i As Long
    For i = 1 To Application.VBE.Windows.count
        If Application.VBE.Windows(i).Caption Like "Properties - *" Then
            Set vbeWindowProperties = Application.VBE.Windows(i)
        End If
    Next
End Function

Function vbeWindowImmediate() As VBIDE.Window
    Set vbeWindowImmediate = Application.VBE.Windows("Immediate")
End Function

Function vbeWindowLocals() As VBIDE.Window
    Set vbeWindowLocals = Application.VBE.Windows("Locals")
End Function

Function vbeWindowWatches() As VBIDE.Window
    Set vbeWindowWatches = Application.VBE.Windows("Watches")
End Function

Function vbeWindowObjectBrowser() As VBIDE.Window
    Set vbeWindowObjectBrowser = Application.VBE.Windows("Object Browser")
End Function

Sub CloseVBEwindows()
    Dim i As Long
    For i = 1 To Application.VBE.Windows.count
        Select Case Application.VBE.Windows(i).Type
            Case 2 To 8
                Application.VBE.Windows(i).Close
        End Select
    Next
End Sub

Sub OpenVBEwindows()
    Dim i As Long
    For i = 1 To Application.VBE.Windows.count
        Select Case Application.VBE.Windows(i).Type
            Case 5 To 7
                Application.VBE.Windows(i).visible = True
        End Select
    Next
End Sub

Function ProcedureLines(Procedure As String, Module As VBComponent) As Collection
    Dim i As Long
    Dim out As New Collection
    For i = 1 To Module.CodeModule.CountOfLines
        If out.count = 0 Then
            If InStr(1, Module.CodeModule.Lines(i, 1), "Sub " & Procedure) > 0 Then out.Add i
            If InStr(1, Module.CodeModule.Lines(i, 1), "Function " & Procedure) > 0 Then out.Add i
        Else
            If InStr(1, Module.CodeModule.Lines(i, 1), "End Sub") > 0 Then
                out.Add i
                Exit For
            End If
            If InStr(1, Module.CodeModule.Lines(i, 1), "End Function") > 0 Then
                out.Add i
                Exit For
            End If
        End If
    Next
    out.Add out(2) - out(1) + 1
    Set ProcedureLines = out
    Debug.Print Procedure
    Debug.Print "Start line = " & out(1)
    Debug.Print "End line = " & out(2)
    Debug.Print "Count of lines = " & out(3)
    Set out = Nothing
End Function

Function ProtectedVBProject(ByVal wb As Workbook) As Boolean
    If wb.VBProject.Protection = 1 Then
        ProtectedVBProject = True
    Else
        ProtectedVBProject = False
    End If
End Function

Sub AddExtensibility(Optional TargetWorkbook As Workbook)
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    On Error Resume Next
    TargetWorkbook.VBProject.REFERENCES.AddFromGuid _
        GUID:="{0002E157-0000-0000-C000-000000000046}", _
        Major:=5, Minor:=3
    On Error GoTo 0
End Sub

Function ProcedureType(TargetWorkbook As Workbook, ProcedureName As String) As String
    '#INCLUDE GetProcText
    '#INCLUDE ModuleOfProcedure
    Dim ProcedureText As String:            ProcedureText = GetProcText(ModuleOfProcedure(TargetWorkbook, ProcedureName), ProcedureName)
    ProcedureType = "Null"
    If InStr(1, ProcedureText, "Sub " & ProcedureName) > 0 Then
        ProcedureType = "Sub"
    ElseIf InStr(1, ProcedureText, "Function " & ProcedureName) > 0 Then
        ProcedureType = "Function"
    End If
End Function

Public Function ActiveCodepaneWorkbook() As Workbook
    Dim TmpStr As String
    TmpStr = Application.VBE.SelectedVBComponent.Collection.parent.fileName
    TmpStr = Right(TmpStr, Len(TmpStr) - InStrRev(TmpStr, "\"))
    Set ActiveCodepaneWorkbook = Workbooks(TmpStr)
End Function

Public Function ActiveModule() As VBComponent
    Set ActiveModule = Application.VBE.SelectedVBComponent
End Function

Public Function ActiveProcedure() As String
    Application.VBE.ActiveCodePane.GetSelection L1&, C1&, L2&, C2&
    ActiveProcedure = Application.VBE.ActiveCodePane _
                      .CodeModule.ProcOfLine(L1&, vbext_pk_Proc)
End Function

Sub RemoveAllCodeFromModule(Module As VBComponent)
    On Error Resume Next
    Module.CodeModule.DeleteLines 1, Module.CodeModule.CountOfLines + 1
End Sub

Function CreateOrSetModule(compName As String, compType As VBIDE.vbext_ComponentType, Optional TargetWorkbook As Workbook) As VBComponent
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim vbComp As VBComponent
    On Error Resume Next
    Set vbComp = TargetWorkbook.VBProject.VBComponents(compName)
    On Error GoTo 0
    If vbComp Is Nothing Then
        Set vbComp = TargetWorkbook.VBProject.VBComponents.Add(compType)
        vbComp.Name = compName
    End If
    Set CreateOrSetModule = vbComp
End Function

Function ComponentTypeToString(componentType As VBIDE.vbext_ComponentType) As String
    Select Case componentType
        Case vbext_ct_ActiveXDesigner
            ComponentTypeToString = "ActiveX Designer"
        Case vbext_ct_ClassModule
            ComponentTypeToString = "Class Module"
        Case vbext_ct_Document
            ComponentTypeToString = "Document Module"
        Case vbext_ct_MSForm
            ComponentTypeToString = "UserForm"
        Case vbext_ct_StdModule
            ComponentTypeToString = "Code Module"
        Case Else
            ComponentTypeToString = "Unknown Type: " & CStr(componentType)
    End Select
End Function

Function ModuleOfWorksheet(TargetSheet As Worksheet) As VBComponent
    Set ModuleOfWorksheet = TargetSheet.parent.VBProject.VBComponents(TargetSheet.CodeName)
End Function

Sub CopyModulelToAllMyProjects(ModuleName As String, Overwrite As Boolean)
    '#INCLUDE dp
    '#INCLUDE ProtectedVBProject
    '#INCLUDE CopyModule
    Dim X, Y As Variant
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    If InStr(1, Y.VBProject.Name, "vbArc", vbTextCompare) > 0 Then
                        If Y.Name <> ThisWorkbook.Name Then
                            dp Y.Name
                            CopyModule ModuleName, ThisWorkbook, Workbooks(Y.Name), Overwrite
                        End If
                    End If
                End If
            End If
            err.clear
        Next
    Next
End Sub

Function CopyModule(ModuleName As String, _
                    FromWorkbook As Workbook, _
                    toWorkbook As Workbook, _
                    OverwriteExisting As Boolean) As Boolean
    Dim FromVBProject As VBIDE.VBProject
    Set FromVBProject = FromWorkbook.VBProject
    Dim ToVBProject As VBIDE.VBProject
    Set ToVBProject = toWorkbook.VBProject
    If ToVBProject.Name = FromVBProject.Name Then Exit Function
    Dim vbComp As VBIDE.VBComponent
    Dim FName As String
    Dim TempVBComp As VBIDE.VBComponent
    On Error Resume Next
    Set vbComp = FromVBProject.VBComponents(ModuleName)
    If vbComp Is Nothing Then
        err.Raise 438, , "No Component by this name"
    End If
    Dim EXT As String
    FName = Environ("Temp") & "\" & ModuleName
    Dim vbcomp2 As VBComponent
    Set vbcomp2 = ToVBProject.VBComponents(ModuleName)
    If OverwriteExisting = True Then
        If Dir(FName, vbNormal + vbHidden + vbSystem) <> vbNullString Then
            err.clear
            Kill FName
            If err.Number <> 0 Then
                CopyModule = False
                Exit Function
            End If
        End If
        If Not vbcomp2 Is Nothing Then
            ToVBProject.VBComponents.Remove vbcomp2
        End If
    Else
        If vbcomp2 Is Nothing Then
        Else
            CopyModule = False
            Exit Function
        End If
    End If
    vbComp.Export fileName:=FName
    If vbComp.Type = vbext_ct_Document Then
        Set TempVBComp = ToVBProject.VBComponents.Import(FName)
        With vbComp.CodeModule
            .DeleteLines 1, .CountOfLines
            s = TempVBComp.CodeModule.Lines(1, TempVBComp.CodeModule.CountOfLines)
            .InsertLines 1, s
        End With
        On Error GoTo 0
        ToVBProject.VBComponents.Remove TempVBComp
    Else
        ToVBProject.VBComponents.Import fileName:=FName
    End If
    Kill FName
    CopyModule = True
End Function

Sub DeleteComponent(vbComp As VBComponent)
    '#INCLUDE GetSheetByCodeName
    '#INCLUDE WorkbookOfModule
    Application.DisplayAlerts = False
    If vbComp.Type = vbext_ct_Document Then
        If vbComp.Name = "ThisWorkbook" Then
            vbComp.CodeModule.DeleteLines 1, vbComp.CodeModule.CountOfLines
        Else
            If WorkbookOfModule(vbComp).SHEETS.count > 1 Then
                GetSheetByCodeName(WorkbookOfModule(vbComp), vbComp.Name).Delete
            Else
                Dim ws As Worksheet
                Set ws = WorkbookOfModule(vbComp).SHEETS.Add
                ws.Name = "All other sheets were deleted"
                GetSheetByCodeName(WorkbookOfModule(vbComp), vbComp.Name).Delete
            End If
        End If
    Else
        WorkbookOfModule(vbComp).VBProject.VBComponents.Remove vbComp
    End If
    Application.DisplayAlerts = True
End Sub

Function GetModuleText(vbComp As VBComponent) As String
    Dim CodeMod As CodeModule
    Set CodeMod = vbComp.CodeModule
    If CodeMod.CountOfLines = 0 Then GetModuleText = "": Exit Function
    GetModuleText = CodeMod.Lines(1, CodeMod.CountOfLines)
End Function

Function GetProjectText(TargetWorkbook As Workbook) As String
    '#INCLUDE getModuleName
    '#INCLUDE GetModuleText
    Dim Module As VBComponent
    Dim txt
    Dim div As String
    div = vbNewLine & "'=============================================" & vbNewLine
    For Each Module In TargetWorkbook.VBProject.VBComponents
        If Module.CodeModule.CountOfLines > 0 Then
            txt = txt & div & _
                  "'    " & getModuleName(Module) & "    " & Module.Type & _
                  div & _
                  GetModuleText(Module)
        End If
    Next
    GetProjectText = txt
End Function

Public Function getProcKind(vbComp As VBComponent, ByVal sProcName As String) As Long
    '#INCLUDE GetProcedureDeclaration
    Dim codeMode As CodeModule
    Set codeMode = vbComp.CodeModule
    Const vbext_pk_Proc As Long = 0
    Const vbext_pk_Let As Long = 1
    Const vbext_pk_Set As Long = 2
    Const vbext_pk_Get As Long = 3
    Dim txt As String
    txt = GetProcedureDeclaration(vbComp, sProcName, 0)
    If InStr(1, txt, "Get " & sProcName) > 0 Then
        getProcKind = 3
    ElseIf InStr(1, txt, "Let " & sProcName) > 0 Then
        getProcKind = 1
    ElseIf InStr(1, txt, "Set " & sProcName) > 0 And Not (InStr(1, txt, "Sub " & sProcName) > 0 Or InStr(1, txt, "Function " & sProcName) > 0) Then
        getProcKind = 2
    Else
        getProcKind = 0
    End If
End Function

Public Function GetProcText(vbComp As VBComponent, _
                            sProcName As Variant, _
                            Optional bInclHeader As Boolean = True) As String
    If vbComp Is Nothing Then
        Stop
    End If
    Dim CodeMod As CodeModule
    Set CodeMod = vbComp.CodeModule
    Dim lProcStart            As Long
    Dim lProcBodyStart        As Long
    Dim lProcNoLines          As Long
    Const vbext_pk_Proc = 0
    On Error GoTo Error_Handler
    lProcStart = CodeMod.ProcStartLine(sProcName, vbext_pk_Proc)
    lProcBodyStart = CodeMod.ProcBodyLine(sProcName, vbext_pk_Proc)
    lProcNoLines = CodeMod.ProcCountLines(sProcName, vbext_pk_Proc)
    If bInclHeader = True Then
        GetProcText = CodeMod.Lines(lProcStart, lProcNoLines)
    Else
        lProcNoLines = lProcNoLines - (lProcBodyStart - lProcStart)
        GetProcText = CodeMod.Lines(lProcBodyStart, lProcNoLines)
    End If
Error_Handler_Exit:
    On Error Resume Next
    Exit Function
Error_Handler:
    Rem debug.Print _
    "Error Source: GetProcText" & vbCrLf & _
    "Error Description: " & err.Description & _
    Switch(Erl = 0, vbNullString, Erl <> 0, vbCrLf & "Line No: " & Erl)
    Resume Error_Handler_Exit
End Function

Public Function GetSheetByCodeName(wb As Workbook, CodeName As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If UCase(sh.CodeName) = UCase(CodeName) Then Set GetSheetByCodeName = sh: Exit For
    Next sh
End Function

Public Sub GotoFirstModule()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    Application.VBE.MainWindow.visible = True
    Application.VBE.MainWindow.WindowState = vbext_ws_Maximize
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then
            vbComp.Activate
            vbComp.CodeModule.CodePane.SetSelection 1, 1, 1, 1
            Exit Sub
        End If
    Next vbComp
End Sub

Public Function IsBlockEnd(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
        Case "Next", "Loop", "Wend", "End Select", "Case", "Else", "#Else", "Else:", "#Else:", "ElseIf", "#ElseIf", "End If", "#End If"
            bOK = True
        Case "End"
            bOK = (Len(strLine) > 3)
    End Select
    IsBlockEnd = bOK
End Function

Public Function IsBlockStart(strLine As String) As Boolean
    Dim bOK As Boolean
    Dim nPos As Integer
    Dim strTemp As String
    nPos = InStr(1, strLine, " ") - 1
    If nPos < 0 Then nPos = Len(strLine)
    strTemp = left$(strLine, nPos)
    Select Case strTemp
        Case "With", "For", "Do", "While", "Select", "Case", "Else", "Else:", "#Else", "#Else:", "Sub", "Function", "Property", "Enum", "Type"
            bOK = True
        Case "If", "#If", "ElseIf", "#ElseIf"
            bOK = (Len(strLine) = (InStr(1, strLine, " Then") + 4))
        Case "public", "Public", "Friend"
            nPos = InStr(1, strLine, " Static ")
            If nPos Then
                nPos = InStr(nPos + 7, strLine, " ")
            Else
                nPos = InStr(Len(strTemp) + 1, strLine, " ")
            End If
            Select Case Mid$(strLine, nPos + 1, InStr(nPos + 1, strLine, " ") - nPos - 1)
                Case "Sub", "Function", "Property", "Enum", "Type"
                    bOK = True
            End Select
    End Select
    IsBlockStart = bOK
End Function

Sub ListProceduresSeparate(Optional wb As Workbook)
    '#INCLUDE ProcListArray
    '#INCLUDE CreateOrSetSheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim vbComp As VBComponent
    Dim procArray As Variant
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("ListOfProcedures", wb)
    ws.Cells.clear
    Dim cell As Range
    Set cell = ws.Range("B1")
    For Each vbComp In wb.VBProject.VBComponents
        cell = vbComp.Name
        procArray = ProcListArray(vbComp)
        If UBound(procArray) <> -1 Then
            cell.OFFSET(1).RESIZE(UBound(procArray) + 1) = WorksheetFunction.Transpose(procArray)
        End If
        Set cell = cell.OFFSET(0, 1)
    Next
    ws.Cells.Columns.AutoFit
End Sub

Sub ListProceduresUnified(Optional wb As Workbook)
    '#INCLUDE ProcListArray
    '#INCLUDE CreateOrSetSheet
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim vbComp As VBComponent
    Dim procArray As Variant
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("ListOfProcedures", ThisWorkbook)
    ws.Cells.clear
    Dim cell As Range
    Set cell = ws.Range("B1")
    For Each vbComp In wb.VBProject.VBComponents
        procArray = ProcListArray(vbComp)
        If UBound(procArray) <> -1 Then
            cell.RESIZE(UBound(procArray) + 1) = WorksheetFunction.Transpose(procArray)
            Set cell = ws.Range("B" & rows.count).End(xlUp).OFFSET(1)
        End If
    Next
    ws.Cells.Columns.AutoFit
    ws.Columns(2).SpecialCells(xlCellTypeConstants).Sort Key1:=Range("B1"), order1:=xlAscending, header:=xlNo
End Sub

Sub getAllWorkbooksLinks()
    '#INCLUDE dp
    '#INCLUDE ProtectedVBProject
    '#INCLUDE WorkbookLinks
    Dim X, Y As Variant
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    dp Y.Name
                    dp "---------------"
                    dp WorkbookLinks(Workbooks(Y.Name))
                End If
            End If
            err.clear
        Next
    Next
End Sub

Function WorkbookLinks(TargetWorkbook As Workbook) As Collection
    Dim coll As New Collection
    Dim aLinks As Variant
    Dim el
    For Each el In Array(xlExcelLinks, xlOLELinks)
        aLinks = TargetWorkbook.LinkSources(el)
        If Not IsEmpty(aLinks) Then
            For i = 1 To UBound(aLinks)
                coll.Add aLinks(i)
            Next i
        End If
    Next
    Set WorkbookLinks = coll
End Function

Sub MacroLinkRemoverActiveWorkbook()
    '#INCLUDE MacroLinkRemover
    MacroLinkRemover ActiveWorkbook
End Sub

Sub MacroLinkRemoverActiveCodepaneWorkbook()
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE MacroLinkRemover
    MacroLinkRemover ActiveCodepaneWorkbook
End Sub

Sub MacroLinkRemover(TargetWorkbook As Workbook)
    '#INCLUDE FixExternalShapes
    '#INCLUDE FixExternalNames
    FixExternalShapes TargetWorkbook
    FixExternalNames TargetWorkbook
End Sub

Sub FixExternalShapes(Optional TargetWorkbook As Workbook)
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim shp As Shape
    Dim MacroLink, NewLink As String
    Dim SplitLink As Variant
    Dim ws As Worksheet
    For Each ws In TargetWorkbook.SHEETS
        For Each shp In ws.Shapes
            On Error GoTo NextShp
            MacroLink = shp.OnAction
            If MacroLink <> "" And InStr(MacroLink, "!") <> 0 Then
                SplitLink = Split(MacroLink, "!")
                NewLink = SplitLink(1)
                If Right(NewLink, 1) = "'" Then
                    NewLink = left(NewLink, Len(NewLink) - 1)
                End If
                shp.OnAction = NewLink
            End If
NextShp:
        Next shp
    Next
End Sub

Sub testFixNames()
    '#INCLUDE ProtectedVBProject
    '#INCLUDE FixExternalNames
    Dim X, Y As Variant
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    FixExternalNames Workbooks(Y.Name)
                End If
            End If
            err.clear
        Next
    Next
End Sub

Sub FixExternalNames(Optional TargetWorkbook As Workbook)
    '#INCLUDE ActiveCodepaneWorkbook
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim objDefinedName As Name
    For Each objDefinedName In TargetWorkbook.Names
        If InStr(objDefinedName.RefersTo, "[") > 0 Then
            With objDefinedName
                .RefersTo = IIf(InStr(1, .RefersTo, "'") > 0, "'", "") & Replace(Mid(.RefersTo, InStrRev(.RefersTo, "]") + 1), """", "")
            End With
        End If
    Next objDefinedName
End Sub

Sub BreakAllLinks(ByVal wb As Object)
    Dim link As Variant, LinkType As Variant
    For Each LinkType In Array(xlLinkTypeExcelLinks, xlOLELinks, xlPublishers, xlSubscribers)
        If Not IsEmpty(wb.LinkSources(Type:=LinkType)) Then
            For Each link In wb.LinkSources(Type:=LinkType)
                wb.BreakLink Name:=link, Type:=LinkType
            Next link
        End If
    Next LinkType
    wb.UpdateLinks = xlUpdateLinksNever
End Sub

Public Function ModuleExists(Name As String, Optional ByVal ExistsInWorkbook As Workbook) As Boolean
    Dim j As Long
    Dim vbComp As VBComponent
    Dim modules As Collection
    Set modules = New Collection
    ModuleExists = False
    If ExistsInWorkbook Is Nothing Then
        Set ExistsInWorkbook = ThisWorkbook
    End If
    If (Name = vbNullString) Then
        GoTo errorname
    End If
    For Each vbComp In ExistsInWorkbook.VBProject.VBComponents
        If ((vbComp.Type = vbext_ct_StdModule) Or (vbComp.Type = vbext_ct_ClassModule)) Then
            modules.Add vbComp.Name
        End If
    Next vbComp
    For j = 1 To modules.count
        If (Name = modules.item(j)) Then
            ModuleExists = True
        End If
    Next j
    j = 0
    If (ModuleExists = False) Then
        GoTo NotFound
    End If
    If (0 <> 0) Then
errorname:
        MsgBox ("Function BootStrap.Is_Module_Loaded Was not passed a Name of Module")
        Exit Function
    End If
    If (0 <> 0) Then
NotFound:
        Exit Function
    End If
End Function

Function ModuleOfProcedure(wb As Workbook, ProcedureName As Variant) As VBComponent
    Dim ProcKind As VBIDE.vbext_ProcKind
    Dim LineNum As Long, NumProc As Long
    Dim procName As String
    Dim vbComp As VBComponent
    For Each vbComp In wb.VBProject.VBComponents
        With vbComp.CodeModule
            LineNum = .CountOfDeclarationLines + 1
            Do Until LineNum >= .CountOfLines
                procName = .ProcOfLine(LineNum, ProcKind)
                LineNum = .ProcStartLine(procName, ProcKind) + .ProcCountLines(procName, ProcKind) + 1
                If UCase(procName) = UCase(ProcedureName) Then
                    Set ModuleOfProcedure = vbComp
                    Exit Function
                End If
            Loop
        End With
    Next vbComp
End Function

Function WorkbookOfProject(vbProj As VBProject) As Workbook
    TmpStr = vbProj.fileName
    TmpStr = Right(TmpStr, Len(TmpStr) - InStrRev(TmpStr, "\"))
    Set WorkbookOfProject = Workbooks(TmpStr)
End Function

Function WorkbookOfModule(vbComp As VBComponent) As Workbook
    '#INCLUDE WorkbookOfProject
    Set WorkbookOfModule = WorkbookOfProject(vbComp.Collection.parent)
End Function


