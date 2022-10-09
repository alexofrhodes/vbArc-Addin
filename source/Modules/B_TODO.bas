Attribute VB_Name = "B_TODO"
Private Sub sysAddTODO()
    Rem Author VBATools
    '#INCLUDE ActiveModule
    Dim Module As VBComponent:      Set Module = ActiveModule
    Dim CurentCodePane As CodePane: Set CurentCodePane = Module.CodeModule.CodePane
    Dim nLine  As Long, lStartLine As Long, lStartColumn As Long, lEndLine As Long, lEndColumn As Long
    Dim sFersLine As String, sSpec As String, txtName As String
    txtName = AUTHOR_NAME
    If txtName = vbNullString Then txtName = Environ("UserName")
    On Error GoTo HELL
    With CurentCodePane
        .GetSelection lStartLine, lStartColumn, lEndLine, lEndColumn
        sSpec = VBA.String$(lStartColumn - 1, " ")
        sFersLine = sSpec & "'* @TODO Created: " & VBA.Format$(Now, ctFormat) & " Author: " & txtName & vbCrLf & sSpec & "'*"
        .CodeModule.InsertLines lStartLine, sFersLine
    End With
HELL:
End Sub

Sub ShowTodo()
    Rem Open Navigation Userform
    '#INCLUDE FindCode
    FindCode "@TODO"
End Sub

Rem @Folder TODO
Sub DebugPrintListOfTodoInModule()
    '#INCLUDE dp
    '#INCLUDE ListOfTodoInModule
    dp ListOfTodoInModule
End Sub

Function ListOfTodoInModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    '#INCLUDE GetModuleText
    '#INCLUDE ArrayMultiFilter
    If Module Is Nothing Then Set Module = ActiveModule
    Dim out As String
    Dim Matches As String
    Matches = Join(ArrayMultiFilter(Split(GetModuleText(Module), vbNewLine), _
                                    Array("@TODO"), True), _
                                    vbNewLine)
    out = out & vbNewLine & "'---------"
    out = out & vbNewLine & "'Module: " & Module.Name
    out = out & vbNewLine & "'---------"
    out = out & vbNewLine & "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    ListOfTodoInModule = out
End Function

Function ListOfToDo(Optional TargetWorkbook As Workbook, Optional includeModuleName As Boolean)
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE GetModuleText
    '#INCLUDE ArrayMultiFilter
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim out As String
    Dim Matches As String
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        Matches = Join(ArrayMultiFilter(Split(GetModuleText(Module), vbNewLine), _
                                        Array("@TODO"), True), _
                                        vbNewLine)
        If Len(Trim(Matches)) <> 0 Then
            If includeModuleName = True Then
                out = out & vbNewLine & "'---------"
                out = out & vbNewLine & "'Module: " & Module.Name
                out = out & vbNewLine & "'---------"
            End If
            out = out & vbNewLine & "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
        End If
    Next
    ListOfToDo = out
End Function

Sub ImportTODOHere()
    Rem Do not move this procedure!!!
    Rem All lines after this procedure will be deleted and replaced.
    '#INCLUDE vbModule
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ListOfToDo
    Dim v
    v = ListOfToDo(, True)
    Dim Module As VBComponent
    Set Module = vbModule("B_TODO", ThisWorkbook)
    Dim ProcEndLine As Long
    ProcEndLine = ProcedureEndLine(Module, "ImportTodoHere")
    With Module.CodeModule
        .DeleteLines ProcEndLine + 1, .CountOfLines - ProcEndLine
        .InsertLines .CountOfLines + 1, vbNewLine & v
    End With
End Sub


