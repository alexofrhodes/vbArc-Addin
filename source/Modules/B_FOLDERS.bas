Attribute VB_Name = "B_FOLDERS"
Rem @Folder Folders
Sub DebugPrintListOfFoldersInModule()
    '#INCLUDE dp
    '#INCLUDE ListOfVbeFoldersInModule
    dp ListOfVbeFoldersInModule
End Sub

Function ListOfVbeFoldersInModule(Optional Module As VBComponent)
    '#INCLUDE ActiveModule
    '#INCLUDE GetModuleText
    '#INCLUDE ArrayMultiFilter
    If Module Is Nothing Then Set Module = ActiveModule
    Dim out As String
    Dim Matches As String
    Matches = Join(ArrayMultiFilter(Split(GetModuleText(Module), vbNewLine), _
                                    Array("@Folder", "@Subfolder"), True), _
                                    vbNewLine)
    If Len(Trim(Matches)) <> 0 Then
        If includeModuleName = True Then
            out = out & vbNewLine & "'---------"
            out = out & vbNewLine & "'Module: " & Module.Name
            out = out & vbNewLine & "'---------"
        End If
        out = out & vbNewLine & "'" & Replace(Matches, vbNewLine, vbNewLine & "'")
    End If
    ListOfVbeFoldersInModule = out
End Function

Function ListOfVbeFolders(Optional TargetWorkbook As Workbook, Optional includeModuleName As Boolean)
    '#INCLUDE ActiveCodepaneWorkbook
    '#INCLUDE GetModuleText
    '#INCLUDE ArrayMultiFilter
    If TargetWorkbook Is Nothing Then Set TargetWorkbook = ActiveCodepaneWorkbook
    Dim out As String
    Dim Matches As String
    Dim Module As VBComponent
    For Each Module In TargetWorkbook.VBProject.VBComponents
        Matches = Join(ArrayMultiFilter(Split(GetModuleText(Module), vbNewLine), _
                                        Array("@Folder", "@Subfolder"), True), _
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
    ListOfVbeFolders = out
End Function

Sub ImportFoldersHere()
    Rem Do not move this procedure!!!
    Rem All lines after this procedure will be deleted and replaced.
    '#INCLUDE vbModule
    '#INCLUDE ProcedureEndLine
    '#INCLUDE ListOfVbeFolders
    Dim v
    v = ListOfVbeFolders(, True)
    Dim Module As VBComponent
    Set Module = vbModule("B_FOLDERS")
    Dim ProcEndLine As Long
    ProcEndLine = ProcedureEndLine(Module, "ImportFoldersHere")
    With Module.CodeModule
        .DeleteLines ProcEndLine + 1, .CountOfLines - ProcEndLine
        .InsertLines .CountOfLines + 1, vbNewLine & v
    End With
End Sub


