Attribute VB_Name = "U_FormNavigator"
Rem @Folder UserformNavigator
Sub FormopenCode(FormName As String, TargetWorkbook As Workbook)
    '#INCLUDE AddModule
    Dim Module As VBComponent
    Set Module = AddModule("vbArc", vbext_ct_StdModule, TargetWorkbook)
    Dim addText As String
    addText = "Sub open" & FormName & vbNewLine
    addText = addText & "On error resume next" & vbNewLine
    addText = addText & FormName & ".show" & vbNewLine
    addText = addText & "On error goto 0" & vbNewLine
    addText = addText & "End Sub"
    Module.CodeModule.InsertLines Module.CodeModule.CountOfLines + 1, addText
End Sub

Function AddModule(compName As String, compType As VBIDE.vbext_ComponentType, Optional TargetWorkbook As Workbook) As VBComponent
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
    Set AddModule = vbComp
End Function

