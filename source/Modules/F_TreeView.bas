Attribute VB_Name = "F_TreeView"
Rem @Folder Treeview Declarations
Public Enum tvImages
    tvProject = 1
    tvSheet = 2
    tvForm = 3
    tvModule = 4
    tvClass = 5
    tvMacro = 6
    tvText = 7
End Enum

Rem @Folder Treeview
Sub FindCode(Optional s As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE TreeviewExpandAllNodes
    '#INCLUDE TreeviewAssignProjectImages
    '#INCLUDE FindCodeEverywhere
    Load uCodeFinder
    If s = "" Then s = CodepaneSelection
    If Len(s) > 0 Then
        FindCodeEverywhere s, uCodeFinder.TreeView1
        TreeviewAssignProjectImages uCodeFinder.TreeView1
        TreeviewExpandAllNodes uCodeFinder.TreeView1
    End If
    uCodeFinder.TextBox1.TEXT = s
    uCodeFinder.Show
End Sub

Sub DebugPrintCodeLinesContaining(f)
    '#INCLUDE ProceduresOfModule
    '#INCLUDE dp
    '#INCLUDE ProtectedVBProject
    '#INCLUDE GetModuleText
    '#INCLUDE GetProjectText
    '#INCLUDE GetProcText
    Const ModuleString = vbNewLine & "    m|"
    Const Procedurestring = "" & vbTab & "p" & "|" & "---" & "| "
    Const FoundString = "" & vbTab & "s" & "|" & vbTab & " |" & "---" & "| "
    Dim X, Y, s, p As Variant
    Dim Module As VBComponent
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    If UBound(Filter(Split(GetProjectText(Workbooks(Y.Name)), vbNewLine), f, True, vbTextCompare)) > -1 Then
                        dp vbNewLine
                        dp "----------------------------------"
                        dp "| " & Y.Name
                        dp "----------------------------------"
                        For Each Module In Workbooks(Y.Name).VBProject.VBComponents
                            If UBound(Filter(Split(GetModuleText(Module), vbNewLine), f, True, vbTextCompare)) > -1 Then
                                dp ModuleString & Module.Name
                                For Each p In ProceduresOfModule(Module)
                                    If UBound(Filter(Split(GetProcText(Module, CStr(p)), vbNewLine), f, True, vbTextCompare)) > -1 Then
                                        dp Procedurestring & CStr(p)
                                        s = Filter(Split(GetProcText(Module, CStr(p)), vbNewLine), f, True, vbTextCompare)
                                        For i = 0 To UBound(s)
                                            dp FoundString & Trim(s(i))
                                        Next i
                                    End If
                                Next p
                            End If
                        Next Module
                    End If
                End If
            End If
            err.clear
        Next Y
    Next X
End Sub

Sub TreeviewExpandAllNodes(TV As TreeView)
    For i = 1 To TV.Nodes.count
        TV.Nodes(i).Expanded = True
    Next
End Sub

Sub TreeviewCollapseAllNodes(TV As TreeView)
    For i = 1 To TV.Nodes.count
        TV.Nodes(i).Expanded = False
    Next
End Sub

Sub TreeviewClear(TV As TreeView)
    For i = TV.Nodes.count To 1 Step -1
        TV.Nodes.Remove i
    Next
End Sub

Sub TreeviewCheckAllChildren(parent As MSComctlLib.node, _
                             Optional check As Boolean = True)
    Rem In userform:
    Rem Sub treeview1_NodeCheck(ByVal node As MSComctlLib.node)
    Rem     TreeviewNodeCheck node, node.Checked
    Rem End Sub
    Dim child As MSComctlLib.node
    parent.Checked = check
    Set child = parent.child
    While Not child Is Nothing
        TreeviewCheckAllChildren child, check
        Set child = child.Next
    Wend
End Sub

Public Function TreeviewGetLevel(theNode As node) As Integer
    TreeviewGetLevel = 1
    Do Until theNode.Root = theNode.FirstSibling
        TreeviewGetLevel = TreeviewGetLevel + 1
        Set theNode = theNode.parent
    Loop
End Function

Sub PopulateTreeviewFromSheetHierarchy( _
    TargetTreeView As TreeView, _
    StartRange As Range, _
    Optional ClearPreviousNodes As Boolean = True, _
    Optional Expanded As Boolean = False)
    Rem example use
    Rem    PopulateTreeviewFromSheetHierarchy me.Treeview1,thisworkbook.sheets("TreeviewHierarchy").range("A1"),true,false
    Rem example of sheet structure
    Rem  |1|2|3
    Rem 1|A| |
    Rem 2| |1|
    Rem 3| | |1.1
    Rem 4| | |1.2
    Rem 5|B| |
    Rem 6| |2|
    Rem 7| | |2.1
    Rem 8| | |2.2
    Dim nP As node
    Dim c As Excel.Range
    On Error Resume Next
    With TargetTreeView
        If ClearPreviousNodes = True Then .Nodes.clear
        For Each c In StartRange.parent.Columns(StartRange.Column).SpecialCells(xlCellTypeConstants)
            Set nP = .Nodes.Add(, , c.Address, c.Value)
        Next
        For Each c In StartRange.CurrentRegion
            If c.Value <> vbNullString And c.Address <> StartRange.Address And c.Column <> 1 Then
                Set nP = .Nodes(c.OFFSET(, -1).End(xlUp).Address)
                If nP Is Nothing Then
                    MsgBox "ERROR: Parent node " & c.OFFSET(, -1).End(xlUp).Value & " not found...", vbExclamation, "Error"
                    Exit Sub
                End If
                .Nodes.Add nP, tvwChild, c.Address, c.Value
                If err.Number <> 0 Then
                    MsgBox "ERROR: The node " & c.Value & " is a duplicate. All node descrptions must be unique", vbExclamation, "Error"
                    Exit Sub
                End If
                nP.Expanded = Expanded
            End If
        Next
        With .Nodes(Range(cROOT).Address)
            .SELECTED = True
            .EnsureVisible
        End With
    End With
    Exit Sub
End Sub

Sub TreeviewAllProjects(TV As TreeView)
    '#INCLUDE ProceduresOfModule
    '#INCLUDE getModuleName
    '#INCLUDE ProtectedVBProject
    Dim nP As node
    Dim nM As node
    Dim nS As node
    Dim X, Y, s, p As Variant
    Dim Module As VBComponent
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    Set nP = TV.Nodes.Add(, , , Y.Name)
                    For Each element In Array(vbext_ct_Document, vbext_ct_MSForm, vbext_ct_StdModule, vbext_ct_ClassModule)
                        For Each Module In Workbooks(Y.Name).VBProject.VBComponents
                            If Module.Type = element Then
                                Set nM = TV.Nodes.Add(nP, tvwChild, , getModuleName(Module))
                                For Each p In ProceduresOfModule(Module)
                                    Set nS = TV.Nodes.Add(nM, tvwChild, , CStr(p))
                                Next p
                            End If
                        Next Module
                    Next
                End If
            End If
            err.clear
        Next Y
    Next X
End Sub

Sub ImageListLoadProjectIcons(imgList As ImageList, TV As TreeView)
    strPath = "C:\Users\acer\Dropbox\SOFTWARE\EXCEL\0 Alex\treeviewicons\"
    With imgList.ListImages
        .Add , "Project", LoadPicture(strPath & "Project.jpg")
        .Add , "Sheet", LoadPicture(strPath & "Sheet.jpg")
        .Add , "Form", LoadPicture(strPath & "Form.jpg")
        .Add , "Module", LoadPicture(strPath & "Module.jpg")
        .Add , "Class", LoadPicture(strPath & "Class.jpg")
        .Add , "Macro", LoadPicture(strPath & "Macro.jpg")
        .Add , "Text", LoadPicture(strPath & "Text.jpg")
    End With
    TV.ImageList = imgList
End Sub

Sub TreeviewAssignProjectImages(TV As TreeView)
    '#INCLUDE ComponentTypeToString
    '#INCLUDE ModuleOfWorksheet
    '#INCLUDE TreeviewGetLevel
    Dim i As Long
    Dim Module As VBComponent
    For i = 1 To TV.Nodes.count
        Select Case TreeviewGetLevel(TV.Nodes.item(i))
            Case 1
                If InStr(1, TV.Nodes.item(i).TEXT, ".") = 0 Then GoTo Skip
                TV.Nodes.item(i).image = tvImages.tvProject
            Case 2
                Set TargetWorkbook = Workbooks(TV.Nodes.item(i).parent.TEXT)
                If InStr(1, TargetWorkbook.Name, ".") = 0 Then GoTo Skip
                ModuleName = TV.Nodes.item(i).TEXT
                Set Module = Nothing
                On Error Resume Next
                Set Module = TargetWorkbook.VBProject.VBComponents(ModuleName)
                On Error GoTo 0
                If Module Is Nothing Then
                    Set Module = ModuleOfWorksheet(TargetWorkbook.Worksheets(TV.Nodes.item(i).TEXT))
                End If
                Select Case ComponentTypeToString(Module.Type)
                    Case "Document Module"
                        TV.Nodes.item(i).image = tvImages.tvSheet
                    Case "UserForm"
                        TV.Nodes.item(i).image = tvImages.tvForm
                    Case "Code Module"
                        TV.Nodes.item(i).image = tvImages.tvModule
                    Case "Class Module"
                        TV.Nodes.item(i).image = tvImages.tvClass
                End Select
            Case 3
                TV.Nodes.item(i).image = tvImages.tvMacro
            Case 4
                TV.Nodes.item(i).image = tvImages.tvText
        End Select
Skip:
    Next i
End Sub

Sub TreeviewGotoProjectElement(TV As TreeView)
    '#INCLUDE GoToModule
    '#INCLUDE TreeviewGetLevel
    Dim Module As VBComponent
    Select Case TreeviewGetLevel(TV.SelectedItem)
        Case Is = 1
        Case Is = 2
            With TV.SelectedItem
                On Error Resume Next
                Set Module = Workbooks(.parent.TEXT).VBProject.VBComponents(.TEXT)
                On Error GoTo 0
                If Module Is Nothing Then Set Module = Workbooks(.parent.TEXT).VBProject.VBComponents(Workbooks(.parent.TEXT).SHEETS(.TEXT).CodeName)
                GoToModule Module
            End With
        Case Is = 3
            With TV.SelectedItem
                On Error Resume Next
                Set Module = Workbooks(.parent.parent.TEXT).VBProject.VBComponents(.parent.TEXT)
                On Error GoTo 0
                If Module Is Nothing Then Set Module = _
                   Workbooks(.parent.parent.TEXT).VBProject.VBComponents(Workbooks(.parent.parent.TEXT).SHEETS(.parent.TEXT).CodeName)
                GoToModule Module
                For i = 1 To Module.CodeModule.CountOfLines
                    If InStr(1, Module.CodeModule.Lines(i, 1), "Sub " & .TEXT) > 0 Or _
                                                                               InStr(1, Module.CodeModule.Lines(i, 1), "Function " & .TEXT) > 0 Then
                        Module.CodeModule.CodePane.SetSelection i, 1, i, 1
                        Exit Sub
                    End If
                Next
            End With
        Case Is = 4
            With TV.SelectedItem
                Set Module = Workbooks(.parent.parent.parent.TEXT).VBProject.VBComponents(.parent.parent.TEXT)
                GoToModule Module
                DoEvents
                For i = 1 To Module.CodeModule.CountOfLines
                    If Trim(Module.CodeModule.Lines(i, 1)) = .TEXT Then
                        Module.CodeModule.CodePane.SetSelection i, 1, i, 1
                        Exit Sub
                    End If
                Next
            End With
    End Select
End Sub

Sub FindCodeEverywhere(f As String, TV As TreeView)
    '#INCLUDE ProceduresOfModule
    '#INCLUDE getModuleName
    '#INCLUDE ProtectedVBProject
    '#INCLUDE GetModuleText
    '#INCLUDE GetProjectText
    '#INCLUDE GetProcText
    Dim nP As node
    Dim nM As node
    Dim nS As node
    Dim nF As node
    Dim X, Y, s, p As Variant
    Dim Module As VBComponent
    On Error Resume Next
    For Each X In Array(Workbooks, AddIns)
        For Each Y In X
            If Not ProtectedVBProject(Workbooks(Y.Name)) Then
                If err.Number = 0 Then
                    If UBound(Filter(Split(GetProjectText(Workbooks(Y.Name)), vbNewLine), f, True, vbTextCompare)) > -1 Then
                        Set nP = TV.Nodes.Add(, , , Y.Name)
                        For Each Module In Workbooks(Y.Name).VBProject.VBComponents
                            If UBound(Filter(Split(GetModuleText(Module), vbNewLine), f, True, vbTextCompare)) > -1 Then
                                Set nM = TV.Nodes.Add(nP, tvwChild, , getModuleName(Module))
                                For Each p In ProceduresOfModule(Module)
                                    If UBound(Filter(Split(GetProcText(Module, CStr(p)), vbNewLine), f, True, vbTextCompare)) > -1 Then
                                        Set nS = TV.Nodes.Add(nM, tvwChild, , CStr(p))
                                        s = Filter(Split(GetProcText(Module, CStr(p)), vbNewLine), f, True, vbTextCompare)
                                        For i = 0 To UBound(s)
                                            Set nF = TV.Nodes.Add(nS, tvwChild, , Trim(s(i)))
                                        Next i
                                    End If
                                Next p
                            End If
                        Next Module
                    End If
                End If
            End If
            err.clear
        Next Y
    Next X
End Sub

Sub TreeviewSelectNodes(TV As TreeView, SingleSelect As Boolean, lvl1crit As String, Optional CriteriaByLevel As Variant)
    Dim nd As node
    For Each nd In TV.Nodes
        If nd.TEXT = lvl1crit Then
            nd.SELECTED = True
            nd.Expanded = True
            If SingleSelect = True Then Exit For
        End If
    Next
    X = nd.index + 1
    Dim crit
    For Each crit In CriteriaByLevel
        For i = X To TV.Nodes.count
            If TV.Nodes.item(i).TEXT = crit Then
                TV.Nodes.item(i).SELECTED = True
                TV.Nodes.item(i).Expanded = True
                If SingleSelect = True Then Exit For
            End If
        Next
        X = i + 1
    Next
End Sub


