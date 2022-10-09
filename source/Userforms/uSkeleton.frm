VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSkeleton 
   Caption         =   "github.com/alexofrhodes"
   ClientHeight    =   9600.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   18960
   OleObjectBlob   =   "uSkeleton.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uSkeleton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uSkeleton
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim wb As Workbook
Dim vbProj As VBProject
Dim vbComp As VBComponent
Dim comp As String
Dim proc

Private Sub exportDeclarationsAndCalls_Click()
    setUp
    Dim var As Variant
    Dim collections As New Collection
    Set collections = FindCalls(wb)
    If collections(1).count > 0 Then
        var = CollectionsToArrayTable(collections)
        ArrayToRange2D var, dataToSheet(wb, "exportCalls", "A2")
        With wb.SHEETS("exportCalls").Range("A1:B1")
            .Value = Array("Procedure", "Calls")
            .Font.Bold = True
            .Font.Size = 14
        End With
        With wb.SHEETS("exportCalls").Cells
            .WrapText = False
            .Columns.AutoFit
            .WrapText = True
            .Columns.AutoFit
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
    Set collections = getDeclarations(wb, True, True, True, True, True, True)
    If collections(1).count > 0 Then
        var = CollectionsToArrayTable(collections)
        ArrayToRange2D var, dataToSheet(wb, "exportDeclarations", "A2")
        With wb.SHEETS("exportDeclarations").Range("A1:F1")
            .Value = Array("Component Type", "Component Name", "Declaration Scope", "Declaration Type", "Declaration Keyword", "Declaration Code")
            .Font.Bold = True
            .Font.Size = 14
        End With
        With wb.SHEETS("exportDeclarations").Cells
            .WrapText = False
            .Columns.AutoFit
            .WrapText = True
            .Columns.AutoFit
            .VerticalAlignment = xlVAlignCenter
        End With
    End If
End Sub

Function GetCallsOfProcedureSkeleton(wb As Workbook, vbComp As VBComponent, procName As String) As Collection
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE GetProcText
    Dim coll As Collection: Set coll = New Collection
    Dim WorkbookProcedure As Variant
    Dim AllProcs As Collection: Set AllProcs = ProceduresOfWorkbook(wb)
    Dim procText As String:    procText = GetProcText(vbComp, procName)
    For Each WorkbookProcedure In AllProcs
        If CStr(WorkbookProcedure) <> procName Then
            If InStr(1, procText, CStr(WorkbookProcedure)) Then
                coll.Add CStr(WorkbookProcedure)
            End If
        End If
    Next WorkbookProcedure
    Set GetCallsOfProcedureSkeleton = coll
End Function

Private Sub Image1_click()
    uDEV.Show
End Sub

Private Sub UserForm_Initialize()
    loadProjects
End Sub

Sub loadProjects()
    '#INCLUDE ProtectedVBProject
    For Each wb In Workbooks
        If Not ProtectedVBProject(wb) Then LProjects.AddItem wb.Name
    Next
    On Error Resume Next
    For Each ad In AddIns
        If Not ProtectedVBProject(Workbooks(ad.Name)) Then
            If err = 0 Then LProjects.AddItem ad.Name
            err.clear
        End If
    Next
End Sub

Private Sub LProjects_Click()
    loadComponents
End Sub

Sub loadComponents()
    '#INCLUDE ComponentTypeToString
    '#INCLUDE SortListboxOnColumn
    '#INCLUDE setUp
    '#INCLUDE ReleaseMe
    '#INCLUDE ControlsResizeColumns
    LComponents.clear: LProcedures.clear: TPROCS.TEXT = "": LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    For Each vbComp In vbProj.VBComponents
        LComponents.AddItem
        LComponents.list(LComponents.ListCount - 1, 0) = ComponentTypeToString(vbComp.Type)
        LComponents.list(LComponents.ListCount - 1, 1) = vbComp.Name
    Next
    ReleaseMe
    SortListboxOnColumn LComponents, 0
    ControlsResizeColumns LComponents
End Sub

Private Sub LComponents_Click()
    LProcedures.clear: TPROCS.TEXT = "": LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    For Each proc In ProcList(vbComp)
        LProcedures.AddItem proc
    Next proc
    TComps.TEXT = GetModuleText(vbComp)
    SortListboxOnColumn LProcedures, 0
End Sub

Private Sub LProcedures_Click()
    LCalls.clear: TCalls.TEXT = "": LDeclarations.clear: TDeclarations.TEXT = "":
    setUp
    TPROCS.TEXT = GetProcText(vbComp, CStr(proc))
    Do While InStr(1, TPROCS.TEXT, "  ") > 0
        TPROCS.TEXT = Replace(TPROCS.TEXT, "  ", " ")
    Loop
    Dim element
    For Each element In GetCallsOfProcedureSkeleton(wb, ModuleOfProcedure(wb, CStr(proc)), CStr(proc))
        LCalls.AddItem element
    Next
    SortListboxOnColumn LCalls, 1
    'test
    Dim coll As Collection:     Set coll = getDeclarations(wb, True, True, True, True, True, True)
    Dim keyCol As Collection:   Set keyCol = coll.item(5)
    Dim decCol As Collection:   Set decCol = coll.item(6)
    Dim i As Long
    Dim tmp As String
    For i = 1 To keyCol.count
        'if the DECLARATION keyword exists inside the procedure
        If InStr(1, TPROCS.TEXT, keyCol.item(i)) > 0 Then
            'and if it is not a VARIABLE inside the procedure
            If InStr(1, TPROCS.TEXT, keyCol.item(i) & " As") = 0 Then
                'avoid duplicates
                If ListboxContains(LDeclarations, keyCol.item(i)) = False Then
                    LDeclarations.AddItem keyCol.item(i)
                End If
            End If
        End If
    Next i
    SortListboxOnColumn LDeclarations, 0
    SortListboxOnColumn LCalls, 0
    ReleaseMe
End Sub

Private Sub LCalls_Click()
    setUp
    proc = LCalls.list(LCalls.ListIndex)
    Set vbComp = ModuleOfProcedure(wb, CStr(proc))
    TCalls.TEXT = GetProcText(vbComp, CStr(proc))
    ReleaseMe
End Sub

Private Sub LDeclarations_Click()
    setUp
    Dim coll As Collection:     Set coll = getDeclarations(wb, True, True, True, True, True, True)
    Dim keyCol As Collection:   Set keyCol = coll.item(5)
    Dim decCol As Collection:   Set decCol = coll.item(6)
    Dim i As Long
    For i = 1 To keyCol.count
        If keyCol.item(i) = LDeclarations.list(LDeclarations.ListIndex) Then
            TDeclarations.TEXT = decCol.item(i)
        End If
    Next i
End Sub

Sub setUp()
    On Error Resume Next
    Set wb = Workbooks(LProjects.list(LProjects.ListIndex))
    Set vbProj = wb.VBProject
    comp = LComponents.list(LComponents.ListIndex, 1)
    Set vbComp = vbProj.VBComponents(comp)
    proc = LProcedures.list(LProcedures.ListIndex)
End Sub

Sub ReleaseMe()
    Set vbProj = Nothing
    Set vbComp = Nothing
    comp = ""
    Set wb = Nothing
End Sub

Function FindCalls(wb As Workbook) As Collection
    '#INCLUDE ProceduresOfWorkbook
    '#INCLUDE ModuleOfProcedure
    '#INCLUDE CollectionToArray
    '#INCLUDE GetCallsOfProcedureSkeleton
    Dim Procedure As Variant
    Dim output As New Collection
    Dim procedures As New Collection
    Dim calls As New Collection
    Dim element As Variant
    Dim tmp As New Collection
    For Each Procedure In ProceduresOfWorkbook(wb)
        Set tmp = GetCallsOfProcedureSkeleton(wb, ModuleOfProcedure(wb, CStr(Procedure)), CStr(Procedure))
        If tmp.count > 0 Then
            procedures.Add Procedure
            calls.Add Join(CollectionToArray(tmp), vbNewLine)
        End If
    Next
    output.Add procedures
    output.Add calls
    Set FindCalls = output
End Function

Function dataToSheet(Optional wb As Workbook, Optional wsName As String, Optional rngAddress As String, Optional confirmClear As Boolean) As Range
    '#INCLUDE answer
    '#INCLUDE sheetExists
    If wb Is Nothing Then Set wb = Workbooks.Add
    Dim ws As Worksheet
    If sheetExists(wsName, wb) Then
        If confirmClear = True Then
            Dim answer As Integer
            answer = MsgBox("Sheet " & wsName & " already exists. Cells will be cleared. Proceed?", vbYesNo)
            If answer = vbNo Then Exit Function
        End If
        Set ws = wb.SHEETS(wsName)
        ws.Cells.clear
    Else
        If wsName = "" Then
            Set ws = wb.SHEETS(1)
        Else
            Set ws = wb.SHEETS.Add
            ws.Name = wsName
        End If
    End If
    If rngAddress <> "" Then
        Set dataToSheet = ws.Range(rngAddress)
    Else
        Set dataToSheet = ws.Range("A1")
    End If
End Function

Function ProcList(vbComp As VBComponent) As Collection
    Dim CodeMod As CodeModule
    Set CodeMod = vbComp.CodeModule
    Dim coll As Collection
    Set coll = New Collection
    Dim LineNum As Long
    Dim NumLines As Long
    Dim procName As String
    Dim ProcKind As VBIDE.vbext_ProcKind
    LineNum = CodeMod.CountOfDeclarationLines + 1
    Do Until LineNum >= CodeMod.CountOfLines
        procName = CodeMod.ProcOfLine(LineNum, ProcKind)
        coll.Add procName
        LineNum = CodeMod.ProcStartLine(procName, ProcKind) + CodeMod.ProcCountLines(procName, ProcKind) + 1
    Loop
    Set ProcList = coll
End Function

Function ControlsResizeColumns(lBox As MSForms.control, Optional ResizeListbox As Boolean)
    '#INCLUDE sheetExists
    If lBox.ListCount = 0 Then Exit Function
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    If sheetExists("ListboxColumnWidth", ThisWorkbook) = False Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = "ListboxColumnwidth"
    Else
        Set ws = ThisWorkbook.Worksheets("ListboxColumnwidth")
        ws.Cells.clear
    End If
    ws.Cells.Font.Size = 12
    ws.Cells.Font.Name = "Calibri"
    '---Listbox to range-----
    Dim rng As Range
    Set rng = ThisWorkbook.SHEETS("ListboxColumnwidth").Range("A1")
    Set rng = rng.RESIZE(UBound(lBox.list) + 1, lBox.columnCount)
    rng = lBox.list
    '---Get ColumnWidths------
    rng.Columns.AutoFit
    Dim sWidth As String
    Dim vR() As Variant
    Dim n As Integer
    Dim cell As Range
    For Each cell In rng.RESIZE(1)
        n = n + 1
        ReDim Preserve vR(1 To n)
        vR(n) = cell.EntireColumn.Width
    Next cell
    sWidth = Join(vR, ";")
    'Debug.Print sWidth
    '---assign ColumnWidths----
    With lBox
        .ColumnWidths = sWidth
        '.RowSource = "A1:A3"
        .BorderStyle = fmBorderStyleSingle
    End With
    'remove worksheet
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    '----Resize Listbox--------
    If ResizeListbox = False Then Exit Function
    Dim w As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    lBox.Width = w + 10
End Function

Function sheetExists(sheetToFind As String, Optional InWorkbook As Workbook) As Boolean
    If InWorkbook Is Nothing Then Set InWorkbook = ThisWorkbook
    On Error Resume Next
    sheetExists = Not InWorkbook.SHEETS(sheetToFind) Is Nothing
End Function

