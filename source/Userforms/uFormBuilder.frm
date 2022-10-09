VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uFormBuilder 
   Caption         =   "Form Transformer"
   ClientHeight    =   8904.001
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11988
   OleObjectBlob   =   "uFormBuilder.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uFormBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uFormBuilder
'* Created    : 06-10-2022 10:35
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private WithEvents Emitter As EventListenerEmitter
Attribute Emitter.VB_VarHelpID = -1

Const codecolumn = 4
Const nameRow = 2
Const firstControlColumn = 7
Dim newType As String

Sub Emitter_MouseOver(ByRef control As Object)
    memo.Caption = control.Tag
End Sub

Private Sub Label9_Click()
    uDEV.Show
End Sub

Private Sub UserForm_Initialize()

    Set Emitter = New EventListenerEmitter
    Emitter.AddEventListenerAll Me
    
    Set UFSheet = ThisWorkbook.SHEETS("FormBuilder")
    PopulateUserforms
    ResizeControlColumns Me.ListUserforms
    Me.ListNewType.list = WorksheetFunction.Transpose(ThisWorkbook.SHEETS("FormBuilderEvents").Range("A1").CurrentRegion.RESIZE(1).Value)
    'If ListUserforms.ListCount > 0 Then ListUserforms.Selected(0) = True
End Sub

Private Sub placeholder_Click()
    
    Dim TargetWorkbook As Workbook
    Set TargetWorkbook = Workbooks(ListUserforms.list(ListUserforms.ListIndex, 1))
    If ListItems.ListIndex = -1 Then
        MsgBox "No item selected"
        Exit Sub
    End If
    If ListControls.list(ListControls.ListIndex) <> ListNewType.list(ListNewType.ListIndex) Then
        MsgBox "You are trying to add " & _
               left(ListNewAction.list(ListNewAction.ListIndex), InStr(1, ListNewAction.list(ListNewAction.ListIndex), "(") - 1) & _
               " from a " & ListNewType.list(ListNewType.ListIndex) & " to a " & ListControls.list(ListControls.ListIndex) & vbNewLine & _
               "Since new action may be unavailable for old control type," & vbNewLine & _
               "please <TRANSFORM> the old control type to the new type first"
        Exit Sub
    End If
    For i = 0 To ListItems.ListCount - 1
        If ListItems.SELECTED(i) = True Then
            addControlCode TargetWorkbook, _
                           ListUserforms.list(ListUserforms.ListIndex), _
                           ListItems.list(i), _
                           ListNewAction.list(ListNewAction.ListIndex), _
                           ""
            'addControlCode TargetWorkbook, _
            ListUserforms.List(ListUserforms.ListIndex), _
            ListItems.List(ListItems.ListIndex), _
            ListNewAction.List(ListNewAction.ListIndex), _
            ""
        End If
    Next

    UFSheet.Columns.Cells.WrapText = False
    
    If ListboxSelectedCount(ListItems) = 1 Then
        ListItems_Change
    Else
        Application.EnableEvents = False
        For i = 0 To ListItems.ListCount - 1
            ListItems.SELECTED(i) = i = 0
        Next
        Application.EnableEvents = True
    End If
    
    'ListOldAction.AddItem ListNewAction.List(ListNewAction.ListIndex)

End Sub

Private Sub RECREATE_Click()
    CreateUserForm Workbooks(ListUserforms.list(ListUserforms.ListIndex, 1))
End Sub

Private Sub transform_Click()
    If ListboxSelectedCount(ListItems) > 1 Then
        MsgBox "Can transform only one control at a time"
        Exit Sub
    End If
    If noemptyListindex = False Then
        MsgBox "Some listboxes have no selection"
        Exit Sub
    End If
    transformControl ListItems.list(ListItems.ListIndex)
End Sub

Sub transformControl(CtrName As String)
    Dim i As Long

    newType = ListNewType.list(ListNewType.ListIndex)
    
    'change control type
    Dim rng As Range
    Set rng = UFSheet.Range("G1", UFSheet.Cells(1, Columns.count).End(xlToLeft))
    Set rng = rng.OFFSET(nameRow - 1)
    Dim ctrCol As Long
    ctrCol = rng.Find(CtrName).Column
    UFSheet.Cells(1, ctrCol) = newType
    
    'set new action
    Dim newAction As String
    newAction = ListNewAction.list(ListNewAction.ListIndex)
    
    'set oldAction and newAction
    Dim oldAction As String
    oldAction = ListOldAction.list(ListOldAction.ListIndex)
    
    'replace code with new control action
    On Error Resume Next
    Set cell = UFSheet.Columns(codecolumn).Find(CtrName & oldAction)
    On Error GoTo 0
    cell.Value = Replace(cell.Value, oldAction, newAction)
    ListOldAction.list(ListOldAction.ListIndex) = newAction
    Reselect
    
End Sub

Sub Reselect()
    Dim i As Long
    'save userform and control selections
    Dim pickedUserform As Long
    pickedUserform = ListUserforms.ListIndex
    Dim pickedControl As String
    pickedControl = ListNewType.list(ListNewType.ListIndex)
    Dim pickedItem As String
    pickedItem = ListItems.list(ListItems.ListIndex)
    DeselectAll
    clearAll
    
    'reselect userform
    For i = 0 To ListUserforms.ListCount - 1
        If i = pickedUserform Then ListUserforms.SELECTED(i) = True
    Next i
    
    'reselect control type if more of same type exist
    For i = 0 To ListControls.ListCount - 1
        If ListControls.list(i) = pickedControl Then
            ListControls.SELECTED(i) = True
        End If
    Next i
    ' else select the new type
    'If ListControls.ListIndex = -1 Then
    For i = 0 To ListControls.ListCount - 1
        If ListControls.list(i) = newType Then
            ListControls.SELECTED(i) = True
        End If
    Next i
    'End If
    'ListItems.Selected(0) = True
    For i = 0 To ListItems.ListCount - 1
        If ListItems.list(i) = pickedItem Then
            ListItems.SELECTED(i) = True
        End If
    Next i
    
    If ListOldAction.ListCount > 0 Then ListOldAction.SELECTED(0) = True
End Sub

Sub DeselectAll()
    Dim ctr As control
    Dim LB As MSForms.ListBox
    For Each ctr In Me.Controls
        'Debug.Print ctr.Name
        If TypeName(ctr) = "ListBox" Then
            '            If ctr.Name <> "ListUserforms" Then
            Set LB = ctr
            If LB.ListCount <> 0 Then
                For i = 0 To LB.ListCount - 1
                    LB.SELECTED(i) = False
                Next i
            End If
            '            End If
        End If
    Next ctr
End Sub

Sub clearAll()
    Dim ctr As control
    Dim LB As MSForms.ListBox
    For Each ctr In Me.Controls
        'Debug.Print ctr.Name
        If TypeName(ctr) = "ListBox" Then
            Set LB = ctr
            If LB.Name <> "ListUserforms" And LB.Name <> "ListNewType" Then
                LB.clear
            End If
        End If
    Next ctr
End Sub

Private Sub deselect_Click()
    DeselectAll
    clearAll
End Sub

Private Sub export_Click()
    ExportSelectedForm
End Sub

Sub ExportSelectedForm()
    Dim pickedUserform As String
    If ListUserforms.ListIndex = -1 Then Exit Sub
    pickedUserform = ListUserforms.list(ListUserforms.ListIndex)
    ExportUserform Workbooks(ListUserforms.list(ListUserforms.ListIndex, 1)).VBProject.VBComponents(pickedUserform)

    clearAll
    DeselectAll
    'Reselect UserForm
    For i = 0 To ListUserforms.ListCount - 1
        If ListUserforms.list(i) = pickedUserform Then
            ListUserforms.SELECTED(i) = True
            Exit For
        End If
    Next i
    If ListControls.ListCount > 0 Then ListControls.SELECTED(0) = True
End Sub

Private Sub ListControls_Click()
    Me.ListOldAction.clear
    PopulateItems
    If ListboxSelectedCount(ListItems) > 0 Then _
                                       ListItems.SELECTED(0) = True
    If ListOldAction.ListCount > 0 Then ListOldAction.SELECTED(0) = True
    Dim i As Long
    For i = 0 To ListNewType.ListCount - 1
        If ListNewType.list(i) = ListControls.list(ListControls.ListIndex) Then
            ListNewType.SELECTED(i) = True
        End If
    Next
End Sub

Private Sub ListItems_Change()
    'populate old actions listbox with selected control's events that exist in userform code
    
    ListOldAction.clear
    
    If ListboxSelectedCount(ListItems) > 1 Then Exit Sub
    Dim cell As Range

    Dim nameStr As String
    nameStr = ListItems.list(ListItems.ListIndex)
    Dim actionStr As String
    Dim openPos As Integer
    Dim closePos As Integer
    Dim cellAddress As String
    On Error Resume Next
    Set cell = UFSheet.Columns(codecolumn).Find(nameStr & "_")        ' & oldCtrAction)
    Dim firstAddress As String
    firstAddress = cell.Address
    On Error GoTo 0
    Do While Not cell Is Nothing And cellAddress <> firstAddress
        openPos = InStr(1, cell, "_")        'InStr(1, cell, nameStr) '
        closePos = InStr(1, cell, ")")        'IIf(InStr(1, cell, ")") > 0, InStr(1, cell, ")"), 0)
        ListOldAction.AddItem Mid(cell, openPos, Abs(closePos - openPos) + 1)        ' cell,openpos +1 ...cell, ")")-1, 0)
        Set cell = UFSheet.Columns(codecolumn).FindNext(cell)        '(nameStr & "_") ' & oldCtrAction)
        cellAddress = cell.Address
    Loop
    If ListOldAction.ListCount > 0 Then ListOldAction.ListIndex = 0
End Sub

Private Sub ListNewType_Click()
    'populate new actions listbox with new type's available events
    Me.ListNewAction.clear
    Dim eventSHEET As Worksheet
    Set eventSHEET = ThisWorkbook.SHEETS("FormBuilderEvents")
    Dim rng As Range
    Set rng = eventSHEET.Range("A1").CurrentRegion.RESIZE(1)
    Dim cell As Range
    Dim col As Long
    col = rng.Find(ListNewType.list(ListNewType.ListIndex)).Column
    Dim row As Long
    row = 2
    Set rng = eventSHEET.Cells(row, col)
    Do While Not IsEmpty(eventSHEET.Cells(row, col))
        Set rng = Union(rng, eventSHEET.Cells(row, col))
        row = row + 1
    Loop
    ListNewAction.list = rng.Value
    '    For Each cell In rng
    '        Me.ListNewAction.AddItem cell
    '    Next
    '
    'if new type has same event as old type then select that
    If Me.ListOldAction.ListIndex <> -1 Then
        If Me.ListNewAction.ListCount = 0 Then Exit Sub
        For i = 0 To Me.ListNewAction.ListCount - 1
            If Me.ListNewAction.list(i) = Me.ListOldAction.list(Me.ListOldAction.ListIndex) Then
                Me.ListNewAction.SELECTED(i) = True
            Else
                Me.ListNewAction.SELECTED(i) = False
            End If
        Next i
    End If
    ListNewAction.ListIndex = 0
End Sub

Private Sub ListUserforms_Click()
    clearAll
    'ExportSelectedForm
    If ListUserforms.list(ListUserforms.ListIndex) = UFSheet.Range("B1").Value Then
        PopulateControls
    Else
        Me.ListControls.clear
    End If
    If ListControls.ListCount > 0 Then ListControls.SELECTED(0) = True
End Sub

Sub PopulateUserforms(Optional TargetWorkbook As Workbook)        '
    Dim Module As VBComponent
    Dim X, WorkbookOrAddin As Variant
    
    If TargetWorkbook Is Nothing Then
        On Error Resume Next
        For Each X In Array(Workbooks, AddIns)
            For Each WorkbookOrAddin In X
                If Not ProtectedVBProject(Workbooks(WorkbookOrAddin.Name)) Then
                    If err.Number = 0 Then
                        For Each Module In Workbooks(WorkbookOrAddin.Name).VBProject.VBComponents
                            If Module.Type = vbext_ct_MSForm Then
                                If Module.Name <> "TransFormer" Then
                                    Me.ListUserforms.AddItem
                                    ListUserforms.list(UBound(ListUserforms.list), 0) = Module.Name
                                    ListUserforms.list(UBound(ListUserforms.list), 1) = WorkbookOfModule(Module).Name
                                End If
                            End If
                        Next Module
                    End If
                End If
                err.clear
            Next
        Next
    
    Else
        For Each Module In TargetWorkbook.VBProject.VBComponents
            If Module.Type = vbext_ct_MSForm Then
                If Module.Name <> "TransFormer" Then
                    Me.ListUserforms.AddItem
                    ListUserforms.list(UBound(ListUserforms.list), 0) = Module.Name
                    ListUserforms.list(UBound(ListUserforms.list), 1) = WorkbookOfModule(Module).Name
                End If
            End If
        Next Module
    End If
End Sub

Sub PopulateControls()

    Me.ListControls.clear

    Dim cell As Range
    Dim rng As Range
    Set rng = UFSheet.Range("G1")
    Dim col As Long
    col = firstControlColumn
    
    'get range with controls
    Do While Not IsEmpty(UFSheet.Cells(1, col))
        Set rng = Union(rng, UFSheet.Cells(1, col))
        col = col + 1
    Loop
    
    'use collection to get unique list because it skips existing keys
    Dim coll As Collection
    Set coll = New Collection

    On Error Resume Next
    For Each cell In rng
        coll.Add cell.TEXT, cell.TEXT
    Next
    On Error GoTo 0
    
    'populate controls listbox
    Dim element As Variant
    For Each element In coll
        Me.ListControls.AddItem element
    Next element
End Sub

Sub PopulateItems()
    Me.ListItems.clear
    Dim col As Long
    col = firstControlColumn

    Do While Not IsEmpty(UFSheet.Cells(1, col))
        If UFSheet.Cells(1, col) = Me.ListControls.list(ListControls.ListIndex) Then
            Me.ListItems.AddItem UFSheet.Cells(nameRow, col)
        End If
        col = col + 1
    Loop
End Sub

Function noemptyListindex() As Boolean
    Dim lbcount As Long
    Dim counter As Long
    Dim ctr As control
    Dim LB As MSForms.ListBox
    For Each ctr In Me.Controls
        'Debug.Print ctr.Name
        If TypeName(ctr) = "ListBox" Then
            lbcount = lbcount + 1
            Set LB = ctr
            If LB.ListIndex <> -1 Then
                counter = counter + 1
            End If
        End If
    Next ctr
    If counter = lbcount Then noemptyListindex = True
End Function

'obsolete
'
'Sub ExportTargetUserform(ufName As String)
'    Dim uForm As Variant
'    Set uForm = UserForms.Add(ufName)
'    ExportUserform uForm
'
'End Sub


'''''''''''''''''''''




Function propertyExists(cell As Range) As Boolean
    Dim lookWhere As Range:     Set lookWhere = ThisWorkbook.SHEETS("FromBuilderProperites").rows(1)
    Dim cellHeader As Range:    Set cellHeader = ThisWorkbook.SHEETS("FormBuilder").Cells(1, cell.Column)
    Dim cellProperty As Range:  Set cellProperty = ThisWorkbook.SHEETS("FormBuilder").Cells(cell.row, "F")
    Dim rng As Range:                 Set rng = ThisWorkbook.SHEETS("FormBuilderProperites").Columns(lookWhere.Find(cellHeader).Column).Cells.SpecialCells(xlCellTypeConstants)
    Dim element As Range
    On Error Resume Next
    Set element = rng.Find(cellProperty)
    On Error GoTo 0
    propertyExists = Not element Is Nothing
End Function

Sub addControlCode(wb As Workbook, FormName As String, CtrName As String, action As String, Code As String)
    Dim uf As UserForm
    Set uf = UserForms.Add(FormName)
    Dim ufComp As VBComponent
    Set ufComp = wb.VBProject.VBComponents(FormName)
    Dim ctl As control
    Dim txt As String
    Dim str As String
    'str = "Rem vbArc"
    'If ufComp.CodeModule.CountOfLines = 0 Then ufComp.CodeModule.AddFromString (str)
    If ufComp.CodeModule.CountOfLines > 0 Then
        str = ufComp.CodeModule.Lines(1, ufComp.CodeModule.CountOfLines)
    End If
    If InStr(1, str, CtrName & action) > 0 Then Exit Sub
    txt = "Sub " & CtrName & action & _
          vbNewLine & _
          Code & _
          vbNewLine & _
          "End Sub"
    Dim var As Variant
             
    var = Split(txt, vbNewLine)
    UFSheet.Cells(Columns.count, 4).End(xlUp).OFFSET(1).RESIZE(Len(txt) - Len(Replace(txt, vbNewLine, "")) - 1).Value = WorksheetFunction.Transpose(var)

    'ufComp.CodeModule.InsertLines ufComp.CodeModule.CountOfLines + 1, txt

End Sub


