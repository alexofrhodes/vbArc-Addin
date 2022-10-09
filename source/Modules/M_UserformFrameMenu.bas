Attribute VB_Name = "M_UserformFrameMenu"
Rem @Folder UserformFrameMenu
Rem -----------------------------------
Rem Put in userform:
Rem -----------------------------------
Rem Private WithEvents Emitter As EventListenerEmitter
Rem
Rem Sub startFrameForm(FORM As Object)
Rem     Dim anc As MSForms.Control
Rem
Rem     For Each c In FORM.Controls
Rem         If TypeName(c) = "Frame" Then
Rem             rem c.Caption = ""
Rem             If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
Rem                 c.Visible = False
Rem                 If InStr(1, c.Tag, "anchor") > 0 Then
Rem                     On Error Resume Next
Rem                     Set anc = Me.Controls(c.Tag)
Rem                     If anc Is Nothing Then Stop
Rem                     On Error GoTo 0
Rem                     c.top = anc.top rem Anchor01.Top
Rem                     c.left = anc.left rem  Anchor01.Left
Rem                     Set anc = Nothing
Rem                 End If
Rem             End If
Rem         End If
Rem     Next
Rem     Set Emitter = New EventListenerEmitter
Rem     Emitter.AddEventListenerAll Me
Rem End Sub
Rem
Rem Private Sub Emitter_LabelMouseOut(Label As MSForms.Label)
Rem     If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
Rem         If Label.BackColor <> &H80B91E Then Label.BackColor = &H534848
Rem     End If
Rem End Sub
Rem
Rem Private Sub Emitter_LabelMouseOver(Label As MSForms.Label)
Rem     If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then
Rem         If Label.BackColor <> &H80B91E Then Label.BackColor = &H808080
Rem     End If
Rem End Sub
Rem
Rem Sub Emitter_LabelClick(ByRef Label As MSForms.Label)
Rem     If InStr(1, Label.Tag, "reframe", vbTextCompare) > 0 Then Reframe Me, Label
Rem End Sub
Rem

Rem Private Sub UserForm_Initialize()
Rem     startFrameForm Me
Rem End Sub
Sub addFrameFormCode(Module As VBComponent)
    '#INCLUDE CopyTemplateFromSheet
    '#INCLUDE dp
    '#INCLUDE Reframe
    '#INCLUDE GetModuleText
    '#INCLUDE CLIP
    If Module.Type <> vbext_ct_MSForm Then
        MsgBox "This is intended for a userform"
        Exit Sub
    End If
    Dim s As String
    s = CopyTemplateFromSheet("FrameForm")
    If InStr(1, GetModuleText(Module), Module.Name & "_Initialize") Then
        MsgBox "Threre is already _Initialize_ code in this form. Code will be put in cilpboard and immediate window."
        dp s
        CLIP s
    Else
        Module.CodeModule.AddFromString s
    End If
End Sub

Sub CreateFrameMenu(Optional Module As Object)
    '#INCLUDE SelectedControl
    '#INCLUDE SelectedControls
    '#INCLUDE ActiveModule
    '#INCLUDE addFrameSidebar
    If ActiveModule.Type <> vbext_ct_MSForm Then Exit Sub
    If SelectedControls.count = 0 Then
        ActiveModule.Designer.BackColor = 4208182
        addFrameSidebar ActiveModule
    Else
        addFrameSidebar SelectedControl
    End If
End Sub

Sub addFrameSidebar(form As Object, Optional dockRight As Boolean)
    '#INCLUDE askFormMenuElements
    '#INCLUDE UnderlineFrameName
    '#INCLUDE CreateOrSetFrame
    Dim f As MSForms.control
    Dim l As MSForms.control
    Set f = CreateOrSetFrame(form, "SideBar" & form.Name)
    f.Tag = "skip"
    f.BackColor = 5457992
    f.ForeColor = vbWhite
    f.BorderStyle = 1
    f.BorderStyle = 0
    f.Width = 80
    If TypeName(form) = "VBComponent" Then
        f.Height = 800
    Else
        f.Height = form.Height
    End If
    dockRight = IIf(TypeName(form) = "VBComponent", False, True)
    If dockRight = True Then
        f.left = form.Width - f.Width
    Else
        f.left = 0
    End If
    UnderlineFrameName form, f
    If TypeName(form) = "VBComponent" Then
        Set l = form.Designer.Controls.Add(ControlIDLabel, "Anchor" & form.Name)
    Else
        Set l = form.Controls.Add(ControlIDLabel, "Anchor" & form.Name)
    End If
    l.visible = False
    l.left = IIf(TypeName(form) = "VBComponent", f.left + f.Width + 9, 1)
    l.top = 12
    l.Width = 1
    l.BackColor = vbWhite
    l.visible = False
    askFormMenuElements form
End Sub

Sub askFormMenuElements(form As Object)
    '#INCLUDE InputboxString
    '#INCLUDE addFrameMenu
    Dim FormElements As String
    FormElements = InputboxString("Form Menus", "Type comma delimited menu names")
    If FormElements = "" Then Exit Sub
    Dim var
    var = Split(FormElements, ",")
    Dim i As Long
    For i = LBound(var) To UBound(var)
        var(i) = Trim(var(i))
    Next
    Dim coll As New Collection
    Dim element
    On Error Resume Next
    For Each element In var
        If Not IsNumeric(left(element, 1)) _
        And InStr(1, element, " ") = 0 Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    On Error GoTo 0
    For Each element In coll
        addFrameMenu form, CStr(element)
    Next
End Sub

Sub addFrameMenu(form As Object, FrameCaptionNoSpace As String)
    '#INCLUDE Reframe
    '#INCLUDE UnderlineFrameName
    '#INCLUDE CreateOrSetFrame
    '#INCLUDE AvailableFormOrFrameRow
    '#INCLUDE AvailableFormOrFrameColumn
    Dim f As MSForms.control
    Dim l As MSForms.control
    Dim Module As VBComponent
    If TypeName(form) = "VBComponent" Then
        Set Module = form
        Set f = Module.Designer.Controls.Add(ControlIDFrame, FrameCaptionNoSpace)
    Else
        Set Module = ThisWorkbook.VBProject.VBComponents(form.parent.Name)
        Set f = CreateOrSetFrame(Module.Designer.Controls(form.Name), FrameCaptionNoSpace)
    End If
    f.Tag = "anchor" & form.Name
    f.Caption = FrameCaptionNoSpace
    f.ForeColor = vbWhite
    f.visible = False
    If TypeName(form) = "VBComponent" Then
        f.left = AvailableFormOrFrameColumn(form.Designer)
    Else
        f.left = 0
    End If
    f.visible = True
    f.BorderStyle = 1
    f.BorderStyle = 0
    f.top = 12
    f.Width = 100
    UnderlineFrameName form, f
    If TypeName(form) = "VBComponent" Then
        Set l = Module.Designer.Controls("SideBar" & form.Name).Controls.Add(ControlIDLabel)
    Else
        Set l = Module.Designer.Controls("SideBar" & form.Name).Add(ControlIDLabel)
    End If
    l.Caption = FrameCaptionNoSpace
    l.ForeColor = vbWhite
    l.visible = False
    l.top = AvailableFormOrFrameRow(Module.Designer.Controls("SideBar" & form.Name))
    l.left = l.left + 3
    l.visible = True
    l.Tag = "reframe"
    l.Width = f.Width
End Sub

Sub AddControlsToFrame(isSubFrame As Boolean)
    '#INCLUDE SelectedControl
    '#INCLUDE SelectedControls
    '#INCLUDE SelectedFrameControl
    '#INCLUDE ActiveModule
    '#INCLUDE InputboxString
    If ActiveModule.Type <> vbext_ct_MSForm Then Exit Sub
    If SelectedControls.count <> 1 Then Exit Sub
    If TypeName(SelectedControl) <> "Frame" Then Exit Sub
    Dim Module As VBComponent
    Dim TargetFrame As MSForms.control
    If isSubFrame = False Then
        Set TargetFrame = SelectedControl
        Set Module = ActiveModule
    Else
        Set TargetFrame = SelectedFrameControl
        Set Module = ThisWorkbook.VBProject.VBComponents(TargetFrame.parent.parent.Name)
    End If
    Dim ControlNames As String
    ControlNames = InputboxString("Form Menus", "Type comma delimited menu names")
    If ControlNames = "" Then Exit Sub
    Dim var
    var = Split(ControlNames, ",")
    Dim i As Long
    For i = LBound(var) To UBound(var)
        var(i) = Trim(var(i))
    Next
    Dim coll As New Collection
    Dim element
    On Error Resume Next
    For Each element In var
        If Not IsNumeric(left(element, 1)) _
        And InStr(1, element, " ") = 0 Then
            coll.Add CStr(element), CStr(element)
        End If
    Next
    On Error GoTo 0
    Dim l As MSForms.control
    For Each element In coll
        Set l = Module.Designer.Controls(TargetFrame.Name).Controls.Add(ControlIDCommandButton, element)
        l.top = 7 + ((TargetFrame.Controls.count - 1) * l.Height)
        l.BackColor = vbWhite
    Next
End Sub

Sub UnderlineFrameName(form As Object, f As MSForms.control)
    If TypeName(form) = "VBComponent" Then
        Set Module = form
    Else
        Set Module = ThisWorkbook.VBProject.VBComponents(form.parent.Name)
    End If
    Set l = Module.Designer.Controls(f.Name).Controls.Add(ControlIDLabel)
    l.top = 6
    l.Height = 1
    l.Width = 100
    l.BackColor = vbWhite
    l.Tag = "skip"
End Sub


