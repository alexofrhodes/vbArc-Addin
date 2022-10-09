Attribute VB_Name = "M_BarmanExample"
Rem @Folder BarmanExamples Declarations
Public Const Mname As String = "MyPopUpMenu"

Rem @Folder BarmanExamples
Sub DeletePopUpMenu()
    'Delete PopUp menu if it exist
    On Error Resume Next
    Application.CommandBars(Mname).Delete
    On Error GoTo 0
End Sub

Sub CreateDisplayPopUpMenu()
    'Delete PopUp menu if it exist
    '#INCLUDE DeletePopUpMenu
    '#INCLUDE Custom_PopUpMenu_1
    Call DeletePopUpMenu
    'Create the PopUpmenu
    Call Custom_PopUpMenu_1
    'Show the PopUp menu
    On Error Resume Next
    Application.CommandBars(Mname).ShowPopup
    On Error GoTo 0
End Sub

Sub Custom_PopUpMenu_1()
    '#INCLUDE TestMacro
    Dim MenuItem As CommandBarPopup
    'Add PopUp menu
    With Application.CommandBars.Add(Name:=Mname, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)
        'First add two buttons
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Button 1"
            .FaceId = 71
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
        End With
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Button 2"
            .FaceId = 72
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
        End With
        'Second Add menu with two buttons
        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
        With MenuItem
            .Caption = "My Special Menu"
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Button 1 in menu"
                .FaceId = 71
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
            End With
            With .Controls.Add(Type:=msoControlButton)
                .Caption = "Button 2 in menu"
                .FaceId = 72
                .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
            End With
        End With
        'Third add one button
        With .Controls.Add(Type:=msoControlButton)
            .Caption = "Button 3"
            .FaceId = 73
            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
        End With
    End With
End Sub

Sub TestMacro()
    MsgBox "Hi There, greetings from the Netherlands"
End Sub

Sub TestProcedure()
    MsgBox "ok"
End Sub

Sub DeleteNotBuiltInCommandbars()
    For Each b In Application.CommandBars
        If b.BuiltIn = False Then b.Delete
    Next
End Sub


