VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uAnswer 
   Caption         =   "Get Answers by vbArc"
   ClientHeight    =   5604
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   11976
   OleObjectBlob   =   "uAnswer.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uAnswer
'* Created    : 06-10-2022 10:34
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Rem @TODO replace old calendar method with new calendar uDatePicker
Rem         for new arguments StartDate and EndDate

Dim columnControls As Long
Dim ExtraOptionColumn As Long
Dim ExtraOptionRow As Long
Dim ctrType As String
Dim optionCaption As String
Dim c As MSForms.control
Dim cFrame As MSForms.Frame
Dim FrameControlColumn As Long
Dim FrameControlRow As Long
Dim cFrameTop As Long
Dim cFrameLeft As Long
Dim PreviousRow As Long

Private WithEvents Calendar1 As cCalendar
Attribute Calendar1.VB_VarHelpID = -1
Private ans As Variant
Const ControlIDCheckBox = "Forms.CheckBox.1"
Const ControlIDComboBox = "Forms.ComboBox.1"
Const ControlIDCommandButton = "Forms.CommandButton.1"
Const ControlIDFrame = "Forms.Frame.1"
Const ControlIDImage = "Forms.Image.1"
Const ControlIDLabel = "Forms.Label.1"
Const ControlIDListBox = "Forms.ListBox.1"
Const ControlIDMultiPage = "Forms.MultiPage.1"
Const ControlIDOptionButton = "Forms.OptionButton.1"
Const ControlIDScrollBar = "Forms.ScrollBar.1"
Const ControlIDSpinButton = "Forms.SpinButton.1"
Const ControlIDTabStrip = "Forms.TabStrip.1"
Const ControlIDTextBox = "Forms.TextBox.1"
Const ControlIDToggleButton = "Forms.ToggleButton.1"
Enum AnswerType
    argInput = 2
    argYesNo = 4
    argCancel = 8
    argTrueFalse = 16
    argDate = 32
    argRange = 64
End Enum

Private msFontName As String
Private mafChrWid(32 To 127) As Double

Private Sub oDate_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oFalse_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oInput_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oNo_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oRange_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oTrue_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub oYes_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Me.ActiveControl.Value = False
    Cancel = True
End Sub

Private Sub UserForm_Initialize()
    Set Calendar1 = New cCalendar
    With Calendar1
        .Add_Calendar_into_Frame Me.Frame1
        .UseDefaultBackColors = True
        .DayLength = 3
        .MonthLength = mlENShort
    End With
    Frame1.visible = False
    Dim ctr As MSForms.control
    For Each ctr In Frame1.Controls
        ctr.visible = Frame1.visible
    Next
End Sub

Private Sub TextBox1_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Then cmdOK_Click
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Set Calendar1 = Nothing
End Sub

Public Function answer( _
       Optional AT As AnswerType = 999, _
       Optional ExtraOptions As Variant, _
       Optional ExtraOptionsPerColumn As Long = 999, _
       Optional extraOptionsFramesVertical As Boolean = True, _
       Optional Caption As String, _
       Optional AskInVBE As Boolean) _
        As Variant
    '#INCLUDE LargestLength
    '#INCLUDE ResizeUserformToFitControls
    '#INCLUDE AddAnswerControlToFrame
    '#INCLUDE StrWidth
    '#INCLUDE MakeUserFormChildOfVBEditor
    '#INCLUDE CreateOrSetFrame
    '#INCLUDE AvailableFormOrFrameRow
    '#INCLUDE AvailableFormOrFrameColumn

                                                    
    Rem example in immediate window:
    Rem uanswer.answer(,Array(array("Forms.OptionButton.1",1,2,3,4,5,6)),2)
    Rem call uanswer.answer(0,array(ControlIDOptionButton,array(1,2,3,4),ControlIDCheckBox,array(5,6,7,8),ControlIDOptionButton,array("a","b","c","d")),2)

    Rem if AT = 0 then don

    If AT = 999 Then AT = argDate + argInput + argRange + argTrueFalse + argYesNo + argCancel
    If AT And argDate Then oDate.visible = True: oDate.Value = True
    If AT And argInput Then oInput.visible = True: TextBox1.visible = True
    If AT And argRange Then oRange.visible = True
    If AT And argTrueFalse Then oTrue.visible = True: oFalse.visible = True
    If AT And argYesNo Then oYes.visible = True: oNo.visible = True
    If AT And argCancel Then oCancel.visible = True
    
    If Not IsMissing(ExtraOptions) Then
        If IsArray(ExtraOptions) Then
            Dim ExtraOptionsGroupCount
            ExtraOptionsGroupCount = UBound(ExtraOptions)
            
            Dim counter As Long
            Dim i As Long
            Dim X As Long
            
            Dim groupCounter As Long
            Dim ControlsCount As Long
    
            ExtraOptionColumn = 186
            Dim ExtraColumnWidth As Long
            
            Dim arrayRow As Long
            
            cFrameTop = 0
            cFrameLeft = ExtraOptionColumn
            
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If Not IsArray(ExtraOptions(X)) Then
                    ctrType = ExtraOptions(X)
                    Set cFrame = CreateOrSetFrame(Me, "myFrame" & X)
                    cFrame.left = cFrameLeft
                    cFrame.top = cFrameTop
                    arrayRow = ExtraOptionRow
                    If IsArray(ExtraOptions(X + 1)) Then
                        ExtraColumnWidth = LargestLength(ExtraOptions(X + 1)) * StrWidth("A", oDate.Font.Name, oDate.Font.Size)
                        For i = 0 To UBound(ExtraOptions(X + 1))
                            ControlsCount = Me.Controls.count
                            optionCaption = ExtraOptions(X + 1)(i)
                            AddAnswerControlToFrame Me, "myFrame" & X
                                
                            If columnControls = ExtraOptionsPerColumn Then
                                columnControls = 0
                                PreviousRow = 0
                                ExtraOptionColumn = ExtraOptionColumn + c.Width + ExtraColumnWidth
                                FrameControlColumn = FrameControlColumn + c.Width + ExtraColumnWidth
                            Else
                            End If
                             
                        Next i
                    Else
                    End If
                    
                End If
                ExtraOptionColumn = 186
                If Not cFrame Is Nothing Then
                    ExtraOptionRow = cFrame.top + cFrame.Height + 6
                Else
                    ExtraOptionRow = AvailableFormOrFrameRow(Me, 185)
                End If
                If Not cFrame Is Nothing Then
                    cFrameTop = cFrame.top + cFrame.Height + 6
                    If extraOptionsFramesVertical = False Then cFrameLeft = cFrame.left + cFrame.Width + 6
                End If
                FrameControlColumn = 0
                Set cFrame = Nothing
            Next X
        End If

    End If
    
    Dim ctr As MSForms.control
    If Not IsMissing(ExtraOptions) Then
        If extraOptionsFramesVertical = True Then
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If WorksheetFunction.IsEven(X) Then Me.Controls("myFrame" & X).visible = False
            Next
            anchorcolumn = IIf(AT = 0, cmdOK.left + cmdOK.Width + 6, AvailableFormOrFrameColumn(Me, cmdOK.left + cmdOK.Width))
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If WorksheetFunction.IsEven(X) Then
                    Me.Controls("myFrame" & X).visible = True
                    Me.Controls("myFrame" & X).left = anchorcolumn
                End If
            Next
        Else
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If WorksheetFunction.IsEven(X) Then
                    Me.Controls("myFrame" & X).top = 6
                End If
            Next
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If WorksheetFunction.IsEven(X) Then Me.Controls("myFrame" & X).visible = False
            Next
            For X = LBound(ExtraOptions) To UBound(ExtraOptions)
                If WorksheetFunction.IsEven(X) Then
                    Me.Controls("myFrame" & X).left = AvailableFormOrFrameColumn(Me)
                    Me.Controls("myFrame" & X).visible = True
                End If
            Next
        End If
    End If
    
    If Len(Caption) > 0 Then uAnswer.Caption = Caption
    ResizeUserformToFitControls Me
    If AskInVBE = True Then
        Application.VBE.MainWindow.visible = True
        Application.VBE.MainWindow.WindowState = vbext_ws_Normal
        MakeUserFormChildOfVBEditor Me.Caption
    End If
    Me.Show
    If IsObject(ans) Then
        Set answer = ans
    Else
        answer = ans
    End If
    Unload Me
End Function

Sub AddAnswerControlToFrame(form As Object, FrameName As String)
    '#INCLUDE ResizeUserformToFitControls
    '#INCLUDE CreateOrSetFrame
    Set cFrame = CreateOrSetFrame(form, FrameName)
    Set c = cFrame.Controls.Add(ctrType)
        
    c.Font.Name = "Consolas"
    c.Font.Size = 9
    columnControls = columnControls + 1
    FrameControlRow = (columnControls - 1) * c.Height
    c.Caption = optionCaption
    c.AutoSize = True
    c.left = FrameControlColumn
    c.top = FrameControlRow
    ResizeUserformToFitControls Me.Controls(FrameName)
End Sub

Sub BringControlsLeft(FormOrFrame As Object)
    '#INCLUDE AvailableFormOrFrameColumn
    Dim c As MSForms.control
    Dim element
    For Each c In FormOrFrame.Controls
        c.left = AvailableFormOrFrameColumn(FormOrFrame)
    Next
End Sub

Private Sub cmdOK_Click()
    Rem example call uanswer.answer(0,array(ControlIDOptionButton,array(1,2,3,4),ControlIDCheckBox,array(5,6,7,8),ControlIDOptionButton,array("a","b","c","d")),2)
    Rem after this selectedoptionscollection is populated with the options chosen for you to use

    Set SelectedOptionsCollection = New Collection
    GetSelectedOptions Me
    For i = 1 To SelectedOptionsCollection.count
        ans = SelectedOptionsCollection(i)
        
        If ans = "INPUT" Then
            ans = TextBox1.TEXT
        ElseIf ans = "RANGE" Then
            Set ans = InputBoxRange
        ElseIf ans = "vbYES" Then
            ans = vbYes
        ElseIf ans = "vbNO" Then
            ans = vbNo
        ElseIf ans = "vbCancel" Then
            ans = vbCancel
        ElseIf ans = "TRUE" Then
            ans = True
        ElseIf ans = "FALSE" Then
            ans = False
        ElseIf ans = "DATE" Then
            ans = Calendar1.Value
        End If
        
        If TypeName(ans) = "Boolean" Then
            ans = CBool(ans)
        ElseIf IsDate(ans) Then
            ans = CDate(ans)
        ElseIf IsNumeric(ans) Then
            ans = CLng(ans)
        ElseIf TypeName(ans) = "String" Then
            ans = CStr(ans)
        End If
        SelectedOptionsCollection.Add ans, , i
        SelectedOptionsCollection.Remove i + 1
    Next
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Function InputBoxRange(Optional sTitle As String, Optional sPrompt As String) As Range
    On Error Resume Next
    Set InputBoxRange = Application.InputBox(title:=sTitle, Prompt:=sPrompt, Type:=8, _
                                             Default:=IIf(TypeName(Selection) = "Range", Selection.Address, ""))
End Function

Private Sub cmdToday_Click()
    Calendar1.Year = Format(Date, "YYYY")
    Calendar1.Month = Format(Date, "MM")
    Calendar1.Day = Format(Date, "DD")
End Sub

Private Sub oDate_Change()
    Dim ctr As MSForms.control
    For Each ctr In Frame1.Controls
        ctr.visible = oDate.Value
    Next
    cmdToday.visible = oDate.Value
    Frame1.visible = oDate.Value
    ResizeUserformToFitControls Me
End Sub

Private Function StrWidth(s As String, sFontName As String, fFontSize As Double) As Double
    '#INCLUDE InitChrWidths
    Dim i As Long
    Dim j As Long

    If sFontName <> msFontName Then
        If Not InitChrWidths(sFontName) Then
            Exit Function
        End If
    End If
    For i = 1 To Len(s)
        j = Asc(Mid(s, i, 1))
        If j >= 32 Then
            StrWidth = StrWidth + fFontSize * mafChrWid(j)
        End If
    Next i
End Function

Private Function InitChrWidths(sFontName As String) As Boolean
    '#INCLUDE StrWidth
    Dim i As Long
    Select Case sFontName
        Case "Consolas"
            For i = 32 To 127
                Select Case i
                    Case 32 To 127
                        mafChrWid(i) = 0.5634
                End Select
            Next i
        Case "Arial"
            For i = 32 To 127
                Select Case i
                    Case 39, 106, 108
                        mafChrWid(i) = 0.1902
                    Case 105, 116
                        mafChrWid(i) = 0.2526
                    Case 32, 33, 44, 46, 47, 58, 59, 73, 91 To 93, 102, 124
                        mafChrWid(i) = 0.3144
                    Case 34, 40, 41, 45, 96, 114, 123, 125
                        mafChrWid(i) = 0.3768
                    Case 42, 94, 118, 120
                        mafChrWid(i) = 0.4392
                    Case 107, 115, 122
                        mafChrWid(i) = 0.501
                    Case 35, 36, 48 To 57, 63, 74, 76, 84, 90, 95, 97 To 101, 103, 104, 110 To 113, 117, 121
                        mafChrWid(i) = 0.5634
                    Case 43, 60 To 62, 70, 126
                        mafChrWid(i) = 0.6252
                    Case 38, 65, 66, 69, 72, 75, 78, 80, 82, 83, 85, 86, 88, 89, 119
                        mafChrWid(i) = 0.6876
                    Case 67, 68, 71, 79, 81
                        mafChrWid(i) = 0.7494
                    Case 77, 109, 127
                        mafChrWid(i) = 0.8118
                    Case 37
                        mafChrWid(i) = 0.936
                    Case 64, 87
                        mafChrWid(i) = 1.0602
                End Select
            Next i
        Case "Calibri"
            For i = 32 To 127
                Select Case i
                    Case 32, 39, 44, 46, 73, 105, 106, 108
                        mafChrWid(i) = 0.2526
                    Case 40, 41, 45, 58, 59, 74, 91, 93, 96, 102, 123, 125
                        mafChrWid(i) = 0.3144
                    Case 33, 114, 116
                        mafChrWid(i) = 0.3768
                    Case 34, 47, 76, 92, 99, 115, 120, 122
                        mafChrWid(i) = 0.4392
                    Case 35, 42, 43, 60 To 63, 69, 70, 83, 84, 89, 90, 94, 95, 97, 101, 103, 107, 118, 121, 124, 126
                        mafChrWid(i) = 0.501
                    Case 36, 48 To 57, 66, 67, 75, 80, 82, 88, 98, 100, 104, 110 To 113, 117, 127
                        mafChrWid(i) = 0.5634
                    Case 65, 68, 86
                        mafChrWid(i) = 0.6252
                    Case 71, 72, 78, 79, 81, 85
                        mafChrWid(i) = 0.6876
                    Case 37, 38, 119
                        mafChrWid(i) = 0.7494
                    Case 109
                        mafChrWid(i) = 0.8742
                    Case 64, 77, 87
                        mafChrWid(i) = 0.936
                End Select
            Next i
        Case "Tahoma"
            For i = 32 To 127
                Select Case i
                    Case 39, 105, 108
                        mafChrWid(i) = 0.2526
                    Case 32, 44, 46, 102, 106
                        mafChrWid(i) = 0.3144
                    Case 33, 45, 58, 59, 73, 114, 116
                        mafChrWid(i) = 0.3768
                    Case 34, 40, 41, 47, 74, 91 To 93, 124
                        mafChrWid(i) = 0.4392
                    Case 63, 76, 99, 107, 115, 118, 120 To 123, 125
                        mafChrWid(i) = 0.501
                    Case 36, 42, 48 To 57, 70, 80, 83, 95 To 98, 100, 101, 103, 104, 110 To 113, 117
                        mafChrWid(i) = 0.5634
                    Case 66, 67, 69, 75, 84, 86, 88, 89, 90
                        mafChrWid(i) = 0.6252
                    Case 38, 65, 71, 72, 78, 82, 85
                        mafChrWid(i) = 0.6876
                    Case 35, 43, 60 To 62, 68, 79, 81, 94, 126
                        mafChrWid(i) = 0.7494
                    Case 77, 119
                        mafChrWid(i) = 0.8118
                    Case 109
                        mafChrWid(i) = 0.8742
                    Case 64, 87
                        mafChrWid(i) = 0.936
                    Case 37, 127
                        mafChrWid(i) = 1.0602
                End Select
            Next i
        Case "Lucida Console"
            For i = 32 To 127
                Select Case i
                    Case 32 To 127
                        mafChrWid(i) = 0.6252
                End Select
            Next i
        Case "Times New Roman"
            For i = 32 To 127
                Select Case i
                    Case 39, 124
                        mafChrWid(i) = 0.1902
                    Case 32, 44, 46, 59
                        mafChrWid(i) = 0.2526
                    Case 33, 34, 47, 58, 73, 91 To 93, 105, 106, 108, 116
                        mafChrWid(i) = 0.3144
                    Case 40, 41, 45, 96, 102, 114
                        mafChrWid(i) = 0.3768
                    Case 63, 74, 97, 115, 118, 122
                        mafChrWid(i) = 0.4392
                    Case 94, 98 To 101, 103, 104, 107, 110, 112, 113, 117, 120, 121, 123, 125
                        mafChrWid(i) = 0.501
                    Case 35, 36, 42, 48 To 57, 70, 83, 84, 95, 111, 126
                        mafChrWid(i) = 0.5634
                    Case 43, 60 To 62, 69, 76, 80, 90
                        mafChrWid(i) = 0.6252
                    Case 65 To 67, 82, 86, 89, 119
                        mafChrWid(i) = 0.6876
                    Case 68, 71, 72, 75, 78, 79, 81, 85, 88
                        mafChrWid(i) = 0.7494
                    Case 38, 109, 127
                        mafChrWid(i) = 0.8118
                    Case 37
                        mafChrWid(i) = 0.8742
                    Case 64, 77
                        mafChrWid(i) = 0.936
                    Case 87
                        mafChrWid(i) = 0.9984
                End Select
            Next i
        Case Else
            MsgBox "Font name """ & sFontName & """ not available!", vbCritical, "StrWidth"
            Exit Function
    End Select
    msFontName = sFontName
    InitChrWidths = True
End Function


