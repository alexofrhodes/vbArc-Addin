Attribute VB_Name = "U_Answers"

Rem @Folder GetAnswers Declarations
Rem
Rem /*
Rem     Public Function answer( _
rem                             Optional AT As AnswerType = 999, _
rem                             Optional ExtraOptions As Variant, _
rem                             Optional ExtraOptionsPerColumn As Long = 999, _
rem                             Optional extraOptionsFramesVertical As Boolean = True, _
rem                             Optional Caption As String, _
rem                             optional AskInVBE as boolean) _
rem                                                         As Variant
Rem */
Rem /*
Rem If AnswerType=999 then show all AT options
Rem If AnswerType=0 then hide all AT options
Rem If you want to show only specific AT options use it like: call uAnswer.answer (argTrueFalse + argYesNo)
Rem */
Rem /*
Rem If you allow only one option then after you call uanswer.anser you can do
Rem
Rem dim myRespone
Rem     myResponse=SelectedOptionsCollection(1)
Rem If myRespone = ...
Rem
Rem Otherwise you can call uanswer.anser and loop all answers: For Each element In SelectedOptionsCollection
    Rem */
    Rem AskInVBE=True will show the userform in VBE window, no need to go back and forth to sheet
    Rem /*
    Rem ExtraOptions is passed as Array(ControlID1,Array(choice1,choice2...),ControlID2,Array(choiceX,choiceY...)
    Rem */
    Rem
    Public SelectedOptionsCollection As Collection

    Rem @Folder GetAnswers
Sub TestGetAnswer()
    '#INCLUDE PrintSelectedOptions
    Dim element
    Call uAnswer.answer(argInput, , , , , True)
    Debug.Print "You said: "
    PrintSelectedOptions
    Stop
    uAnswer.answer argDate, , , , "select start date"
    Debug.Print "Start date: "
    PrintSelectedOptions
    Stop
    Call uAnswer.answer(, Array(ControlIDToggleButton, Array(1, 2, 3, 4)), 2)
    Debug.Print "You selected toggles:"
    PrintSelectedOptions
    Stop
    Call uAnswer.answer(0, Array(ControlIDOptionButton, Array(1, 2, 3, 4), _
                                 ControlIDCheckBox, Array(5, 6, 7, 8), _
                                 ControlIDToggleButton, Array("a", "b", "c", "d")), _
                        2, False)
    PrintSelectedOptions
End Sub

Sub PrintSelectedOptions()
    If TypeName(SelectedOptionsCollection) <> "Nothing" Then
        If SelectedOptionsCollection.count > 0 Then
            For Each element In SelectedOptionsCollection: Debug.Print element: Next
        End If
    End If
End Sub

Sub GetSelectedOptions(Frame As Variant)
    Dim ctr As MSForms.control
    Dim out As New Collection
    For Each ctr In Frame.Controls
        If ctr.visible = True Then
            Select Case UCase(TypeName(ctr))
                Case UCase("CheckBox"), UCase("OptionButton"), UCase("ToggleButton")
                    If ctr.Value = True Then
                        If ctr.Name = "oInput" Then
                            SelectedOptionsCollection.Add Frame.Controls("Textbox1").Value
                        Else
                            SelectedOptionsCollection.Add ctr.Caption
                        End If
                    End If
            End Select
        End If
    Next
End Sub

Function CreateOrSetFrame(form As Object, Optional FrameName As String, Optional LTWH As Variant) As MSForms.Frame
    Dim cFrame As MSForms.Frame
    On Error Resume Next
    Set cFrame = form.Controls(FrameName)
    On Error GoTo 0
    If cFrame Is Nothing Then
        If TypeName(form) = "VBComponent" Then
            Set cFrame = form.Designer.Controls.Add("Forms.Frame.1")
        Else
            Set cFrame = form.Controls.Add("Forms.Frame.1")
        End If
    End If
    If Not IsMissing(FrameName) Then cFrame.Name = FrameName
    If Not IsMissing(LTWH) Then
        cFrame.left = LTWH(0)
        cFrame.top = LTWH(1)
        cFrame.Width = LTWH(2)
        cFrame.Height = LTWH(3)
    End If
    Set CreateOrSetFrame = cFrame
End Function

Function AvailableFormOrFrameRow(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myHeight
    For Each ctr In FormOrFrame.Controls
        If ctr.visible = True Then
            If ctr.left >= AfterWidth And ctr.top >= AfterHeight Then
                If ctr.top + ctr.Height > myHeight Then myHeight = ctr.top + ctr.Height
            End If
        End If
    Next
    AvailableFormOrFrameRow = myHeight + 6
End Function

Function AvailableFormOrFrameColumn(FormOrFrame As Object, Optional AfterWidth As Long = 0, Optional AfterHeight As Long = 0) As Long
    Dim ctr As MSForms.control
    Dim myWidth
    For Each ctr In FormOrFrame.Controls
        If ctr.visible = True Then
            If ctr.left >= AfterWidth And ctr.top >= AfterHeight Then
                If ctr.left + ctr.Width > myWidth Then myWidth = ctr.left + ctr.Width
            End If
        End If
    Next
    AvailableFormOrFrameColumn = myWidth + 6
End Function


