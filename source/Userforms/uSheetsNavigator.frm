VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uSheetsNavigator 
   Caption         =   "Sheet Nav"
   ClientHeight    =   6372
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   2964
   OleObjectBlob   =   "uSheetsNavigator.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "uSheetsNavigator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uSheetsNavigator
'* Created    : 06-10-2022 10:40
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Sub SortListbox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
    On Error GoTo eh
    Dim vaItems As Variant
    Dim i As Long, j As Long
    Dim c As Integer
    Dim vTemp As Variant
    vaItems = oLb.list
    If sType = 1 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                If sDir = 1 Then
                    If vaItems(i, sCol) > vaItems(j, sCol) Then
                        For c = 0 To oLb.columnCount - 1
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                ElseIf sDir = 2 Then
                    If vaItems(i, sCol) < vaItems(j, sCol) Then
                        For c = 0 To oLb.columnCount - 1
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
            Next j
        Next i
    ElseIf sType = 2 Then
        For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
            For j = i + 1 To UBound(vaItems, 1)
                If sDir = 1 Then
                    If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.columnCount - 1
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                ElseIf sDir = 2 Then
                    If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
                        For c = 0 To oLb.columnCount - 1
                            vTemp = vaItems(i, c)
                            vaItems(i, c) = vaItems(j, c)
                            vaItems(j, c) = vTemp
                        Next c
                    End If
                End If
            Next j
        Next i
    End If
    oLb.list = vaItems
    Exit Sub
eh:
    LoadSheetBox
End Sub

Sub LoadSheetBox()
    SheetBox.clear
    Dim sh As Worksheet
    For Each sh In ActiveWorkbook.SHEETS
        If sh.visible = xlSheetVisible Then SheetBox.AddItem sh.Name
    Next
End Sub

Sub SortSheetBox()
    If Me.oDefault.Value = True Then
        Call LoadSheetBox
    Else
        Dim lBox As MSForms.ListBox
        Set lBox = Me.SheetBox
        Dim OnColumn As Integer
        OnColumn = 0
        Dim TextOrNumbers As Integer
        TextOrNumbers = 1
        Dim AscendingOrDescending As Integer
        AscendingOrDescending = IIf(Me.oAZ.Value = True, 1, 2)
        Call SortListbox(lBox, OnColumn, TextOrNumbers, AscendingOrDescending)
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim ans As String
    ans = MsgBox("No undo." & vbNewLine & _
                 "You may only close the workbook without changes to restore original order." & vbNewLine & vbNewLine & "Proceed?", vbYesNo)
    If ans = vbNo Then Exit Sub
    Dim i As Long
    For i = 0 To SheetBox.ListCount - 1
        ActiveWorkbook.SHEETS(SheetBox.list(i)).Move Before:=SHEETS(i + 1)
    Next i
End Sub

Private Sub CommandButton5_Click()
    TextBox2.TEXT = ""
End Sub

Private Sub CommandButton6_Click()
    On Error Resume Next
    Dim note As String
    note = ActiveSheet.Name
    With TextBox1
        ActiveWorkbook.SHEETS(.TEXT).Activate
        .TEXT = note
    End With
End Sub

Private Sub CommandButton7_Click()
    LoadSheetBox
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub SheetBox_Click()
    TextBox1.TEXT = ActiveSheet.Name
    With SheetBox
        ActiveWorkbook.SHEETS(.list(.ListIndex)).Activate
    End With
End Sub

Private Sub oAZ_Click()
    SortSheetBox
End Sub

Private Sub oDefault_Click()
    SortSheetBox
End Sub

Private Sub oZA_Click()
    SortSheetBox
End Sub

Private Sub TextBox2_Change()
    LoadSheetBox
    For i = SheetBox.ListCount - 1 To 0 Step -1
        If InStr(1, LCase(SheetBox.list(i)), LCase(TextBox2.TEXT)) = 0 Then
            SheetBox.RemoveItem (i)
        End If
    Next
End Sub

Private Sub UserForm_Initialize()
    Call LoadUserformPosition
    TextBox1.TEXT = ActiveSheet.Name
    Call LoadSheetBox
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Call SaveUserformPosition
End Sub

Sub LoadUserformPosition()
    If GetSetting("My Settings Folder", Me.Name, "Left Position") = "" _
                                                                    And GetSetting("My Settings Folder", Me.Name, "Top Position") = "" Then
        Me.StartUpPosition = 1
    Else
        Me.left = GetSetting("My Settings Folder", Me.Name, "Left Position")
        Me.top = GetSetting("My Settings Folder", Me.Name, "Top Position")
    End If
End Sub

Sub SaveUserformPosition()
    SaveSetting "My Settings Folder", Me.Name, "Left Position", Me.left
    SaveSetting "My Settings Folder", Me.Name, "Top Position", Me.top
End Sub


