Attribute VB_Name = "U_Addins"
Rem @Folder Addins Declarations
#If VBA7 Then
    Private Declare PtrSafe Function URLDownloadToFile Lib "urlmon" _
    Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
    ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#Else
    Private Declare Function URLDownloadToFile Lib "urlmon" _
                             Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, _
                                                         ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If

Rem @Folder Addins
Sub AddinManagerButtonClicked(control As IRibbonControl)
    Select Case UCase(control.ID)
        Case UCase("buttonAddinManager")
            uAddinManager.Show
        Case UCase("AddinManagerToggle")
            ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
    End Select
End Sub

Sub DownloadFile(FileUrl As String, SaveAs As String)
    URLDownloadToFile 0, FileUrl, SaveAs, 0, 0
End Sub

Function vbArcAddins() As Variant
    '#INCLUDE TXTReadFromUrl
    Dim v
    v = Filter(Split(TXTReadFromUrl("https://github.com/alexofrhodes/vbArc-addins/raw/main/ListOfAddins.txt"), "  " & vbLf), "xlam", True)
    Dim i As Long
    vbArcAddins = v
End Function

Sub DownloadFileFromURL(FileUrl, saveFullName)
    Dim objXmlHttpReq As Object
    Dim objStream As Object
    Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
    objXmlHttpReq.Open "GET", FileUrl, False, "username", "password"
    objXmlHttpReq.send
    If objXmlHttpReq.Status = 200 Then
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Open
        objStream.Type = 1
        objStream.Write objXmlHttpReq.responseBody
        objStream.SaveToFile saveFullName, 2
        objStream.Close
    End If
End Sub

Public Function LastModified(filespec)
End Function

Sub AddinsModified()
    On Error Resume Next
    Application.CalculateFull
    On Error GoTo 0
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim cell As Range
    Set cell = Sheet1.Range("E2")
    Dim ans As Long
    Dim counter As Long
    Dim ctr As MSForms.control
    Dim ReplaceWhat As String, ReplaceWith As String
    Dim spaceCounter As Long
    spaceCounter = ThisWorkbook.SHEETS("uAddins_Settings").Range("K1")
    spaceCounter = spaceCounter + 1
    Do While cell <> ""
        If cell.TEXT = "A" Then
            counter = counter + 1
            Set ctr = uUpdateFiles.Controls.Add("Forms.CheckBox.1")
            ctr.Caption = Sheet1.Range("A1") & vbTab & cell.OFFSET(0, -4) & vbTab & vbTab & cell.OFFSET(0, -2) & vbTab & vbTab & _
                                                                                                               Sheet1.Range("G1") & vbTab & vbTab & cell.OFFSET(0, 2) & vbTab & vbTab & cell.OFFSET(0, 4)
            ctr.left = 6
            ctr.top = 36 - ctr.Height + (counter * ctr.Height)
            ctr.Height = ctr.Height
            ctr.Width = 1000
            ctr.WordWrap = False
            ctr.AutoSize = True
        ElseIf cell.TEXT = "B" Then
            counter = counter + 1
            Set ctr = uUpdateFiles.Controls.Add("Forms.CheckBox.1")
            ReplaceWhat = Sheet1.Range("G1") & cell.OFFSET(0, 2)
            ReplaceWith = Sheet1.Range("A1") & cell.OFFSET(0, -4)
            ctr.Font.Name = "Consolas"
            ctr.Caption = ReplaceWhat & Space(spaceCounter - Len(ReplaceWhat)) & vbTab & cell.OFFSET(0, 4) & vbTab & _
                                                                                                           ReplaceWith & Space(spaceCounter - Len(ReplaceWhat)) & vbTab & cell.OFFSET(0, -2)
            ctr.left = 6
            ctr.top = 36 - ctr.Height + (counter * ctr.Height)
            ctr.Height = ctr.Height
            ctr.Width = 1000
            ctr.WordWrap = False
            ctr.AutoSize = True
        End If
        Set cell = cell.OFFSET(1, 0)
    Loop
    Application.CalculateFull
End Sub

Sub UpdateFiles()
    '#INCLUDE WorkbookIsOpen
    Dim c As MSForms.control
    Dim cell As Range
    Dim r As Range
    Dim workbookName As String, WorkbookFullName As String
    Dim LocalWorkbook As String, GithubWorkbook As String
    For Each c In uUpdateFiles.Controls
        If UCase(TypeName(c)) = UCase("CheckBox") Then
            If c.Value = True Then
                FindMe = Mid(Split(c.Caption, " ")(0), InStrRev(Split(c.Caption, " ")(0), "\") + 1)
                For Each r In ThisWorkbook.SHEETS("uAddins_Settings").Range("A1").CurrentRegion.RESIZE(, 1)
                    If r.TEXT = FindMe Then
                        Set cell = r
                        Exit For
                    End If
                Next
                LocalWorkbook = cell.OFFSET(0, 1)
                GithubWorkbook = cell.OFFSET(0, 7)
                workbookName = cell.TEXT
                Dim WasOpen As Boolean
                Dim fso As Object
                Set fso = CreateObject("Scripting.FileSystemObject")
                If cell.OFFSET(0, 4) = "A" Then
                    WasOpen = WorkbookIsOpen(workbookName)
                    If WasOpen = True Then
                        Workbooks(workbookName).IsAddin = True
                        Workbooks(workbookName).Close
                    End If
                    fso.CopyFile GithubWorkbook, LocalWorkbook, True
                    If WasOpen = True Then Workbooks.Open LocalWorkbook
                ElseIf cell.OFFSET(0, 4) = "B" Then
                    fso.CopyFile LocalWorkbook, GithubWorkbook, True
                End If
                If err.Number = 0 Then uUpdateFiles.Controls.Remove c.Name
            End If
        End If
    Next
End Sub


