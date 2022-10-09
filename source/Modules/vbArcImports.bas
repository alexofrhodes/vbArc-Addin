Attribute VB_Name = "vbArcImports"
Rem @Folder vbArcImports
Sub ImportSelectedProcedure()
    '#INCLUDE CodepaneSelection
    '#INCLUDE CopyProcedures
    CopyProcedures CodepaneSelection, Workbooks("projectstart.xlam"), ThisWorkbook, False
End Sub

Function CheckPath(Path) As String
    '#INCLUDE HttpExists
    '#INCLUDE FileExists
    '#INCLUDE FolderExists
    Dim retval
    retval = "I"
    If (retval = "I") And FileExists(Path) Then retval = "F"
    If (retval = "I") And FolderExists(Path) Then retval = "D"
    If (retval = "I") And HttpExists(Path) Then retval = "U"
    CheckPath = retval
End Function

Function HttpExists(ByVal sURL As String) As Boolean
    Dim oXHTTP As Object
    Set oXHTTP = CreateObject("MSXML2.XMLHTTP")
    If Not UCase(sURL) Like "HTTP:*" Then
        sURL = "http://" & sURL
    End If
    On Error GoTo haveError
    oXHTTP.Open "HEAD", sURL, False
    oXHTTP.send
    HttpExists = IIf(oXHTTP.Status = 200, True, False)
    Exit Function
haveError:
    Rem Debug.Print err.Description
    HttpExists = False
End Function

Function TXTReadFromUrl(URL As String) As String
    On Error GoTo Err_GetFromWebpage
    Dim objWeb As Object
    Dim strXML As String
    Set objWeb = CreateObject("Msxml2.ServerXMLHTTP")
    objWeb.Open "GET", URL, False
    objWeb.setRequestHeader "Content-Type", "text/xml"
    objWeb.setRequestHeader "Cache-Control", "no-cache"
    objWeb.setRequestHeader "Pragma", "no-cache"
    objWeb.send
    strXML = objWeb.responseText
    TXTReadFromUrl = strXML
End_GetFromWebpage:
    Set objWeb = Nothing
    Exit Function
Err_GetFromWebpage:
    MsgBox err.Description & " (" & err.Number & ")"
    Resume End_GetFromWebpage
End Function

