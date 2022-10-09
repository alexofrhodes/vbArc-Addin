Attribute VB_Name = "F_Declararions"
Rem @Folder Declarations
Sub ListDeclarationsToSheet(TargetWorkbook As Workbook)
    '#INCLUDE ArrayToRange2D
    '#INCLUDE CreateOrSetSheet
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("DeclarationsList", ActiveWorkbook)
    ArrayToRange2D CollectionsToArrayTable(getDeclarations(TargetWorkbook, True, True, True, True, True, True)), ws.Range("A1")
    ws.Range("A1").CurrentRegion.Sort ws.Range("D1")
End Sub

Sub testGetDeclarations()
    '#INCLUDE CreateOrSetSheet
    '#INCLUDE getDeclarations
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("DeclarationsTest", ThisWorkbook)
    Dim element As Variant
    Dim subelement As Variant
    Dim r As Long, c As Long
    r = 1
    c = 1
    For Each element In getDeclarations(ThisWorkbook, True, True, True, True, True, True)
        For Each subelement In element
            ws.Cells(r, c) = subelement
            r = r + 1
        Next
        r = 1
        c = c + 1
    Next
End Sub

Function getDeclaredKeywordsOfWorkbook(TargetWorkbook As Workbook) As Variant
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    getDeclaredKeywordsOfWorkbook = WorksheetFunction.Transpose(CollectionsToArrayTable(getDeclarations(TargetWorkbook, , , True)))
End Function

Function getDeclaredEnumOfWorkbook(TargetWorkbook As Workbook) As String
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    Dim out As String
    For Each c In CollectionsToArrayTable(getDeclarations(ThisWorkbook, , , , True))
        If InStr(1, CStr(c), "Enum ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredEnumOfWorkbook = out
End Function

Function getDeclaredTypeOfWorkbook(TargetWorkbook As Workbook) As String
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    Dim out As String
    For Each c In CollectionsToArrayTable(getDeclarations(ThisWorkbook, , , , True))
        If InStr(1, CStr(c), "Type ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredTypeOfWorkbook = out
End Function

Function getDeclaredSubOfWorkbook(TargetWorkbook As Workbook) As String
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    Dim out As String
    For Each c In CollectionsToArrayTable(getDeclarations(ThisWorkbook, , , , True))
        If InStr(1, CStr(c), "Sub ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredSubOfWorkbook = out
End Function

Function getDeclaredFunctionOfWorkbook(TargetWorkbook As Workbook) As String
    '#INCLUDE CollectionsToArrayTable
    '#INCLUDE getDeclarations
    Dim out As String
    For Each c In CollectionsToArrayTable(getDeclarations(ThisWorkbook, , , , True))
        If InStr(1, CStr(c), "Function ") > 0 Then out = IIf(out = "", c, out & vbNewLine & c)
    Next
    getDeclaredFunctionOfWorkbook = out
End Function

Function getDeclarations( _
         wb As Workbook, _
         Optional includeScope As Boolean, _
         Optional includeType As Boolean, _
         Optional includeKeywords As Boolean, _
         Optional includeDeclarations As Boolean, _
         Optional includeComponentName As Boolean, _
         Optional includeComponentType As Boolean) _
        As Collection
    '#INCLUDE ComponentTypeToString
    '#INCLUDE getWord
    Dim output As Collection: Set output = New Collection
    Dim declarationsCollection As Collection: Set declarationsCollection = New Collection
    Dim keywordsCollection As Collection: Set keywordsCollection = New Collection
    Dim vbComp As VBComponent
    Dim CodeMod As CodeModule
    Dim str As Variant
    Dim i As Long
    Dim element As Variant
    Dim originalDeclarations As Variant
    Dim tmp As String
    Dim helper As String
    Dim typeCollection As Collection: Set typeCollection = New Collection
    Dim componentCollection As Collection: Set componentCollection = New Collection
    Dim componentTypeCollection As Collection: Set componentTypeCollection = New Collection
    Dim scopeCollection As Collection: Set scopeCollection = New Collection
    For Each vbComp In wb.VBProject.VBComponents
        If vbComp.Type <> vbext_ct_ClassModule And vbComp.Type <> vbext_ct_Document Then
            Set CodeMod = vbComp.CodeModule
            If CodeMod.CountOfDeclarationLines > 0 Then
                str = CodeMod.Lines(1, CodeMod.CountOfDeclarationLines)
                str = Replace(str, "_" & vbNewLine, "")
                originalDeclarations = str
                tmp = str
                Do While InStr(1, str, "End Type") > 0
                    tmp = Mid(str, InStr(1, str, "Type "), InStr(1, str, "End Type") - InStr(1, str, "Type ") + 8)
                    str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, str, "End Enum") > 0
                    tmp = Mid(str, InStr(1, str, "Enum "), InStr(1, str, "End Enum") - InStr(1, str, "Enum ") + 8)
                    str = Replace(str, tmp, Split(tmp, vbNewLine)(0))
                Loop
                Do While InStr(1, str, "  ") > 0
                    str = Replace(str, "  ", " ")
                Loop
                str = Split(str, vbNewLine)
                tmp = originalDeclarations
                For Each element In str
                    If InStr(1, CStr(element), "Enum ", vbTextCompare) > 0 Then
                        keywordsCollection.Add getWord(CStr(element), " ", "Enum")
                        declarationsCollection.Add getWord(tmp, , "Enum " & keywordsCollection.item(keywordsCollection.count), "End Enum", , , True)
                        typeCollection.Add "Enum"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf InStr(1, CStr(element), "Type ", vbTextCompare) > 0 Then
                        keywordsCollection.Add getWord(CStr(element), " ", "Type")
                        declarationsCollection.Add getWord(tmp, , "Type " & keywordsCollection.item(keywordsCollection.count), "End Type", , , True)
                        typeCollection.Add "Type"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf InStr(1, CStr(element), "Const ", vbTextCompare) > 0 Then
                        keywordsCollection.Add getWord(CStr(element), " ", "Const")
                        declarationsCollection.Add CStr(element)
                        typeCollection.Add "Const"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf InStr(1, CStr(element), "Sub ", vbTextCompare) > 0 Then
                        keywordsCollection.Add getWord(CStr(element), " ", "Sub")
                        declarationsCollection.Add CStr(element)
                        typeCollection.Add "Sub"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf InStr(1, CStr(element), "Function ", vbTextCompare) > 0 Then
                        keywordsCollection.Add getWord(CStr(element), " ", "Function")
                        declarationsCollection.Add CStr(element)
                        typeCollection.Add "Function"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf element Like "*(*) As *" Then
                        helper = left(element, InStr(1, CStr(element), "(") - 1)
                        helper = Mid(helper, InStrRev(helper, " ") + 1)
                        keywordsCollection.Add helper
                        declarationsCollection.Add CStr(element)
                        typeCollection.Add "Other"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    ElseIf element Like "* As *" Then
                        keywordsCollection.Add getWord(CStr(element), " ", , "As")
                        declarationsCollection.Add CStr(element)
                        typeCollection.Add "Other"
                        componentCollection.Add vbComp.Name
                        componentTypeCollection.Add ComponentTypeToString(vbComp.Type)
                        scopeCollection.Add IIf(InStr(1, declarationsCollection.item(declarationsCollection.count), "Public", vbTextCompare), "Public", "Private")
                    Else
                    End If
                Next element
            End If
        End If
    Next vbComp
    If includeScope = True Then output.Add scopeCollection
    If includeType = True Then output.Add typeCollection
    If includeKeywords = True Then output.Add keywordsCollection
    If includeDeclarations = True Then output.Add declarationsCollection
    If includeComponentType = True Then output.Add componentTypeCollection
    If includeComponentName = True Then output.Add componentCollection
    Set getDeclarations = output
End Function

Function getWord(str As Variant, Optional delim As String _
                                , Optional afterWord As String _
                                 , Optional beforeWord As String _
                                  , Optional counter As Integer _
                                   , Optional outer As Boolean _
                                    , Optional includeWords As Boolean) As String
    Dim i As Long
    If afterWord = "" And beforeWord = "" And counter = 0 Then MsgBox ("Pass at least 1 parameter betweenn -AfterWord- , -BeforeWord- , -counter-"): Exit Function
    If TypeName(str) = "String" Then
        If delim <> "" Then
            str = Split(str, delim)
            If UBound(str) <> 0 Then
                If afterWord = "" And beforeWord = "" And counter <> 0 Then If counter - 1 <= UBound(str) Then getWord = str(counter - 1): Exit Function
                For i = LBound(str) To UBound(str)
                    If afterWord <> "" And beforeWord = "" Then If i <> 0 Then If str(i - 1) = afterWord Then getWord = str(i): Exit Function
                    If afterWord = "" And beforeWord <> "" Then If i <> UBound(str) Then If str(i + 1) = beforeWord Then getWord = str(i): Exit Function
                    If afterWord <> "" And beforeWord <> "" Then If i <> 0 And i <> UBound(str) Then If str(i - 1) = afterWord And str(i + 1) = beforeWord Then getWord = str(i): Exit Function
                Next i
            End If
        Else
            If InStr(1, str, afterWord) > 0 And InStr(1, str, beforeWord) > 0 Then
                If includeWords = False Then
                    getWord = Mid(str, InStr(1, str, afterWord) + Len(afterWord))
                Else
                    getWord = Mid(str, InStr(1, str, afterWord))
                End If
                If outer = True Then
                    If includeWords = False Then
                        getWord = left(getWord, InStrRev(getWord, beforeWord) - 1)
                    Else
                        getWord = left(getWord, InStrRev(getWord, beforeWord) + Len(beforeWord) - 1)
                    End If
                Else
                    If includeWords = False Then
                        getWord = left(getWord, InStr(1, getWord, beforeWord) - 1)
                    Else
                        getWord = left(getWord, InStr(1, getWord, beforeWord) + Len(beforeWord) - 1)
                    End If
                End If
                Exit Function
            End If
        End If
    Else
    End If
    getWord = vbNullString
End Function


