Attribute VB_Name = "U_Finder"

Rem @Folder Finder
Option Compare Text
Public Enum OPERATOR
    IS_LIKE
    IS_EQUAL
    NOT_EQUAL
    IS_CONTAINS
    NOT_CONTAINS
    STARTS_WITH
    ENDS_WITH
    GREATER_THAN
    GREATER_OR_EQUAL
    LESS_THAN
    LESS_OR_EQUAL
    IS_BETWEEN
    NOT_BETWEEN
End Enum

Function FindDateRange(rng As Range) As Range
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
    Next cell
    Set FindDateRange = CopyRange
End Function

Function FindNumericRange(rng As Range) As Range
    Dim StartTime
    StartTime = Now
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsNumeric(cell) And Not IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
        Rem        If Now() - startTime > TimeSerial(0, 0, 10) Then Stop
    Next cell
    Set FindNumericRange = CopyRange
End Function

Function FindStringRange(rng As Range) As Range
    Dim CopyRange As Range
    Dim cell As Range
    For Each cell In rng
        If IsDate(cell) Then
            If CopyRange Is Nothing Then
                Set CopyRange = cell
            Else
                Set CopyRange = Union(CopyRange, cell)
            End If
        End If
    Next cell
    Set FindStringRange = CopyRange
End Function

Sub ListboxToRange(lBox As MSForms.ListBox, rng As Range)
    rng.RESIZE(lBox.ListCount, lBox.columnCount) = lBox.list
End Sub

Function ArrayColumn(arr As Variant, col As Long) As Variant
    ArrayColumn = WorksheetFunction.index(arr, 0, col)
End Function

Private Function SortCompare(one As Variant, two As Variant) As Boolean
    Select Case True
        Case Len(one) < Len(two)
            SortCompare = True
        Case Len(one) > Len(two)
            SortCompare = False
        Case Len(one) = Len(two)
            SortCompare = LCase$(one) < LCase$(two)
    End Select
End Function

Function FindIfGetRow(FirstValue, operation As OPERATOR, SecondValue, offsetRow As Integer, offsetColumn As Integer, _
                      Optional wb As Workbook, Optional ws As Worksheet, Optional delim As String = ",") As Collection
    '#INCLUDE ArrayToString
    '#INCLUDE compare
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim output As New Collection
    Dim arr, element
    Dim str As String
    Dim c As Range
    Dim firstAddress As String
    If TypeName(ws) = "Nothing" Then
        For Each ws In wb.Worksheets
            With ws.Cells
                Set c = .Find(FirstValue, LookIn:=xlValues)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    Do
                        If compare(c.OFFSET(offsetRow, offsetColumn).Value, operation, SecondValue) = True Then
                            arr = ws.Range(ws.Cells(c.row, c.CurrentRegion.Column), ws.Cells(c.row, c.CurrentRegion.Column + c.CurrentRegion.Columns.count - 1)).Value
                            str = ArrayToString(arr, delim)
                            output.Add str
                            Debug.Print str
                        End If
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End With
        Next
    Else
        With ws.Cells
            Set c = .Find(FirstValue, LookIn:=xlValues)
            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    If compare(c.OFFSET(offsetRow, offsetColumn).Value, operation, SecondValue) = True Then
                        arr = ws.Range(ws.Cells(c.row, c.CurrentRegion.Column), ws.Cells(c.row, c.CurrentRegion.Column + c.CurrentRegion.Columns.count - 1)).Value
                        str = ArrayToString(arr, delim)
                        output.Add str
                        Debug.Print str
                    End If
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstAddress
            End If
        End With
    End If
    Set FindIfGetRow = output
End Function

Function FindAllGetRow(FirstValue, Optional wb As Workbook, Optional ws As Worksheet) As Collection
    '#INCLUDE dp
    '#INCLUDE ArrayToString
    If wb Is Nothing Then Set wb = ActiveWorkbook
    Dim output As New Collection
    Dim arr, element
    Dim str As String
    Dim c As Range
    Dim firstAddress As String
    If TypeName(ws) = "Nothing" Then
        For Each ws In wb.Worksheets
            With ws.Cells
                Set c = .Find(FirstValue, LookIn:=xlValues)
                If Not c Is Nothing Then
                    firstAddress = c.Address
                    Rem dp vbNewLine & ws.Name
                    Do
                        Rem dp c.Address
                        arr = ws.Range(ws.Cells(c.row, c.CurrentRegion.Column), ws.Cells(c.row, c.CurrentRegion.Column + c.CurrentRegion.Columns.count - 1)).Value
                        str = ArrayToString(arr)
                        Do While InStr(1, str, ",,") > 0
                            str = Replace(str, ",,", ",")
                        Loop
                        str = str
                        output.Add str
                        Debug.Print str
                        Set c = .FindNext(c)
                    Loop While Not c Is Nothing And c.Address <> firstAddress
                End If
            End With
        Next
    Else
        With ws.Cells
            Set c = .Find(FirstValue, LookIn:=xlValues)
            If Not c Is Nothing Then
                firstAddress = c.Address
                Do
                    arr = ws.Range(ws.Cells(c.row, c.CurrentRegion.Column), ws.Cells(c.row, c.CurrentRegion.Column + c.CurrentRegion.Columns.count - 1)).Value
                    str = ArrayToString(arr)
                    output.Add str
                    Debug.Print str
                    Set c = .FindNext(c)
                Loop While Not c Is Nothing And c.Address <> firstAddress
            End If
        End With
    End If
    Set FindAllGetRow = output
End Function

Function compare(inputValue, operation As OPERATOR, FirstComparison, Optional SecondComparison, Optional CaseSensitive As Boolean) As Boolean
    '#INCLUDE InStrExact
    If TypeName(inputValue) = "Range" Then inputValue = inputValue.Value
    Select Case TypeName(inputValue)
        Case "String()", "Variant", "Variant()", "Collection"
            MsgBox "Not able to proccess this case at the moment"
            Stop
    End Select
    If TypeName(inputValue) = "String" Then
        If CaseSensitive = True Then
            inputValue = UCase(inputValue)
            FirstComparison = UCase(FirstComparison)
            If Not IsMissing(SecondComparison) Then SecondComparison = UCase(SecondComparison)
        End If
    ElseIf IsDate(inputValue) Then
        inputValue = CDate(inputValue)
    ElseIf IsNumeric(inputValue) Then
        inputValue = CDbl(inputValue)
    End If
    If IsDate(FirstComparison) Then
        FirstComparison = CDate(FirstComparison)
        If Not IsMissing(SecondComparison) Then
            If IsDate(SecondComparison) Then SecondComparison = CDate(SecondComparison)
        End If
    ElseIf IsNumeric(FirstComparison) Then
        FirstComparison = CDbl(FirstComparison)
        If Not IsMissing(SecondComparison) Then
            If IsNumeric(SecondComparison) Then SecondComparison = CDbl(SecondComparison)
        End If
    End If
    If operation = OPERATOR.IS_LIKE Then
        If inputValue Like FirstComparison Then
            compare = True
        End If
    ElseIf operation = OPERATOR.IS_CONTAINS Then
        If InStrExact(1, CStr(inputValue), CStr(FirstComparison)) > 0 Then compare = True
    ElseIf operation = OPERATOR.NOT_CONTAINS Then
        If InStrExact(1, CStr(inputValue), CStr(FirstComparison)) = 0 Then compare = True
    ElseIf operation = OPERATOR.NOT_EQUAL Then
        If inputValue <> FirstComparison Then compare = True
    ElseIf operation = OPERATOR.STARTS_WITH Then
        If inputValue Like FirstComparison & "*" Then compare = True
    ElseIf operation = OPERATOR.ENDS_WITH Then
        If inputValue Like "*" & FirstComparison Then compare = True
    ElseIf operation = OPERATOR.IS_EQUAL Then
        If inputValue = FirstComparison Then compare = True
    ElseIf operation = OPERATOR.GREATER_THAN Then
        If inputValue > FirstComparison Then compare = True
    ElseIf operation = OPERATOR.GREATER_OR_EQUAL Then
        If inputValue >= FirstComparison Then compare = True
    ElseIf operation = OPERATOR.IS_BETWEEN Then
        If FirstComparison < inputValue And inputValue < SecondComparison Then compare = True
    ElseIf operation = OPERATOR.NOT_BETWEEN Then
        If Not (FirstComparison < inputValue And inputValue < SecondComparison) Then compare = True
    ElseIf operation = OPERATOR.LESS_THAN Then
        If inputValue < FirstComparison Then compare = True
    ElseIf operation = OPERATOR.LESS_OR_EQUAL Then
        If inputValue <= FirstComparison Then compare = True
    End If
End Function

Public Function RangeFindAll(ByRef SearchRange As Range, ByVal Value As Variant, Optional ByVal LookIn As XlFindLookIn = xlValues, Optional LookAt As XlLookAt = xlPart) As Range
    Dim FoundValues As Range, Found As Range, Prev As Range, Looper As Boolean: Looper = True
    Do While Looper
        If Not Prev Is Nothing Then Set Found = SearchRange.Find(Value, Prev, LookIn, LookAt)
        If Found Is Nothing Then Set Found = SearchRange.Find(Value, LookIn:=LookIn, LookAt:=LookAt)
        If Not Found Is Nothing Then
            If FoundValues Is Nothing Then
                Set FoundValues = Found
            Else
                If Not Intersect(Found, FoundValues) Is Nothing Then Looper = False
                Set FoundValues = Union(FoundValues, Found)
            End If
            Set Prev = Found
        End If
        If Found Is Nothing And Prev Is Nothing Then Exit Function
    Loop
    Set RangeFindAll = FoundValues
    Set FoundValues = Nothing
    Set Found = Nothing
    Set Prev = Nothing
End Function


