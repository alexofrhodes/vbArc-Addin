Attribute VB_Name = "F_Arrays"
Rem @Folder ARRAY
Function ArrayMultiFilter(SourceArray As Variant, Matches As Variant, Optional Include As Boolean, Optional CompareMode As VbCompareMethod = vbBinaryCompare) As Variant
    Dim X&, arr, sJoined$
    For X = LBound(Matches) To UBound(Matches)
        arr = VBA.Filter(SourceArray, Matches(X), Include, CompareMode)
        sJoined = sJoined & VBA.Join(arr, ",") & ","
    Next X
    sJoined = left(sJoined, Len(sJoined) - 1)
    ArrayMultiFilter = Split(sJoined, ",")
End Function

Public Function ArrayRemoveEmptyElemets(varArray As Variant) As Variant
    Dim tempArray() As Variant
    Dim oldIndex As Integer
    Dim newIndex As Integer
    ReDim tempArray(LBound(varArray) To UBound(varArray))
    For oldIndex = LBound(varArray) To UBound(varArray)
        If Not Trim(varArray(oldIndex) & " ") = "" Then
            tempArray(newIndex) = varArray(oldIndex)
            newIndex = newIndex + 1
        End If
    Next oldIndex
    ReDim Preserve tempArray(LBound(varArray) To newIndex - 1)
    ArrayRemoveEmptyElemets = tempArray
End Function

Sub ArrayToRange1d(arr As Variant, Optional rng As Range)
    '#INCLUDE GetInputRange
    If rng Is Nothing Then
        If GetInputRange(rng, "select range", "select range") = False Then Exit Sub
    End If
    Dim dif As Long, difC As Long
    dif = IIf(LBound(arr, 1) = "0", 0, 1)
    rng.RESIZE(UBound(arr, 1) + dif) = arr
    rng.TextToColumns rng, , , , , , True
End Sub

Public Function SortArray( _
       ByVal sortableArray As Variant, _
       Optional ByVal descendingFlag As Boolean) _
        As Variant
    Dim i As Integer
    Dim swapOccuredBool As Boolean
    Dim arrayLength As Integer
    arrayLength = UBound(sortableArray) - LBound(sortableArray) + 1
    Dim sortedArray() As Variant
    ReDim sortedArray(arrayLength)
    Dim dif As Long
    dif = IIf(LBound(sortableArray) = 1, 1, 0)
    For i = 0 To arrayLength - 1
        sortedArray(i) = sortableArray(i + dif)
    Next
    Dim temporaryValue As Variant
    Do
        swapOccuredBool = False
        For i = 0 To arrayLength - 1
            If (sortedArray(i)) < sortedArray(i + 1) Then
                temporaryValue = sortedArray(i)
                sortedArray(i) = sortedArray(i + 1)
                sortedArray(i + 1) = temporaryValue
                swapOccuredBool = True
            End If
        Next
    Loop While swapOccuredBool
    If descendingFlag = True Then
        SortArray = sortedArray
    Else
        Dim ascendingArray() As Variant
        ReDim ascendingArray(arrayLength)
        For i = 0 To arrayLength - 1
            ascendingArray(i) = sortedArray(arrayLength - i - 1)
        Next
        SortArray = ascendingArray
    End If
End Function

Public Function DuplicatesInArray(ArrayOfValues) As String
    On Error GoTo Err_DuplicatesInArray
    Dim intUB As Integer
    Dim intElem As Integer
    Dim intLoop As Integer
    Dim intCount As Integer
    Dim varValue
    Dim varLoop
    Dim strResults As String
    intUB = UBound(ArrayOfValues)
    strResults = ""
    For intElem = 0 To intUB
        intCount = 0
        varValue = ArrayOfValues(intElem)
        If Not IsNull(varValue) Then
            For intLoop = 0 To intUB
                varLoop = ArrayOfValues(intLoop)
                If Not IsNull(varLoop) Then
                    If varLoop = varValue Then
                        intCount = intCount + 1
                    End If
                End If
            Next intLoop
            If intCount > 1 Then
                If InStr(strResults, varValue & ", ") = 0 Then
                    strResults = strResults & varValue & ", "
                End If
            End If
        End If
    Next intElem
    If Len(strResults) > 0 Then
        DuplicatesInArray = left(strResults, Len(strResults) - 2)
    Else
        DuplicatesInArray = ""
    End If
Exit_DuplicatesInArray:
    On Error Resume Next
    Exit Function
Err_DuplicatesInArray:
    MsgBox err.Number & " " & err.Description, vbCritical, "DuplicatesInArray()"
    DuplicatesInArray = ""
    Resume Exit_DuplicatesInArray
End Function

Public Sub SampleSortByLength()
    '#INCLUDE CustomQuickSort
    Dim sample As Variant
    sample = Split("this,is,a,random,phrase", ",")
    CustomQuickSort sample, LBound(sample), UBound(sample)
    Dim i As Integer
    For i = LBound(sample) To UBound(sample)
        Debug.Print sample(i)
    Next i
End Sub

Private Function SortByLengthCompare(one As Variant, two As Variant) As Boolean
    Select Case True
        Case Len(one) < Len(two)
            SortByLengthCompare = True
        Case Len(one) > Len(two)
            SortByLengthCompare = False
        Case Len(one) = Len(two)
            SortByLengthCompare = LCase$(one) < LCase$(two)
    End Select
End Function

Public Sub CustomQuickSort(list As Variant, first As Long, last As Long)
    '#INCLUDE SortByLengthCompare
    Dim pivot As String
    Dim low As Long
    Dim high As Long
    low = first
    high = last
    pivot = list((first + last) \ 2)
    Do While low <= high
        Do While low < last And SortByLengthCompare(list(low), pivot)
            low = low + 1
        Loop
        Do While high > first And SortByLengthCompare(pivot, list(high))
            high = high - 1
        Loop
        If low <= high Then
            Dim swap As String
            swap = list(low)
            list(low) = list(high)
            list(high) = swap
            low = low + 1
            high = high - 1
        End If
    Loop
    If (first < high) Then CustomQuickSort list, first, high
    If (low < last) Then CustomQuickSort list, low, last
End Sub

Function Transpose2DArray(inputArray As Variant) As Variant
    Dim X As Long, yUbound As Long
    Dim Y As Long, xUbound As Long
    Dim tempArray As Variant
    xUbound = UBound(inputArray, 2)
    yUbound = UBound(inputArray, 1)
    ReDim tempArray(1 To xUbound, 1 To yUbound)
    For X = 1 To xUbound
        For Y = 1 To yUbound
            tempArray(X, Y) = inputArray(Y, X)
        Next Y
    Next X
    Transpose2DArray = tempArray
End Function

Public Function ArrayDimensionLength(SourceArray As Variant) As Integer
    Dim i As Integer
    Dim test As Long
    On Error GoTo catch
    Do
        i = i + 1
        test = UBound(SourceArray, i)
    Loop
catch:
    ArrayDimensionLength = i - 1
End Function

Public Function IsArrayAllocated(ByRef arr As Variant) As Boolean
    On Error Resume Next
    IsArrayAllocated = IsArray(arr) And (Not IsError(LBound(arr, 1))) And LBound(arr, 1) <= UBound(arr, 1)
End Function

Rem @AUTHOR ROBERT TODAR
Public Function ArrayToString(SourceArray As Variant, Optional Delimiter As String = ",") As String
    '#INCLUDE ArrayDimensionLength
    Dim temp As String
    Select Case ArrayDimensionLength(SourceArray)
        Case 1
            temp = Join(SourceArray, Delimiter)
        Case 2
            Dim RowIndex As Long
            Dim ColIndex As Long
            For RowIndex = LBound(SourceArray, 1) To UBound(SourceArray, 1)
                For ColIndex = LBound(SourceArray, 2) To UBound(SourceArray, 2)
                    temp = temp & SourceArray(RowIndex, ColIndex)
                    If ColIndex <> UBound(SourceArray, 2) Then temp = temp & Delimiter
                Next ColIndex
                If RowIndex <> UBound(SourceArray, 1) Then temp = temp & vbNewLine
            Next RowIndex
    End Select
    ArrayToString = temp
End Function

Public Sub ArrayToRange2D(arr2d As Variant, rng As Range)
    rng.RESIZE(UBound(arr2d, 1), UBound(arr2d, 2)) = arr2d
End Sub

Function Filter2DArray(ByVal sArray, ByVal ColIndex As Long, ByVal FindStr As String, ByVal HasTitle As Boolean)
    Dim tmpArr, i As Long, j As Long, arr, dic, TmpStr, tmp, Chk As Boolean, TmpVal As Double
    On Error Resume Next
    Set dic = CreateObject("Scripting.Dictionary")
    tmpArr = sArray
    ColIndex = ColIndex + LBound(tmpArr, 2) - 1
    Chk = (InStr("><=", left(FindStr, 1)) > 0)
    For i = LBound(tmpArr, 1) - HasTitle To UBound(tmpArr, 1)
        If Chk Then
            TmpVal = CDbl(tmpArr(i, ColIndex))
            If Evaluate(TmpVal & FindStr) Then dic.Add i, ""
        Else
            If UCase(tmpArr(i, ColIndex)) Like UCase(FindStr) Then dic.Add i, ""
        End If
    Next
    If dic.count > 0 Then
        tmp = dic.keys
        ReDim arr(LBound(tmpArr, 1) To UBound(tmp) + LBound(tmpArr, 1) - HasTitle, LBound(tmpArr, 2) To UBound(tmpArr, 2))
        For i = LBound(tmpArr, 1) - HasTitle To UBound(tmp) + LBound(tmpArr, 1) - HasTitle
            For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
                arr(i, j) = tmpArr(tmp(i - LBound(tmpArr, 1) + HasTitle), j)
            Next
        Next
        If HasTitle Then
            For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
                arr(LBound(tmpArr, 1), j) = tmpArr(LBound(tmpArr, 1), j)
            Next
        End If
    End If
    Filter2DArray = arr
End Function

Function RotateArray(inputArray, Optional ShiftByNum = 1) As Variant
    Rem @TODO - Rotate right
    Rem rotates array left
    '#INCLUDE Len2
    Dim ub As Long: ub = UBound(inputArray)
    Dim LB As Long: LB = LBound(inputArray)
    Dim dif As Long: dif = 1 - LB
    Dim NewArray() As Variant
    Dim element As Variant
    Dim counter As Long
    Dim fromStart As Long: fromStart = LB
    For counter = LB To ub
        ReDim Preserve NewArray(1 To counter + dif)
        If counter + ShiftByNum <= ub Then
            NewArray(UBound(NewArray)) = inputArray(counter + ShiftByNum)
        ElseIf UBound(NewArray) <= Len2(inputArray) Then
            NewArray(UBound(NewArray)) = inputArray(fromStart)
            fromStart = fromStart + 1
        End If
    Next
    RotateArray = NewArray
End Function

Public Function Len2( _
       ByVal val As Variant) _
        As Integer
    If IsArray(val) And Right(TypeName(val), 2) = "()" Then
        Len2 = UBound(val) - LBound(val) + 1
    ElseIf TypeName(val) = "String" Then
        Len2 = Len(val)
    ElseIf IsNumeric(val) Then
        Len2 = Len(CStr(val))
    Else
        Len2 = val.count
    End If
End Function

Public Function IsInArray( _
       ByVal value1 As Variant, _
       ByVal array1 As Variant, _
       Optional CaseSensitive As Boolean) _
        As Boolean
    Dim individualElement As Variant
    If CaseSensitive = True Then value1 = UCase(value1)
    For Each individualElement In array1
        If CaseSensitive = True Then individualElement = UCase(individualElement)
        If individualElement = value1 Then
            IsInArray = True
            Exit Function
        End If
    Next
    IsInArray = False
End Function

Function joinArrays(arr1 As Variant, arr2 As Variant) As Variant
    Dim arrToReturn() As Variant, myCollection As New Collection
    For Each X In arr1: myCollection.Add X: Next
    For Each Y In arr2: myCollection.Add Y: Next
    If myCollection.count = 0 Then
        joinArrays = Array("")
        Exit Function
    End If
    ReDim arrToReturn(1 To myCollection.count)
    For i = 1 To myCollection.count: arrToReturn(i) = myCollection.item(i): Next
    joinArrays = arrToReturn
End Function

Function appendArray(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    Dim holdarr As Variant
    Dim ub1 As Long
    Dim ub2 As Long
    Dim i As Long
    Dim newind As Long
    If IsEmpty(arr1) Or Not IsArray(arr1) Then
        arr1 = Array()
    End If
    If IsEmpty(arr2) Or Not IsArray(arr2) Then
        arr2 = Array()
    End If
    ub1 = UBound(arr1)
    ub2 = UBound(arr2)
    If ub1 = -1 Then
        appendArray = arr2
        Exit Function
    End If
    If ub2 = -1 Then
        appendArray = arr1
        Exit Function
    End If
    holdarr = arr1
    ReDim Preserve holdarr(ub1 + ub2 + 1)
    newind = UBound(arr1) + 1
    For i = 0 To ub2
        If VarType(arr2(i)) = vbObject Then
            Set holdarr(newind) = arr2(i)
        Else
            holdarr(newind) = arr2(i)
        End If
        newind = newind + 1
    Next i
    appendArray = holdarr
End Function

Public Function Combine2Array(ByVal arr1 As Variant, ByVal arr2 As Variant) As Variant
    '#INCLUDE NumberOfArrayDimensions
    Dim LowRowArr1 As Long
    Dim HighRowArr1 As Long
    Dim LowColumnArr1 As Long
    Dim HighColumnArr1 As Long
    Dim NumOfRowsArr1 As Long
    Dim NumOfColumnsArr1 As Long
    Dim LowRowArr2 As Long
    Dim HighRowArr2 As Long
    Dim LowColumnArr2 As Long
    Dim HighColumnArr2 As Long
    Dim NumOfRowsArr2 As Long
    Dim NumOfColumnsArr2 As Long
    Dim output As Variant
    Dim OutputRow As Long
    Dim OutputColumn As Long
    Dim RowIdx As Long
    Dim ColIdx As Long
    If (IsArray(arr1) = False) Or (IsArray(arr2) = False) Then
        Combine2Array = Null
        MsgBox "Both need to be array"
        Exit Function
    End If
    If (NumberOfArrayDimensions(arr1) <> 2) Or (NumberOfArrayDimensions(arr2) <> 2) Then
        Combine2Array = Null
        MsgBox "Both need to be 2D array"
        Exit Function
    End If
    LowRowArr1 = LBound(arr1, 1)
    HighRowArr1 = UBound(arr1, 1)
    LowColumnArr1 = LBound(arr1, 2)
    HighColumnArr1 = UBound(arr1, 2)
    NumOfRowsArr1 = HighRowArr1 - LowRowArr1 + 1
    NumOfColumnsArr1 = HighColumnArr1 - LowColumnArr1 + 1
    LowRowArr2 = LBound(arr2, 1)
    HighRowArr2 = UBound(arr2, 1)
    LowColumnArr2 = LBound(arr2, 2)
    HighColumnArr2 = UBound(arr2, 2)
    NumOfRowsArr2 = HighRowArr2 - LowRowArr2 + 1
    NumOfColumnsArr2 = HighColumnArr2 - LowColumnArr2 + 1
    If NumOfColumnsArr1 <> NumOfColumnsArr2 Then
        Combine2Array = Null
        MsgBox "Both array must have same number of column"
        Exit Function
    End If
    ReDim output(0 To NumOfRowsArr1 + NumOfRowsArr2 - 1, 0 To NumOfColumnsArr1 - 1)
    For RowIdx = LowRowArr1 To HighRowArr1
        OutputColumn = 0
        For ColIdx = LowColumnArr1 To HighColumnArr1
            output(OutputRow, OutputColumn) = arr1(RowIdx, ColIdx)
            OutputColumn = OutputColumn + 1
        Next ColIdx
        OutputRow = OutputRow + 1
    Next RowIdx
    For RowIdx = LowRowArr2 To HighRowArr2
        OutputColumn = 0
        For ColIdx = LowColumnArr2 To HighColumnArr2
            output(OutputRow, OutputColumn) = arr2(RowIdx, ColIdx)
            OutputColumn = OutputColumn + 1
        Next ColIdx
        OutputRow = OutputRow + 1
    Next RowIdx
    Combine2Array = output
End Function

Public Function NumberOfArrayDimensions(arr As Variant) As Byte
    Dim Ndx As Byte
    Dim Res As Long
    On Error Resume Next
    Do
        Ndx = Ndx + 1
        Res = UBound(arr, Ndx)
    Loop Until err.Number <> 0
    NumberOfArrayDimensions = Ndx - 1
End Function


