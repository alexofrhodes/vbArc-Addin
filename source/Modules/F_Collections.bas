Attribute VB_Name = "F_Collections"
Rem @Folder Collection
Function collectionToString(coll As Collection, delim As String) As String
    Dim element
    Dim out As String
    For Each element In coll
        out = IIf(out = "", element, out & delim & element)
    Next
    collectionToString = out
End Function

Public Function CollectionOfUnique( _
       FullCollection As Collection) _
        As Variant
    Dim UniqueCollection As Collection: Set UniqueCollection = New Collection
    Dim eachItem As Variant
    On Error Resume Next
    For Each eachItem In FullCollection
        UniqueCollection.Add eachItem, CStr(eachItem)
    Next
    On Error GoTo 0
    Set CollectionOfUnique = UniqueCollection
End Function

Public Function SortCollection(colInput As Collection) As Collection
    Dim iCounter As Integer
    Dim iCounter2 As Integer
    Dim temp As Variant
    Set SortCollection = New Collection
    For iCounter = 1 To colInput.count - 1
        For iCounter2 = iCounter + 1 To colInput.count
            If colInput(iCounter) > colInput(iCounter2) Then
                temp = colInput(iCounter2)
                colInput.Remove iCounter2
                colInput.Add temp, , iCounter
            End If
        Next iCounter2
    Next iCounter
    Set SortCollection = colInput
End Function

Public Function CollectionContains(Kollection As Collection, Optional key As Variant, Optional item As Variant) As Boolean
    Dim strKey As String
    Dim var As Variant
    If Not IsMissing(key) Then
        strKey = CStr(key)
        On Error Resume Next
        CollectionContains = True
        var = Kollection(strKey)
        If err.Number = 91 Then GoTo CheckForObject
        If err.Number = 5 Then GoTo NotFound
        On Error GoTo 0
        Exit Function
CheckForObject:
        If IsObject(Kollection(strKey)) Then
            CollectionContains = True
            On Error GoTo 0
            Exit Function
        End If
NotFound:
        CollectionContains = False
        On Error GoTo 0
        Exit Function
    ElseIf Not IsMissing(item) Then
        CollectionContains = False
        For Each var In Kollection
            If var = item Then
                CollectionContains = True
                Exit Function
            End If
        Next var
    Else
        CollectionContains = False
    End If
End Function

Function CollectionToArray(c As Collection) As Variant
    Dim a() As Variant: ReDim a(0 To c.count - 1)
    Dim i As Long
    For i = 1 To c.count
        a(i - 1) = c.item(i)
    Next
    CollectionToArray = a
End Function

Function CollectionsToArrayTable(collections As Collection) As Variant
    Dim columnCount As Long
    columnCount = collections.count
    Dim rowCount As Long
    rowCount = collections.item(1).count
    Dim var As Variant
    ReDim var(1 To rowCount, 1 To columnCount)
    Dim cols As Long
    Dim rows As Long
    For rows = 1 To rowCount
        For cols = 1 To collections.count
            var(rows, cols) = collections(cols).item(rows)
        Next cols
    Next rows
    CollectionsToArrayTable = var
End Function


