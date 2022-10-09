Attribute VB_Name = "F_References"
Option Explicit
Rem @Folder ReferencesUserform Declarations
Private Const HKEY_CLASSES_ROOT As Long = &H80000000
Private Const READ_CONTROL As Long = &H20000
Private Const STANDARD_RIGHTS_READ As Long = (READ_CONTROL)
Private Const KEY_QUERY_VALUE As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY As Long = &H10
Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_READ As Long = (( _
STANDARD_RIGHTS_READ _
Or KEY_QUERY_VALUE _
Or KEY_ENUMERATE_SUB_KEYS _
Or KEY_NOTIFY) _
And (Not SYNCHRONIZE))
Private Const ERROR_SUCCESS As Long = 0&
Private Const ERROR_NO_MORE_ITEMS As Long = 259&
Private Declare PtrSafe Function RegOpenKeyEx _
Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
ByVal HKEY As Long, _
ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
ByRef phkResult As Long) As Long
Private Declare PtrSafe Function RegEnumKey _
Lib "advapi32.dll" Alias "RegEnumKeyA" ( _
ByVal HKEY As Long, _
ByVal dwIndex As Long, _
ByVal lpName As String, _
ByVal cbName As Long) As Long
Private Declare PtrSafe Function RegQueryValue _
Lib "advapi32.dll" Alias "RegQueryValueA" ( _
ByVal HKEY As Long, _
ByVal lpSubKey As String, _
ByVal lpValue As String, _
ByRef lpcbValue As Long) As Long
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal HKEY As Long) As Long

Rem @Folder ReferencesUserform
Private Sub RefList()
    '#INCLUDE CreateOrSetSheet
    Dim R1 As Long
    Dim R2 As Long
    Dim hHK1 As Long
    Dim hHK2 As Long
    Dim hHK3 As Long
    Dim hHK4 As Long
    Dim i As Long
    Dim i2 As Long
    Dim lpPath As String
    Dim lpGUID As String
    Dim lpName As String
    Dim lpValue As String
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("References", ThisWorkbook)
    ws.Cells.clear
    ws.Cells(1, 1).Value = "Reference Description"
    ws.Cells(1, 2).Value = "GUID"
    ws.Cells(1, 3).Value = "Path"
    ws.Cells(1, 4).Value = "Version"
    lpPath = String$(128, vbNullChar)
    lpValue = String$(128, vbNullChar)
    lpName = String$(128, vbNullChar)
    lpGUID = String$(128, vbNullChar)
    R1 = RegOpenKeyEx(HKEY_CLASSES_ROOT, "TypeLib", ByVal 0&, KEY_READ, hHK1)
    If R1 = ERROR_SUCCESS Then
        i = 1
        Do While Not R1 = ERROR_NO_MORE_ITEMS
            R1 = RegEnumKey(hHK1, i, lpGUID, Len(lpGUID))
            If R1 = ERROR_SUCCESS Then
                R2 = RegOpenKeyEx(hHK1, lpGUID, ByVal 0&, KEY_READ, hHK2)
                If R2 = ERROR_SUCCESS Then
                    i2 = 0
                    Do While Not R2 = ERROR_NO_MORE_ITEMS
                        R2 = RegEnumKey(hHK2, i2, lpName, Len(lpName))
                        If R2 = ERROR_SUCCESS Then
                            RegQueryValue hHK2, lpName, lpValue, Len(lpValue)
                            RegOpenKeyEx hHK2, lpName, ByVal 0&, KEY_READ, hHK3
                            RegOpenKeyEx hHK3, "0", ByVal 0&, KEY_READ, hHK4
                            RegQueryValue hHK4, "win32", lpPath, Len(lpPath)
                            i2 = i2 + 1
                            ws.Cells(i + 1, 1) = lpValue
                            ws.Cells(i + 1, 2) = lpGUID
                            ws.Cells(i + 1, 3) = lpPath
                            ws.Cells(i + 1, 4) = lpName
                        End If
                    Loop
                End If
            End If
            i = i + 1
        Loop
        RegCloseKey hHK1
        RegCloseKey hHK2
        RegCloseKey hHK3
        RegCloseKey hHK4
    End If
    Rem   ws.Columns("A:A").EntireColumn.AutoFit
    Rem   ws.Columns("B:C").ColumnWidth = 70
    ws.Range("A1").CurrentRegion.Sort Key1:=Range("A1"), header:=xlYes
End Sub

Sub ListAllReferences()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("References")
    ws.Cells(1, 1).Value = "Reference Description"
    ws.Cells(1, 2).Value = "GUID"
    ws.Cells(1, 3).Value = "Path"
    ws.Cells(1, 4).Value = "Version"
    Dim myRef As Reference
    Dim refs As VBIDE.REFERENCES
    Set refs = Application.VBE.ActiveVBProject.REFERENCES
    Dim i As Long
    i = 2
    For Each myRef In refs
        ws.Cells(i, 1) = IIf(myRef.Description <> "", myRef.Description, myRef.Name)
        ws.Cells(i, 2) = myRef.GUID
        ws.Cells(i, 3) = myRef.fullPath
        ws.Cells(i, 4) = myRef.Major & "." & myRef.Minor
        i = i + 1
    Next myRef
End Sub

Sub xlfVBEAddReferences()
    Dim oRefs As REFERENCES
    Set oRefs = Application.VBE.ActiveVBProject.REFERENCES
    On Error GoTo OnError
    oRefs.AddFromFile "C:\Windows\System32\msxml6.dll"
OnError:
End Sub

Sub xlfVBEAddReferencesGUID()
    Dim oRefs As REFERENCES
    Set oRefs = Application.VBE.ActiveVBProject.REFERENCES
    On Error GoTo OnError
    Rem Syntax: AddFromGuid(Guid, Major, Minor)
    oRefs.AddFromGuid "{F5078F18-C551-11D3-89B9-0000F81FE221}", 6, 0
OnError:
End Sub

Sub xlfVBERemoveReference1()
    Dim oRef As Reference
    Dim oRefs As REFERENCES
    Set oRefs = Application.VBE.ActiveVBProject.REFERENCES
    For Each oRef In oRefs
        If oRef.Name = "MSXML2" Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub xlfVBERemoveReference2()
    Dim oRef As Reference
    Dim oRefs As REFERENCES
    Set oRefs = Application.VBE.ActiveVBProject.REFERENCES
    For Each oRef In oRefs
        If oRef.Description = "Microsoft XML, v6.0" Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub

Sub RemoveReferenceByGUID(TargetWorkbook As Workbook, refGUID As String)
    '#INCLUDE dp
    Dim oRefs As REFERENCES
    Set oRefs = TargetWorkbook.VBProject.REFERENCES
    Dim oRef As Reference
    For Each oRef In oRefs
        dp oRef.Name
        If oRef.GUID = refGUID Then
            oRefs.Remove oRef
            Exit For
        End If
    Next oRef
End Sub


