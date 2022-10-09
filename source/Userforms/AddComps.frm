VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} AddComps 
   Caption         =   "Add Components"
   ClientHeight    =   5412
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   9096.001
   OleObjectBlob   =   "AddComps.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "AddComps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : AddComps
'* Created    : 06-10-2022 10:33
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub CommandButton1_Click()
    If pmWorkbook Is Nothing Then Set pmWorkbook = ActiveWorkbook
    Dim coll As Collection
    Set coll = New Collection
    Dim element As Variant
    coll.Add Split(Me.tModule.TEXT, vbNewLine)
    coll.Add Split(Me.tClass.TEXT, vbNewLine)
    coll.Add Split(Me.tUserform.TEXT, vbNewLine)
    coll.Add Split(Me.tDocument.TEXT, vbNewLine)
    Dim typeCounter As Long
    For Each element In coll
        If UBound(element) <> -1 Then
            typeCounter = typeCounter + 1
            AddComponent pmWorkbook, typeCounter, element
        End If
    Next element
    MsgBox typeCounter & " components added to " & pmWorkbook.Name
End Sub

Sub AddComponent(wb As Workbook, Module_Class_Form_Sheet As Long, componentArray As Variant)
    '#INCLUDE ModuleExists
    Dim compType As Long
    compType = Module_Class_Form_Sheet
    Dim vbProj As VBProject
    Set vbProj = wb.VBProject
    Dim vbComp As VBComponent
    Dim NewSheet As Worksheet
    Dim i As Long
    Dim counter As Long
    On Error GoTo ErrorHandler
    For i = LBound(componentArray) To UBound(componentArray)
        If componentArray(i) <> vbNullString Then
            Select Case compType
                Case Is = 1, 2, 3
                    If ModuleExists(CStr(componentArray(i))) = False Then
                        If compType = 1 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_StdModule)
                        If compType = 2 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_ClassModule)
                        If compType = 3 Then Set vbComp = vbProj.VBComponents.Add(vbext_ct_MSForm)
                    End If
                    vbComp.Name = componentArray(i)
                Case Is = 4
                    If compType = 4 Then
                        Set NewSheet = CreateOrSetSheet(CStr(componentArray(i)), wb)
                        NewSheet.Name = componentArray(i)
                    End If
            End Select
        End If
loop1:
    Next i
    On Error GoTo 0
    Exit Sub
ErrorHandler:
    counter = counter + 1
    componentArray(i) = componentArray(i) & counter
    GoTo loop1
End Sub

