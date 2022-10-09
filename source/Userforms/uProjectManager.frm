VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uProjectManager 
   Caption         =   "github.com/AlexOfRhodes"
   ClientHeight    =   5508
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   3840
   OleObjectBlob   =   "uProjectManager.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uProjectManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : uProjectManager
'* Created    : 06-10-2022 10:39
'* Author     : Anastasiou Alex
'* Contacts   : AnastasiouAlex@gmail.com
'*
'* GITHUB     : https://github.com/AlexOfRhodes
'* YOUTUBE    : https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg
'* VK         : https://vk.com/video/playlist/735281600_1
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Private Sub goToFolder_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    FollowLink Environ("USERprofile") & "\Documents\vbArc\"
End Sub

Private Sub iAdd_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oAdd.Value = True
    optionsBlank
    iAdd.BorderStyle = fmBorderStyleSingle
    ExportOptionsHide
End Sub

Private Sub iExport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oExport.Value = True
    optionsBlank
    iExport.BorderStyle = fmBorderStyleSingle
    ExportOptionsShow
End Sub

Private Sub iImport_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oImport.Value = True
    optionsBlank
    iImport.BorderStyle = fmBorderStyleSingle
    ExportOptionsHide
End Sub

Private Sub iRename_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oRename.Value = True
    optionsBlank
    iRename.BorderStyle = fmBorderStyleSingle
    ExportOptionsHide
End Sub

Private Sub iRemove_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oDelete.Value = True
    optionsBlank
    iRemove.BorderStyle = fmBorderStyleSingle
End Sub

Private Sub iRefresh_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    oRefresh.Value = True
    optionsBlank
    iRefresh.BorderStyle = fmBorderStyleSingle
    ExportOptionsHide
End Sub

Sub optionsBlank()
    iAdd.BorderStyle = fmBorderStyleNone
    iExport.BorderStyle = fmBorderStyleNone
    iImport.BorderStyle = fmBorderStyleNone
    iRename.BorderStyle = fmBorderStyleNone
    iRefresh.BorderStyle = fmBorderStyleNone
    iRemove.BorderStyle = fmBorderStyleNone
End Sub

Sub ExportOptionsHide()
    ExportOptions.visible = False
    ImageExport.visible = False
    '    FrameBottom.Top = 36
    '    Me.Height = 198
End Sub

Sub ExportOptionsShow()
    '    FrameBottom.top = 132
    ExportOptions.visible = True
    ImageExport.visible = True
    '    Me.Height = 295
End Sub

Private Sub SelectFromList_Click()
    If listOpenBooks.ListIndex = -1 Then
        MsgBox "No book selected"
        Exit Sub
    End If
    Set pmWorkbook = Workbooks(listOpenBooks.list(listOpenBooks.ListIndex))
    SelectAction
End Sub

Private Sub UserForm_Initialize()
    LoadBooksAndAddins
    SortListboxOnColumn Me.listOpenBooks, 0
    LoadUserformOptions Me
    FormatColourFormatters
End Sub

Sub LoadBooksAndAddins()
    Rem list workbooks
    '#INCLUDE ListboxContains
    '#INCLUDE ProtectedVBProject
    Dim wb As Workbook
    For Each wb In Workbooks
        If Len(wb.Path) > 0 Then
            If ProtectedVBProject(wb) = False Then listOpenBooks.AddItem wb.Name
        End If
    Next
    Rem list addins
    Dim vbProj As VBProject
    Dim wbPath As String
    For Each vbProj In Application.VBE.VBProjects
        On Error GoTo ErrorHandler
        wbPath = vbProj.fileName
        If Right(wbPath, 4) = "xlam" Or Right(wbPath, 3) = "xla" Then
            Dim wbName As String
            wbName = Mid(wbPath, InStrRev(wbPath, "\") + 1)
            If ProtectedVBProject(Workbooks(wbName)) = False Then
                If ListboxContains(listOpenBooks, wbName) = False Then listOpenBooks.AddItem wbName
            End If
        End If
Skip:
    Next vbProj
    Exit Sub
ErrorHandler:
    If err.Number = 76 Then GoTo Skip
End Sub

Private Sub ActiveFile_Click()
    Set pmWorkbook = ActiveWorkbook
    SelectAction
End Sub

Sub SelectAction()
    '#INCLUDE ImportComponents
    '#INCLUDE ExportProject
    '#INCLUDE RefreshComponents
    '#INCLUDE ProtectedVBProject
    '#INCLUDE ListAllProcedureImports
    
    If ProtectedVBProject(pmWorkbook) = True Then
        MsgBox "Project of " & pmWorkbook.Name & " is protected."
        Exit Sub
    End If
    
    '    ListAllProcedureImports pmWorkbook
    
    Select Case True
        Case oExport.Value = True
            Me.Hide
            ExportProject wb:=pmWorkbook _
                               , bExportComponents:=chExportComponents.Value _
                                                     , bSeparateProcedures:=chExportProcedures.Value _
                                                                             , PrintCode:=chPrintCode.Value _
                                                                                           , ExportSheets:=chExportSheets.Value _
                                                                                                            , ExportForms:=chExportForms.Value _
                                                                                                                            , bExportXML:=chExportXML.Value _
                                                                                                                                           , bExportReferences:=chExportReferences.Value _
                                                                                                                                                                 , bExportDeclarations:=chExportDeclarations.Value _
                                                                                                                                                                                         , bExportUnified:=chExportUnified.Value _
                                                                                                                                                                                                            , bWorkbookBackup:=chWorkbookBackup.Value
                        
            Me.Show
        Case oImport.Value = True
            ImportComponents pmWorkbook
        Case oRefresh.Value = True
            RefreshComponents pmWorkbook
        Case oDelete.Value = True
            RemoveComps.Show
        Case oRename.Value = True
            RenameComps.Show
        Case oAdd.Value = True
            AddComps.Show
        Case Else
    End Select
End Sub

Private Sub cInfo_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    uDEV.Show
End Sub

Private Sub Image1_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ExportSettings.Show
End Sub

Private Sub SelectFile_Click()
    Dim fPath As String
    fPath = PickExcelFile
    If fPath = "" Then Exit Sub
    Set pmWorkbook = Workbooks.Open(fileName:=fPath, UpdateLinks:=0, ReadOnly:=False)
    SelectAction
    Set pmWorkbook = Nothing
End Sub

Private Sub LBLcolourCode_Click()
    ColorPaletteDialog ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J1"), LBLcolourCode
End Sub

Private Sub LBLcolourComment_Click()
    ColorPaletteDialog ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J4"), LBLcolourComment
End Sub

Private Sub LBLcolourKey_Click()
    ColorPaletteDialog ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J3"), LBLcolourKey
End Sub

Private Sub LBLcolourOdd_Click()
    ColorPaletteDialog ThisWorkbook.SHEETS("ProjectManagerTXTColour").Range("J2"), LBLcolourOdd
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    SaveUserformOptions Me, , False
End Sub


