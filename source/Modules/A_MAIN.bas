Attribute VB_Name = "A_MAIN"

Rem @Folder Main
Public Const AUTHOR_GITHUB = "https://github.com/AlexOfRhodes"
Public Const AUTHOR_YOUTUBE = "https://www.youtube.com/channel/UC5QH3fn1zjx0aUjRER_rOjg"
Public Const AUTHOR_VK = "https://vk.com/video/playlist/735281600_1"
Public Const AUTHOR_NAME = "Anastasiou Alex"
Public Const AUTHOR_EMAIL = "AnastasiouAlex@gmail.com"
Public Const AUTHOR_COPYRIGHT = ""
Public Const AUTHOR_OTHERTEXT = ""
Public Const AUTHOR_MEDIA = "'* GITHUB     : " & AUTHOR_GITHUB & vbNewLine & _
"'* YOUTUBE    : " & AUTHOR_YOUTUBE & vbNewLine & _
"'* VK         : " & AUTHOR_VK & vbNewLine
Public Const PROJECT_URL = "https://github.com/alexofrhodes/vbArc-Addin/"
Public Const PROJECT_DOWNLAOD_URL = "https://github.com/alexofrhodes/vbArc-Addin/raw/main/vbArc-Addin.xlsm"
Public Const PROJECT_CHANGELOG_URL = "https://github.com/alexofrhodes/vbArc-Addin/raw/main/ChangeLog.md"
Public Const PROJECT_NAME = "vbArc-Addin"
Public Const PROJECT_VERSION_URL = "https://github.com/alexofrhodes/vbArc-Addin/raw/main/changelog.md"
Public ShowInVBE As Boolean
Public Const SNIP_FOLDER As String = "C:\Users\acer\My Documents\vbArc\SNIPPETS\"
Public myRibbon As IRibbonUI
#If VBA7 Then
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As LongPtr)
#Else
    Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
#End If

Function IMAGE_FOLDER() As String
    '#INCLUDE FoldersCreate
    Dim myPath As String
    myPath = ThisWorkbook.Path & "\Ribbon Images\"
    FoldersCreate myPath
    IMAGE_FOLDER = myPath
End Function

Sub vbArcRibbon_OnLoad(ribbon As IRibbonUI)
    #If VBA7 Then
        Dim StoreRibbonPointer As LongPtr
    #Else
        Dim StoreRibbonPointer As Long
    #End If
    Set myRibbon = ribbon
    StoreRibbonPointer = ObjPtr(ribbon)
    ThisWorkbook.SHEETS("vbArc_Addin_Settings").Range("B1").Value = StoreRibbonPointer
End Sub

Sub vbArcRibbon_RefreshRibbon()
    Rem PURPOSE: Refresh Ribbon UI
    '#INCLUDE GetRibbon
    Dim myRibbon As Object
    On Error GoTo RestartExcel
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(ThisWorkbook.SHEETS("vbArc_Addin_Settings").Range("B1").Value)
    End If
    Rem Redo Ribbon Load
    myRibbon.Invalidate
    On Error GoTo 0
    Exit Sub
RestartExcel:
    MsgBox "Ribbon UI Refresh Failed"
End Sub

Sub InvalidateControl(controlID)
    Rem PURPOSE: Refresh Ribbon UI
    '#INCLUDE GetRibbon
    On Error GoTo RestartExcel
    If myRibbon Is Nothing Then
        Set myRibbon = GetRibbon(ThisWorkbook.SHEETS("vbArc_Addin_Settings").Range("B1").Value)
    End If
    myRibbon.InvalidateControl controlID
    On Error GoTo 0
    Exit Sub
RestartExcel:
    MsgBox "Ribbon UI Refresh Failed"
End Sub

#If VBA7 Then
Function GetRibbon(ByVal lRibbonPointer As LongPtr) As Object
#Else
Function GetRibbon(ByVal lRibbonPointer As Long) As Object
#End If
Dim objRibbon As Object
CopyMemory objRibbon, lRibbonPointer, LenB(lRibbonPointer)
Set GetRibbon = objRibbon
Set objRibbon = Nothing
End Function

Sub vbArcRibbon_getSize(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE vbArcRibbon_ReturnValue
    returnedVal = vbArcRibbon_ReturnValue(control.ID, "Size")
End Sub

Sub vbArcRibbon_getLabel(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE vbArcRibbon_ReturnValue
    returnedVal = vbArcRibbon_ReturnValue(control.ID, "Label")
End Sub

Sub vbArcRibbon_getScreenTip(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE vbArcRibbon_ReturnValue
    returnedVal = vbArcRibbon_ReturnValue(control.ID, "ScreenTip")
End Sub

Sub vbArcRibbon_getSuperTip(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE vbArcRibbon_ReturnValue
    returnedVal = vbArcRibbon_ReturnValue(control.ID, "superTip")
End Sub

Sub vbArcRibbon_getVisible(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE vbArcRibbon_ReturnValue
    returnedVal = vbArcRibbon_ReturnValue(control.ID, "visible")
End Sub

Sub vbArcRibbon_getImage(control As IRibbonControl, ByRef returnedVal)
    '#INCLUDE IMAGE_FOLDER
    '#INCLUDE vbArcRibbon_ReturnValue
    '#INCLUDE FileExists
    '#INCLUDE LoadPictureGDI
    Dim image
    Dim ImageName As String
    ImageName = vbArcRibbon_ReturnValue(control.ID, "Image")
    If InStr(1, ImageName, ".") > 0 Then
        On Error GoTo ErrorHandler
        Dim strPath As String
        strPath = IMAGE_FOLDER
        If FileExists(strPath & ImageName) Then
            Set returnedVal = LoadPictureGDI(strPath & ImageName)
        Else
            returnedVal = "WordPicture"
        End If
    Else
        returnedVal = ImageName
    End If
ErrorHandler:
End Sub

Sub ShowImagePicker()
    If ActiveSheet.Name <> "vbArc_Addin_Settings" Then Exit Sub
    If ActiveSheet.Cells(2, ActiveCell.Column) = "image" Then
        uImageMso.Show
    Else
        MsgBox "Select a cell in column ""image"""
    End If
End Sub

Sub ShowLocalImagePicker()
    '#INCLUDE IMAGE_FOLDER
    '#INCLUDE FolderExists
    If ThisWorkbook.SHEETS("vbArc_Addin_Settings").Cells(2, ActiveCell.Column).TEXT Like "IMAGE" Then
        MsgBox "Select a cell in column ""IMAGE"""
        Exit Sub
    End If
    Dim initialFileName As String
    initialFileName = IMAGE_FOLDER
    Dim strFile As String
    Dim fd As Office.FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Filters.clear
        .title = "Choose an Image file"
        .AllowMultiSelect = False
        If FolderExists(initialFileName) Then
            .initialFileName = initialFileName
        End If
        If .Show = True Then
            strFile = .SelectedItems(1)
            strFile = Mid(strFile, InStrRev(strFile, "\") + 1)
            ActiveCell.Value = strFile
        End If
    End With
End Sub

Sub SetControlValue(controlID, TargetProperty, controlValue)
    Dim TargetSheet As Worksheet
    Set TargetSheet = ThisWorkbook.SHEETS("vbArc_Addin_Settings")
    Dim PropertyColumn As Long
    PropertyColumn = TargetSheet.rows(2).Find(TargetProperty).Column
    Dim ControlRow As Long
    ControlRow = TargetSheet.Columns(2).Find(controlID).row
    TargetSheet.Cells(ControlRow, PropertyColumn).Value = controlValue
End Sub

Function vbArcRibbon_ReturnValue(controlID, TargetProperty)
    Dim TargetSheet As Worksheet
    Set TargetSheet = ThisWorkbook.SHEETS("vbArc_Addin_Settings")
    Dim PropertyColumn As Long
    PropertyColumn = TargetSheet.rows(2).Find(TargetProperty).Column
    Dim ControlRow As Long
    ControlRow = TargetSheet.Columns(2).Find(controlID).row
    vbArcRibbon_ReturnValue = TargetSheet.Cells(ControlRow, PropertyColumn)
End Function

Function ControlLabel(control As IRibbonControl)
    '#INCLUDE vbArcRibbon_ReturnValue
    ControlLabel = vbArcRibbon_ReturnValue(control.ID, "label")
End Function

Sub vbArcRibbon_ButtonAction(control As IRibbonControl)
    '#INCLUDE CreateAllBars
    '#INCLUDE SaveThisAddin
    '#INCLUDE ShowUserformAuthorCard
    '#INCLUDE ShowUserformSnippetsWorkbook
    '#INCLUDE ShowUserformProjectManager
    Select Case control.ID
        Case "MainButtonAuthorCard"
            ShowUserformAuthorCard
        Case "MainButtonUpdate"
            Rem @TODO
        Case "MainButtonReload"
            CreateAllBars
        Case "MainButtonToggleIsAddin"
            ThisWorkbook.IsAddin = Not ThisWorkbook.IsAddin
        Case "MainButtonSave"
            SaveThisAddin
        Case "MainProjectManager"
            ShowUserformProjectManager
        Case "MainSnippetsWorkbook"
            ShowUserformSnippetsWorkbook
        Case "MainFinder"
            uFinder.Show
        Case "MainFormNavigator"
            uFormNavigator.Show
        Case "MainWorksheetNavigator"
            uSheetsNavigator.Show
        Case "MainRangeManager"
            uRangeControl.Show
        Case "MainImageManager"
            uImageControl.Show
        Case "MainFileManager"
            uFileManager.Show
        Case "MainSessionManager"
            uSessions.Show
        Case "MainAddinsManager"
            uAddinManager.Show
        Case "MainXray"
            uSkeleton.Show
        Case "MainNotekeeper"
            uMemoryKnots.Show
        Case "MainMouseRecorder"
            uMouseRecorder.Show
    End Select
End Sub

Sub SaveThisAddin()
    Dim WasOpen As Boolean: WasOpen = ThisWorkbook.IsAddin
    If Right(ThisWorkbook.Name, 4) = "xlam" Then ThisWorkbook.IsAddin = True
    ThisWorkbook.Save
    If WasOpen = False Then ThisWorkbook.IsAddin = False
End Sub

Sub ShowUserformAuthorCard()
    uDEV.Show
End Sub

Sub ShowUserformSnippetsWorkbook()
    ShowInVBE = False
    uSnippets.Show
End Sub

Sub ShowUserformSnippetsVBE()
    ShowInVBE = True
    Application.VBE.MainWindow.visible = True
    Application.VBE.MainWindow.SetFocus
    uSnippets.Show
End Sub

Sub ShowUserformProjectManager()
    uProjectManager.Show
End Sub

Sub ShowUserformComponentsRemove()
    '#INCLUDE ActiveCodepaneWorkbook
    Set pmWorkbook = ActiveCodepaneWorkbook
    RemoveComps.Show
End Sub

Sub ShowUserformComponentsAdd()
    '#INCLUDE ActiveCodepaneWorkbook
    Set pmWorkbook = ActiveCodepaneWorkbook
    AddComps.Show
End Sub

Sub ShowUserformComponentsRename()
    '#INCLUDE ActiveCodepaneWorkbook
    Set pmWorkbook = ActiveCodepaneWorkbook
    RenameComps.Show
End Sub

Sub ShowUserformReferences()
    uReferences.Show
End Sub

Sub ShowUserformFormBuilder()
    uFormBuilder.Show
End Sub

Sub ShowFormBuilderSheet()
    ThisWorkbook.IsAddin = False
    With ThisWorkbook.SHEETS("FormBuilder")
        .visible = xlSheetVisible
        .Activate
    End With
End Sub

Sub HideFormBuilderSheet()
    If Right(ThisWorkbook.Name, 4) = "xlam" Then ThisWorkbook.IsAddin = True
End Sub

Sub ShowUserformProjectExplorer()
    uProjectExplorer.Show
End Sub

Sub ShowUserformSkeleton()
    uSkeleton.Show
End Sub

Sub ShowUserformPickImageMSO()
    uImageMso.Show
End Sub

Sub AddReadmeToWorkbook()
    '#INCLUDE WorksheetExists
    If WorksheetExists("README", ActiveWorkbook) Then
        MsgBox "Sheet ""README"" already exists."
        Exit Sub
    Else
        ThisWorkbook.SHEETS("README").Copy ActiveWorkbook.SHEETS(1)
        ActiveWorkbook.SHEETS("README").visible = True
    End If
End Sub


