VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEditOpenXML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mbCreateBackup As Boolean
Private mvSourceFile As Variant
Private msSheet2Change As String
Private msSheetId As String
Private msSheetFileName As String
Private mbAddedZip As Boolean
Private mvXMLFolderRoot As Variant
Private mvxmlfolder As Variant
Public Enum XMLFolder
    XMLFolder_root = 1
    XMLFolder_rels = 2
    XMLFolder_xl = 3
    XMLFolder_customUI = 4
    XMLFolder_docProps = 5
End Enum

Public Sub ExtractRibbonX(sFullFile As String, sSaveFile As String)
    Dim cEditOpenXML As clsEditOpenXML
    Dim sXML As String
    Dim oXMLDoc As MSXML2.DOMDocument
    Set cEditOpenXML = New clsEditOpenXML
    With cEditOpenXML
        .CreateBackup = False
        .SourceFile = sFullFile
        .UnzipFile
        sXML = .GetXMLFromFile("customUI.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            Set oXMLDoc = New DOMDocument
            oXMLDoc.loadXML sXML
            oXMLDoc.Save sSaveFile
        End If
        sXML = .GetXMLFromFile("customUI14.xml", XMLFolder_customUI)
        If Len(sXML) > 0 Then
            Set oXMLDoc = New DOMDocument
            oXMLDoc.loadXML sXML
            oXMLDoc.Save Replace(sSaveFile, ".xml", "14.xml")
        End If
    End With
    Set cEditOpenXML = Nothing
End Sub

Public Sub WriteRibbonXML2File(sFile As String, sXML As String)
    Dim cEditOpenXML As clsEditOpenXML
   
    Set cEditOpenXML = New clsEditOpenXML
   
    With cEditOpenXML
        .SourceFile = sFile
       
        .UnzipFile
       
        .WriteXML2File sXML, "customUI.xml", XMLFolder_customUI
       
        .ZipAllFilesInFolder
    End With
   
    Set cEditOpenXML = Nothing
End Sub

Public Function GetXMLFromFile(sFilename As String, sXMLFolder As XMLFolder) As String
    Dim oXMLDoc As MSXML2.DOMDocument
    If Len(XMLFolder(sXMLFolder)) = 0 Then
        GetXMLFromFile = ""
    Else
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(sXMLFolder) & sFilename
        GetXMLFromFile = oXMLDoc.xml
        Set oXMLDoc = Nothing
    End If
End Function

Public Sub WriteXML2File(sXML As String, sFilename As String, sXMLFolder As XMLFolder)
    Dim oXMLDoc As MSXML2.DOMDocument
    Set oXMLDoc = New MSXML2.DOMDocument
    If sXMLFolder = XMLFolder_customUI Then
        If Not FolderExists(XMLFolder(XMLFolder_customUI)) Then
            MkDir XMLFolder(XMLFolder_customUI)
            oXMLDoc.loadXML sXML
            oXMLDoc.Save XMLFolder(sXMLFolder) & sFilename
            AddCustomUIToRels
        End If
    End If
    oXMLDoc.loadXML sXML
    oXMLDoc.Save XMLFolder(sXMLFolder) & sFilename
End Sub

Public Sub AddCustomUIToRels()
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oXMLElement As MSXML2.IXMLDOMNode
    Dim oXMLAttrib As MSXML2.IXMLDOMAttribute
    Dim oNamedNodeMap As MSXML2.IXMLDOMNamedNodeMap
    Dim oXMLRelsList As MSXML2.IXMLDOMNodeList
    Set oXMLDoc = New MSXML2.DOMDocument
    oXMLDoc.Load XMLFolder(XMLFolder_rels) & ".rels"
    Set oXMLElement = oXMLDoc.createNode(1, "Relationship", "http://schemas.openxmlformats.org/package/2006/relationships")
    Set oNamedNodeMap = oXMLElement.Attributes
    Set oXMLAttrib = oXMLDoc.createAttribute("Id")
    oXMLAttrib.nodeValue = "cuID"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLAttrib = oXMLDoc.createAttribute("Type")
    oXMLAttrib.nodeValue = "http://schemas.microsoft.com/office/2006/relationships/ui/extensibility"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLAttrib = oXMLDoc.createAttribute("Target")
    oXMLAttrib.nodeValue = "customUI/customUI.xml"
    oNamedNodeMap.setNamedItem oXMLAttrib
    Set oXMLRelsList = oXMLDoc.selectNodes("/Relationships")
    oXMLRelsList.item(0).appendChild oXMLElement
    oXMLDoc.Save XMLFolder(XMLFolder_rels) & ".rels"
    Set oXMLAttrib = Nothing
    Set oXMLElement = Nothing
    Set oXMLDoc = Nothing
End Sub

Private Function GetSheetIdFromSheetName(sSheetName) As String
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLTemp As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.baseName = "sheets" Then
                    For Each oXMLTemp In oXMLChildNode.ChildNodes
                        If oXMLTemp.Attributes.getNamedItem("name").nodeValue = sSheetName Then
                            GetSheetIdFromSheetName = oXMLTemp.Attributes.getNamedItem("r:id").nodeValue
                            Exit Function
                        End If
                    Next
                End If
            Next
        Next
    End If
End Function

Public Function GetSheetFileNameFromId(sSheetId As String) As String
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    If XMLFolder(XMLFolder_xl) <> "" And Sheet2Change <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "_rels\workbook.xml.rels"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.Attributes.getNamedItem("Id").nodeValue = sSheetId Then
                    GetSheetFileNameFromId = oXMLChildNode.Attributes.getNamedItem("Target").nodeValue
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Private Function GetSheetNameFromId(sId As String) As String
    Dim oXMLDoc As MSXML2.DOMDocument
    Dim oxmlNode As MSXML2.IXMLDOMNode
    Dim oXMLChildNode As MSXML2.IXMLDOMNode
    Dim oXMLChildChildNode As MSXML2.IXMLDOMNode
    If mvxmlfolder(XMLFolder_xl) <> "" Then
        Set oXMLDoc = New MSXML2.DOMDocument
        oXMLDoc.Load XMLFolder(XMLFolder_xl) & "workbook.xml"
        For Each oxmlNode In oXMLDoc.ChildNodes
            For Each oXMLChildNode In oxmlNode.ChildNodes
                If oXMLChildNode.nodeName = "sheets" Then
                    For Each oXMLChildChildNode In oXMLChildNode.ChildNodes
                        If oXMLChildChildNode.Attributes.getNamedItem("r:id").nodeValue = "rId" & val(sId) + 1 Then
                            GetSheetNameFromId = oXMLChildChildNode.Attributes.getNamedItem("name").nodeValue
                            Exit Function
                        End If
                    Next
                    Exit Function
                End If
            Next
        Next
    End If
End Function

Public Sub ZipAllFilesInFolder()
    Dim oShellApp As Object
    Dim sDate As String
    Dim sDefPath As String
    Dim vFileNameZip As Variant
    Dim fso As Object
    Dim lFileCt As Long
    Set fso = CreateObject("scripting.filesystemobject")
    sDate = Format(Now, " dd-mmm-yy h-mm-ss")
    vFileNameZip = SourceFile & sDate & ".zip"
    NewZip vFileNameZip
    Set oShellApp = CreateObject("Shell.Application")
    lFileCt = oShellApp.Namespace(FolderName & "Unzipped " & fileName & Application.PathSeparator).items.count
    oShellApp.Namespace(vFileNameZip).CopyHere oShellApp.Namespace(FolderName & "Unzipped " & fileName & Application.PathSeparator).items
    On Error Resume Next
    Do Until oShellApp.Namespace(vFileNameZip).items.count = lFileCt
        Application.Wait (Now + TimeValue("0:00:01"))
    Loop
    DoEvents
    Kill SourceFile
    Name vFileNameZip As SourceFile
    On Error Resume Next
    fso.DeleteFolder FolderName & "Unzipped " & fileName, True
    On Error GoTo 0
    Set oShellApp = Nothing
End Sub

Public Sub UnzipFile()
    Dim fso As Object
    Dim oShellApp As Object
    Set fso = CreateObject("scripting.filesystemobject")
    XMLFolderRoot = FolderName
    If Right(XMLFolderRoot, 1) <> Application.PathSeparator Then
        XMLFolderRoot = XMLFolderRoot & "\UnZipped " & fileName & Application.PathSeparator
    Else
        XMLFolderRoot = XMLFolderRoot & "UnZipped " & fileName & Application.PathSeparator
    End If
    On Error Resume Next
    fso.DeleteFolder XMLFolderRoot & "*", True
    Kill XMLFolderRoot & "*.*"
    If FolderExists(XMLFolderRoot) = False Then
        MkDir XMLFolderRoot
    End If
    Set oShellApp = CreateObject("Shell.Application")
    oShellApp.Namespace(XMLFolderRoot).CopyHere oShellApp.Namespace(SourceFile).items
    On Error Resume Next
    fso.DeleteFolder Environ("Temp") & "\Temporary Directory*", True
    Set oShellApp = Nothing
    Set fso = Nothing
    Exit Sub
End Sub

Sub NewZip(sPath)
    If Len(Dir(sPath)) > 0 Then Kill sPath
    Open sPath For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String(18, 0)
    Close #1
End Sub

Public Property Get CreateBackup() As Boolean
    CreateBackup = mbCreateBackup
End Property

Public Property Let CreateBackup(ByVal bCreateBackup As Boolean)
    mbCreateBackup = bCreateBackup
End Property

Private Sub Class_Initialize()
    CreateBackup = True
End Sub

Public Property Get SourceFile() As Variant
    SourceFile = mvSourceFile
End Property

Public Property Let SourceFile(ByVal vSourceFile As Variant)
    mvSourceFile = vSourceFile
    If CreateBackup Then
        If Len(Dir(vSourceFile & "(backup)")) > 0 Then
            Kill vSourceFile & "(backup)"
        End If
        FileCopy vSourceFile, vSourceFile & "(backup)"
    End If
    If Not vSourceFile Like "*.zip" Then
        Name vSourceFile As vSourceFile & ".zip"
        mvSourceFile = mvSourceFile & ".zip"
        AddedZip = True
    End If
End Property

Public Property Get FolderName() As Variant
    FolderName = Mid(SourceFile, 1, InStrRev(SourceFile, Application.PathSeparator))
End Property

Public Property Get fileName() As Variant
    If SourceFile <> "" Then
        fileName = Mid(SourceFile, InStrRev(SourceFile, Application.PathSeparator) + 1, Len(SourceFile))
    End If
End Property

Public Property Get XMLFolder(sXMLFolder As XMLFolder) As String
    Select Case sXMLFolder
        Case Is = XMLFolder_root
            XMLFolder = mvXMLFolderRoot
        Case Is = XMLFolder_customUI
            XMLFolder = mvXMLFolderRoot & "customUI" & Application.PathSeparator
        Case Is = XMLFolder_docProps
            XMLFolder = mvXMLFolderRoot & "docProps" & Application.PathSeparator
        Case Is = XMLFolder_rels
            XMLFolder = mvXMLFolderRoot & "_rels" & Application.PathSeparator
        Case Is = XMLFolder_xl
            XMLFolder = mvXMLFolderRoot & "xl" & Application.PathSeparator
    End Select
End Property

Public Property Get Sheet2Change() As String
    Sheet2Change = msSheet2Change
End Property

Public Property Let Sheet2Change(ByVal sSheet2Change As String)
    msSheet2Change = sSheet2Change
    SheetId = GetSheetIdFromSheetName(sSheet2Change)
    If SheetId <> "" Then
        SheetFileName = GetSheetFileNameFromId(SheetId)
    End If
End Property

Public Property Get SheetId() As String
    SheetId = msSheetId
End Property

Public Property Let SheetId(ByVal sSheetId As String)
    msSheetId = sSheetId
End Property

Public Property Get SheetFileName() As String
    SheetFileName = msSheetFileName
End Property

Public Property Let SheetFileName(ByVal sSheetFileName As String)
    msSheetFileName = sSheetFileName
End Property

Private Property Get AddedZip() As Boolean
    AddedZip = mbAddedZip
End Property

Private Property Let AddedZip(ByVal bAddedZip As Boolean)
    mbAddedZip = bAddedZip
End Property

Private Sub Class_Terminate()
    Dim fso As Object
    If AddedZip Then
        Name SourceFile As left(SourceFile, Len(SourceFile) - 4)
    End If
    On Error Resume Next
    fso.DeleteFolder XMLFolderRoot, True
End Sub

Private Property Get XMLFolderRoot() As Variant
    XMLFolderRoot = mvXMLFolderRoot
End Property

Private Property Let XMLFolderRoot(ByVal vXMLFolderRoot As Variant)
    mvXMLFolderRoot = vXMLFolderRoot
End Property

