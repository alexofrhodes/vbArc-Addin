Attribute VB_Name = "M_Templates"
Rem @Folder TEMPLATES Declarations
Rem DspErrMsg Constants and Variables
Global Const Success        As Boolean = True
Global Const Failure        As Boolean = False
Global Const NoError        As Long = 0
Global Const LogError       As Long = 997
Global Const RtnError       As Long = 998
Global Const DspError       As Long = 999
Public bLogOnly             As Boolean
Public bDebug               As Boolean
Rem timer constants
Public Const mblncTimer As Boolean = True
Public mvarTimerName
Public mvarTimerStart

Rem @Folder TEMPLATES
Rem This is the main function that basically displays a message box formatted based on what the Err object contains and if we want to put our project in debug mode. It returns the button the user clicks: vbAbort, vbCancel, vbIgnore, vbRetry
Public Function DspErrMsg(ByVal sRoutineName As String, _
                          Optional ByVal sAddText As String = "") As VbMsgBoxResult
    If bLogOnly Then
        Debug.Print Now(), ThisWorkbook.Name & "!" & sRoutineName, err.Description, sAddText
    Else
        DspErrMsg = MsgBox( _
                    Prompt:="Error#" & err.Number & vbLf & err.Description & vbLf & sAddText, _
                    BUTTONS:=IIf(bDebug, vbAbortRetryIgnore, vbCritical) + _
                    IIf(err.Number < 1, 0, vbMsgBoxHelpButton), _
                    title:=sRoutineName, _
                    HelpFile:=err.HelpFile, _
                    Context:=err.HelpContext)
    End If
End Function

Rem templates
Function ErrHandlerTemplate(ProcedureName As String) As String
    '#INCLUDE DspErrMsg
    Dim s As String
    s = "ErrHandler:  'Error Handling, Clean Up and Routine Termination"
    s = s & vbNewLine & "    Select Case err.Number"
    s = s & vbNewLine & "        Case Is = NoError:                                        'No error - do nothing"
    s = s & vbNewLine & "        Case Is = 555:                                               'Add specific error handling here"
    s = s & vbNewLine & "        Case Is = RtnError: PROCEDURENAME = CVErr(xlErrDiv0)           'Return Error code to spreadsheet"
    s = s & vbNewLine & "        Case Is = LogError: Debug.Print cRoutine, err.Description 'Log the event and go on"
    s = s & vbNewLine & "        Case Else:"
    s = s & vbNewLine & "            Select Case DspErrMsg(cModule & ""."" & cRoutine)"
    s = s & vbNewLine & "                Case Is = vbAbort:  Stop: Resume       'Debug mode"
    s = s & vbNewLine & "                Case Is = vbRetry:  Resume             'Try again"
    s = s & vbNewLine & "                Case Is = vbIgnore:                    'End routine"
    s = s & vbNewLine & "            End Select"
    s = s & vbNewLine & "     End Select"
    ErrHandlerTemplate = s
End Function

Sub InjectTemplateModule()
    '#INCLUDE Inject
    '#INCLUDE TemplateModule
    Inject TemplateModule
End Sub

Function TemplateModule(Optional Module As VBComponent) As String
    '#INCLUDE DevInfo
    '#INCLUDE compare
    Dim ModuleName As String
    If Module Is Nothing Then
        ModuleName = "MODULE_NAME"
    Else
        ModuleName = Module.Name
    End If
    Dim s As String
    s = s & DevInfo & vbNewLine & vbNewLine
    s = s & "'   Version:    <Last Update Date goes here>" & vbNewLine
    s = s & "'   Description: General purpose library included in all projects" & vbNewLine & vbNewLine
    s = s & "'   Changelog" & vbNewLine
    s = s & "'   Date" & vbTab & vbTab & "Modification" & vbNewLine
    s = s & "'   " & Format(Date, "dd/mm/yy") & vbTab & "Initial Development" & vbNewLine
    s = s & vbNewLine
    s = s & "'Options" & vbNewLine
    s = s & "    Option Explicit" & vbNewLine
    s = s & "    Option Private Module" & vbNewLine
    s = s & "    Option Compare Text" & vbNewLine & vbNewLine
    s = s & "'Private Constants" & vbNewLine
    s = s & "    Private Const cModule    As String = " & ModuleName & vbNewLine
    TemplateModule = s
End Function

Function CopyTemplateFromSheet(Template As String)
    '#INCLUDE CodepaneSelection
    '#INCLUDE CLIP
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("Templates")
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0
    If Not cell Is Nothing Then
        If Len(CodepaneSelection) = 0 Then CLIP cell.OFFSET(0, 1)
    End If
    CopyTemplateFromSheet = cell.OFFSET(0, 1)
End Function

Sub InjectTemplateFromSheet(Template As String)
    '#INCLUDE Inject
    '#INCLUDE CodepaneSelection
    Dim cell As Range
    Dim ws As Worksheet
    Set ws = ThisWorkbook.SHEETS("Templates")
    On Error Resume Next
    Set cell = ws.Columns(1).SpecialCells(xlCellTypeConstants).Find(Template, LookAt:=xlWhole)
    On Error GoTo 0
    If Not cell Is Nothing Then
        If Len(CodepaneSelection) = 0 Then Inject cell.OFFSET(0, 1)
    End If
End Sub

Sub InjectScreenUpdating()
    '#INCLUDE InjectTemplateFromSheet
    InjectTemplateFromSheet "ScreenUpdating"
End Sub

Sub InjectForCounter()
    '#INCLUDE InjectTemplateFromSheet
    InjectTemplateFromSheet "forcounter"
End Sub

Sub InjectIf()
    '#INCLUDE InjectTemplateFromSheet
    InjectTemplateFromSheet "ifelse"
End Sub

Sub InjectOnErrorResumeNext()
    '#INCLUDE InjectTemplateFromSheet
    InjectTemplateFromSheet "onerror"
End Sub

Sub InjectEnableEvents()
    '#INCLUDE InjectTemplateFromSheet
    InjectTemplateFromSheet "EnableEvents"
End Sub

Sub InjectTemplateProcedure()
    '#INCLUDE Inject
    '#INCLUDE TemplateProcedure
    Inject TemplateProcedure
End Sub

Function TemplateProcedure(Optional FunctionName As String = "PROCEDURE_NAME") As String
    '#INCLUDE DevInfo
    '#INCLUDE ErrHandlerTemplate
    '#INCLUDE StartTimer
    '#INCLUDE EndTimer
    Dim Q As String: Q = """"
    Dim s As String
    s = s & "Function " & FunctionName & "(ByVal MyParameter as String) As Variant" & vbNewLine
    s = s & DevInfo & vbNewLine
    s = s & vbNewLine
    s = s & "'   Description:Procedure Description" & vbNewLine
    s = s & "'   Inputs:     MyParameter  Describe its purpose" & vbNewLine
    s = s & "'   Outputs:    Success: <return this>" & vbNewLine
    s = s & "'               Failure: <return this>" & vbNewLine
    s = s & "'   Requisites  Routines    ModuleName.ProcedureName" & vbNewLine
    s = s & "'               Classes      Class Module Name" & vbNewLine
    s = s & "'               Forms        User Form Name" & vbNewLine
    s = s & "'               Tables       Table Name" & vbNewLine
    s = s & "'               References   Reference" & vbNewLine
    s = s & "'   Notes <add if needed>" & vbNewLine
    s = s & "'   Example: ?" & FunctionName & "(MyParameter)" & vbNewLine & vbNewLine
    s = s & "'   Changelog" & vbNewLine
    s = s & "'   Date        Modification" & vbNewLine
    s = s & "'   " & Format(Date, "DD/MM/YY") & "    Initial Release" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Check Inputs and Requisites" & vbNewLine
    s = s & "    If sParameter = cvbNullString then Err.Raise DspError, , ""Parameter missing""" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Declarations" & vbNewLine
    s = s & "    Const cRoutine as String = " & Q & FunctionName & Q
    s = s & "'   Error Handling Initialization" & vbNewLine
    s = s & "    On Error GoTo ErrHandler" & vbNewLine
    s = s & "    " & FunctionName & " = Failure    'Assume failure" & vbNewLine
    s = s & vbNewLine
    s = s & "'   Initialize Variables" & vbNewLine & vbNewLine
    s = s & "'   Procedure" & vbNewLine
    s = s & "    Application.screenupdating=false" & vbNewLine
    s = s & "    StartTimer " & FunctionName & vbNewLine & "    " & vbNewLine
    s = s & "    " & FunctionName & " = Success    'Successful finish" & vbNewLine
    s = s & "    EndTimer" & vbNewLine & vbNewLine
    s = s & "NormalExit:" & vbNewLine
    s = s & "    Application.screenupdating=false" & vbNewLine
    s = s & "    Exit Sub" & vbNewLine
    s = s & vbNewLine
    s = s & ErrHandlerTemplate(FunctionName) & vbNewLine
    s = s & "End Function"
    TemplateProcedure = s
End Function

Rem timer
Public Function StartTimer(TimerName)
    On Error GoTo ERR_HANDLER
    If mblncTimer Then
        mvarTimerName = TimerName
        mvarTimerStart = Timer
    End If
    On Error Resume Next
    Exit Function
ERR_HANDLER:
    MsgBox err.Number & " " & err.Description, vbCritical, "StartTimer()"
End Function

Public Function EndTimer()
    '#INCLUDE FoldersCreate
    '#INCLUDE TxtAppend
    On Error GoTo ERR_HANDLER
    Dim strFile As String
    Dim strContent As String
    If mblncTimer Then
        Dim strPath As String
        strPath = Environ("USERprofile") & "\My Documents\vbArc\Timers\"
        FoldersCreate strPath
        strFile = strPath & mvarTimerName & ".txt"
        Rem strFile = ThisWorkbook.path & "\" _
        & Left(ThisWorkbook.Name, InStr(1, ThisWorkbook.Name, ".") - 1) _
        & "TimerLog.txt"
        If Len(Dir(strFile)) = 0 Then
            strContent = _
                       "Timestamp" & vbTab & vbTab & vbTab & vbTab & _
                       "ElapsedTime" & vbTab & vbTab & _
                       "TimerName"
            TxtAppend strFile, strContent
        End If
        strContent = Now() & vbTab & vbTab & _
                           Format(Timer - mvarTimerStart, "0.000000") & vbTab & vbTab & vbTab & _
                           mvarTimerName
        TxtAppend strFile, strContent
    End If
    On Error Resume Next
    Exit Function
ERR_HANDLER:
    MsgBox err.Number & " " & err.Description, vbCritical, "EndTimer()"
End Function

Public Sub dp(var As Variant)
    '#INCLUDE printRange
    '#INCLUDE printArray
    '#INCLUDE printCollection
    '#INCLUDE printDictionary
    Dim element     As Variant
    Dim i As Long
    Select Case TypeName(var)
        Case Is = "String", "Long", "Integer", "Boolean"
            Debug.Print var
        Case Is = "Variant()", "String()", "Long()", "Integer()"
            printArray var
        Case Is = "Collection"
            printCollection var
        Case Is = "Dictionary"
            printDictionary var
        Case Is = "Range"
            printRange var
        Case Is = "Date"
            Debug.Print var
        Case Else
    End Select
End Sub

Public Sub printRange(var As Variant)
    '#INCLUDE dp
    '#INCLUDE Combine2Array
    If var.Areas.count = 1 Then
        dp var.Value
    Else
        Dim out As Variant
        Dim temp As Variant
        Dim i As Long
        For i = 1 To var.Areas.count
            temp = var.Areas(i).Value
            If IsEmpty(out) Then
                out = temp
            Else
                out = Combine2Array(out, temp)
            End If
        Next
        dp out
    End If
End Sub

Private Sub printArray(var As Variant)
    '#INCLUDE DPH
    '#INCLUDE ArrayDimensions
    If ArrayDimensions(var) = 1 Then
        Debug.Print Join(var, vbNewLine)
    ElseIf ArrayDimensions(var) > 1 Then
        DPH var
    End If
End Sub

Private Sub printCollection(var As Variant)
    '#INCLUDE dp
    Dim elem        As Variant
    For Each elem In var
        dp elem
    Next elem
End Sub

Private Sub printDictionary(var As Variant)
    '#INCLUDE dp
    Dim i As Long: Dim iCount As Long
    Dim arrKeys
    Dim sKey        As String
    Dim varItem
    With var
        iCount = .count
        arrKeys = .keys
        iCount = UBound(arrKeys, 1)
        For i = 0 To iCount
            sKey = arrKeys(i)
            If IsObject(.item(sKey)) Then
                Debug.Print sKey & " : "
                dp (.item(sKey))
            Else
                Debug.Print sKey & " : " & .item(sKey)
            End If
        Next i
    End With
End Sub

Private Sub DPH(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '#INCLUDE DebugPrintHairetu
    Call DebugPrintHairetu(Hairetu, HyoujiMaxNagasa, HairetuName)
End Sub

Public Function ArrayDimensions(ByVal vArray As Variant) As Long
    Dim dimnum      As Long
    Dim ErrorCheck As Long
    On Error GoTo FinalDimension
    For dimnum = 1 To 60000
        ErrorCheck = LBound(vArray, dimnum)
    Next
FinalDimension:
    ArrayDimensions = dimnum - 1
End Function

Private Sub DebugPrintHairetu(ByVal Hairetu, Optional HyoujiMaxNagasa%, Optional HairetuName$)
    '#INCLUDE ShortenToByteCharacters
    Dim i&, j&, k&, m&, n&
    Dim TateMin&, TateMax&, YokoMin&, YokoMax&
    Dim WithTableHairetu
    Dim NagasaList, MaxNagasaList
    Dim NagasaOnajiList
    Dim OutputList
    Const SikiriMoji$ = "|"
    Dim Jigen2%
    On Error Resume Next
    Jigen2 = UBound(Hairetu, 2)
    On Error GoTo 0
    If Jigen2 = 0 Then
        Hairetu = Application.Transpose(Hairetu)
    End If
    TateMin = LBound(Hairetu, 1)
    TateMax = UBound(Hairetu, 1)
    YokoMin = LBound(Hairetu, 2)
    YokoMax = UBound(Hairetu, 2)
    ReDim WithTableHairetu(1 To TateMax - TateMin + 1 + 1, 1 To YokoMax - YokoMin + 1 + 1)
    For i = 1 To TateMax - TateMin + 1
        WithTableHairetu(i + 1, 1) = TateMin + i - 1
            For j = 1 To YokoMax - YokoMin + 1
                WithTableHairetu(1, j + 1) = YokoMin + j - 1
                    WithTableHairetu(i + 1, j + 1) = Hairetu(i - 1 + TateMin, j - 1 + YokoMin)
                    Next j
                Next i
                n = UBound(WithTableHairetu, 1)
                m = UBound(WithTableHairetu, 2)
                ReDim NagasaList(1 To n, 1 To m)
                ReDim MaxNagasaList(1 To m)
                Dim TmpStr$
                For j = 1 To m
                    For i = 1 To n
                        If j > 1 And HyoujiMaxNagasa <> 0 Then
                            TmpStr = WithTableHairetu(i, j)
                            WithTableHairetu(i, j) = ShortenToByteCharacters(TmpStr, HyoujiMaxNagasa)
                            End If
                            NagasaList(i, j) = LenB(StrConv(WithTableHairetu(i, j), vbFromUnicode))
                            MaxNagasaList(j) = WorksheetFunction.Max(MaxNagasaList(j), NagasaList(i, j))
                        Next i
                    Next j
                    ReDim NagasaOnajiList(1 To n, 1 To m)
                    Dim TmpMaxNagasa&
                    For j = 1 To m
                        TmpMaxNagasa = MaxNagasaList(j)
                        For i = 1 To n
                            NagasaOnajiList(i, j) = WithTableHairetu(i, j) & WorksheetFunction.Rept(" ", TmpMaxNagasa - NagasaList(i, j))
                        Next i
                    Next j
                    ReDim OutputList(1 To n)
                    For i = 1 To n
                        For j = 1 To m
                            If j = 1 Then
                                OutputList(i) = NagasaOnajiList(i, j)
                            Else
                                OutputList(i) = OutputList(i) & SikiriMoji & NagasaOnajiList(i, j)
                            End If
                        Next j
                    Next i
                    Debug.Print HairetuName
                    For i = 1 To n
                        Debug.Print OutputList(i)
                    Next i
                End Sub

Private Function ShortenToByteCharacters(Mojiretu$, ByteNum%)
    '#INCLUDE CalculateByteCharacters
    '#INCLUDE TextDecomposition
    Dim OriginByte%
    Dim output
    OriginByte = LenB(StrConv(Mojiretu, vbFromUnicode))
    If OriginByte <= ByteNum Then
        output = Mojiretu
    Else
        Dim RuikeiByteList, BunkaiMojiretu
        RuikeiByteList = CalculateByteCharacters(Mojiretu)
        BunkaiMojiretu = TextDecomposition(Mojiretu)
        Dim AddMoji$
        AddMoji = "."
        Dim i&, n&
        n = Len(Mojiretu)
        For i = 1 To n
            If RuikeiByteList(i) < ByteNum Then
                output = output & BunkaiMojiretu(i)
            ElseIf RuikeiByteList(i) = ByteNum Then
                If LenB(StrConv(BunkaiMojiretu(i), vbFromUnicode)) = 1 Then
                    output = output & AddMoji
                Else
                    output = output & AddMoji & AddMoji
                End If
                Exit For
            ElseIf RuikeiByteList(i) > ByteNum Then
                output = output & AddMoji
                Exit For
            End If
        Next i
    End If
    ShortenToByteCharacters = output
End Function

Private Function CalculateByteCharacters(Mojiretu$)
    Dim MojiKosu%
    MojiKosu = Len(Mojiretu)
    Dim output
    ReDim output(1 To MojiKosu)
    Dim i&
    Dim TmpMoji$
    For i = 1 To MojiKosu
        TmpMoji = Mid(Mojiretu, i, 1)
        If i = 1 Then
            output(i) = LenB(StrConv(TmpMoji, vbFromUnicode))
        Else
            output(i) = LenB(StrConv(TmpMoji, vbFromUnicode)) + output(i - 1)
        End If
    Next i
    CalculateByteCharacters = output
End Function

Private Function TextDecomposition(Mojiretu$)
    Dim i&, n&
    Dim output
    n = Len(Mojiretu)
    ReDim output(1 To n)
    For i = 1 To n
        output(i) = Mid(Mojiretu, i, 1)
    Next i
    TextDecomposition = output
End Function

Function DpHeader(str As Variant, Optional lvl As Integer = 1, Optional Character As String = "'", _
                  Optional top As Boolean, Optional bottom As Boolean) As String
    '#INCLUDE LargestLength
    If lvl < 1 Then lvl = 1
    If Character = "" Then Character = "'"
    Dim Indentation As Integer
    Indentation = (lvl * 4) - 4 + 1
    Dim QUOTE As String: QUOTE = "'"
    Dim s As String
    Dim element As Variant
    If top = True Then s = vbNewLine & QUOTE & String(Indentation + LargestLength(str), Character) & vbNewLine
    If TypeName(str) <> "String" Then
        For Each element In str
            s = s & QUOTE & String(Indentation, Character) & element & vbNewLine
        Next
    Else
        s = s & QUOTE & String(Indentation, Character) & str
    End If
    If bottom = True Then s = s & QUOTE & String(Indentation + LargestLength(str), Character)
    DpHeader = s
End Function


