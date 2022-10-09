Attribute VB_Name = "F_Userforms"
Rem @Folder Userforms
Option Explicit
Option Compare Text
Rem @Subfolder Userforms>Transparent Declarations
Rem Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const GWL_STYLE As Long = (-16)
Private Const GWL_EXSTYLE As Long = (-20)
Private Const WS_CAPTION As Long = &HC00000
Private Const WS_EX_DLGMODALFRAME As Long = &H1
Private Const WS_EX_LAYERED = &H80000
Private Const LWA_COLORKEY = &H1
Private Const LWA_ALPHA = &H2
Private m_sngDownX As Single
Private m_sngDownY As Single
Rem Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Rem Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Rem @Subfolder Userforms>Parent Declarations
Public Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Public Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Public Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Public Const FORMAT_MESSAGE_FROM_STRING = &H400
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Public Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Public Const FORMAT_MESSAGE_TEXT_LEN = 160
Public Const MAX_PATH = 260
Public Const GWL_HWNDPARENT As Long = -8
Public Const GW_OWNER = 4
Public Declare PtrSafe Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare PtrSafe Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public VBEditorHWnd As Long
Public ApplicationHWnd As Long
Public ExcelDeskHWnd As Long
Public ActiveWindowHWnd As Long
Public UserFormHWnd As Long
Public WindowsDesktopHWnd As Long
Public Const GA_ROOT As Long = 2
Public Const GA_ROOTOWNER As Long = 3
Public Const GA_PARENT As Long = 1
Public Declare PtrSafe Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Public Declare PtrSafe Function GetAncestor Lib "user32.dll" (ByVal hWnd As Long, ByVal gaFlags As Long) As Long
Public Declare PtrSafe Function GetDesktopWindow Lib "user32" () As Long
Public Declare PtrSafe Function GetParent Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare PtrSafe Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Const C_EXCEL_APP_WINDOWCLASS = "XLMAIN"
Public Const C_EXCEL_DESK_WINDOWCLASS = "XLDESK"
Public Const C_EXCEL_WINDOW_WINDOWCLASS = "EXCEL7"
Public Const USERFORM_WINDOW_CLASS = "ThunderDFrame"
Public Const C_VBA_USERFORM_WINDOWCLASS = "ThunderDFrame"
Rem Form on top
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_TOP = 0
Public Const HWND_BOTTOM = 1
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Declare PtrSafe Function SetWindowPos Lib "user32" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal X As LongPtr, ByVal Y As LongPtr, ByVal cx As LongPtr, ByVal cy As LongPtr, ByVal uFlags As LongPtr) As Long
Rem ---
#If VBA7 Then
    Public Declare PtrSafe Function SetParent Lib "user32" ( _
    ByVal hwndChild As LongPtr, _
    ByVal hWndNewParent As LongPtr) As LongPtr
    Public Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
    ByVal hWnd As LongPtr) As Long
    Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As LongPtr
    Public Declare PtrSafe Function FormatMessage Lib "kernel32" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long _
    , lpSource As Any _
    , ByVal dwMessageId As Long _
    , ByVal dwLanguageId As Long _
    , ByVal lpBuffer As String _
    , ByVal nSize As Long _
    , Arguments As LongPtr) As Long
#Else
    Public Declare  Function SetParent Lib "user32" ( _
    ByVal hWndChild As Long, _
    ByVal hWndNewParent As Long) As Long
    Public Declare  Function SetForegroundWindow Lib "user32" ( _
    ByVal hwnd As Long) As Long
    Public Declare  Function FindWindow Lib "user32" Alias "FindWindowA" ( _
    ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long
    Public Declare  Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageA" ( _
    ByVal dwFlags As Long, _
    ByRef lpSource As Any, _
    ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, _
    ByVal lpBuffer As String, _
    ByVal nSize As Long, _
    ByRef Arguments As Long) As Long
#End If
Rem Closeby
Public Enum CloseBy
    user = 0
    Code = 1
    WindowsOS = 2
    TaskManager = 3
End Enum

Rem FlashControl
Public Declare PtrSafe Function getTickCount Lib "kernel32" Alias "GetTickCount" () As Long
Public Const Black As Long = &H80000012
Public Const Red As Long = &HFF&
Rem ControlID
Public Const ControlIDCheckBox = "Forms.CheckBox.1"
Public Const ControlIDComboBox = "Forms.ComboBox.1"
Public Const ControlIDCommandButton = "Forms.CommandButton.1"
Public Const ControlIDFrame = "Forms.Frame.1"
Public Const ControlIDImage = "Forms.Image.1"
Public Const ControlIDLabel = "Forms.Label.1"
Public Const ControlIDListBox = "Forms.ListBox.1"
Public Const ControlIDMultiPage = "Forms.MultiPage.1"
Public Const ControlIDOptionButton = "Forms.OptionButton.1"
Public Const ControlIDScrollBar = "Forms.ScrollBar.1"
Public Const ControlIDSpinButton = "Forms.SpinButton.1"
Public Const ControlIDTabStrip = "Forms.TabStrip.1"
Public Const ControlIDTextBox = "Forms.TextBox.1"
Public Const ControlIDToggleButton = "Forms.ToggleButton.1"
Rem other
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLongA Lib "user32" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Rem @Folder Userforms
Rem @Subfolder Userforms>Transparent
Rem MakeFormTransparent me
Rem MakeFormBorderless Me
Public Sub MakeFormTransparent(frm As Object, Optional color As Variant)
    '#INCLUDE MakeFormBorderless
    Dim formhandle As Long
    Dim bytOpacity As Byte
    formhandle = CLng(FindWindow(vbNullString, frm.Caption))
    If IsMissing(color) Then color = vbWhite
    bytOpacity = 100
    SetWindowLong formhandle, GWL_EXSTYLE, GetWindowLong(formhandle, GWL_EXSTYLE) Or WS_EX_LAYERED
    frm.BackColor = color
    SetLayeredWindowAttributes formhandle, color, bytOpacity, LWA_COLORKEY
End Sub

Public Sub MakeFormBorderless(frm As Object)
    Dim lngWindow As Long
    Dim lFrmHdl As Long
    lFrmHdl = CLng(FindWindow(vbNullString, frm.Caption))
    lngWindow = GetWindowLong(lFrmHdl, GWL_STYLE)
    lngWindow = lngWindow And (Not WS_CAPTION)
    SetWindowLong lFrmHdl, GWL_STYLE, lngWindow
    lngWindow = GetWindowLong(lFrmHdl, GWL_EXSTYLE)
    lngWindow = lngWindow And Not WS_EX_DLGMODALFRAME
    SetWindowLong lFrmHdl, GWL_EXSTYLE, lngWindow
    DrawMenuBar lFrmHdl
End Sub

Rem @Subfolder Userforms>Parent
Public Sub UserformOnTop(form As Object)
    Const C_VBA6_USERFORM_CLASSNAME = "ThunderDFrame"
    Dim ret As Long
    Dim formHWnd As Long
    formHWnd = CLng(FindWindow(C_VBA6_USERFORM_CLASSNAME, form.Caption))
    If formHWnd = 0 Then
        Debug.Print err.LastDllError
    End If
    ret = SetWindowPos(formHWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE)
    If ret = 0 Then
        Debug.Print err.LastDllError
    End If
End Sub

Public Sub MakeUserFormChildOfVBEditor(GivenFormCaption As String)
    '#INCLUDE DisplayErrorText
    #If VBA7 Then
        Dim VBEWindowPointer As LongPtr
        Dim UserFormWindowPointer As LongPtr
        Dim ReturnOfSetParentAPI As LongPtr
    #Else
        Dim VBEWindowPointer As Long
        Dim UserFormWindowPointer As Long
        Dim ReturnOfSetParentAPI As Long
    #End If
    Dim ErrorNumber As Long
    VBEWindowPointer = Application.VBE.MainWindow.hWnd
    UserFormWindowPointer = FindWindow(lpClassName:=USERFORM_WINDOW_CLASS, lpWindowName:=GivenFormCaption)
    Const ERROR_NUMBER_FOR_SETPARENT_API = 0
    ReturnOfSetParentAPI = SetParent(hwndChild:=UserFormWindowPointer, hWndNewParent:=VBEWindowPointer)
    If ReturnOfSetParentAPI = ERROR_NUMBER_FOR_SETPARENT_API Then
        ErrorNumber = err.LastDllError
        DisplayErrorText "Error With SetParent", ErrorNumber
    Else
        Debug.Print GivenFormCaption & " is child of VBE Window."
    End If
    SetForegroundWindow UserFormWindowPointer
End Sub

Sub DisplayErrorText(Context As String, ErrNum As Long)
    Rem  Displays a standard error message box. For this
    Rem  procedure, ErrNum should be the number returned
    Rem  by the GetLastError API function or the value
    Rem  of Err.LastDllError. It is NOT the number
    Rem  returned by Err.Number.
    '#INCLUDE GetSystemErrorMessageText
    Dim ErrText As String
    ErrText = GetSystemErrorMessageText(ErrNum)
    MsgBox Context & vbCrLf & _
           "Error Number: " & CStr(ErrNum) & vbCrLf & _
           "Error Text:   " & ErrText, vbOKOnly
End Sub

Function GetSystemErrorMessageText(ErrorNumber As Long) As String
    Rem  This function gets the system error message text that corresponds to the error code returned by the
    Rem  GetLastError API function or the Err.LastDllError property. It may be used ONLY for these error codes.
    Rem  These are NOT the error numbers returned by Err.Number (for these errors, use Err.Description to get
    Rem  the description of the message).
    Rem  The error number MUST be the value returned by GetLastError or Err.LastDLLError.
    Rem
    Rem  In general, you should use Err.LastDllError rather than GetLastError because under some circumstances the value of
    Rem  GetLastError will be reset to 0 before the value is returned to VB. Err.LastDllError will always reliably return
    Rem  the last error number raised in a DLL.
    '#INCLUDE TrimToNull
    Dim ErrorText As String
    Dim ErrorTextLength As Long
    Dim FormatMessageResult As Long
    Dim LanguageID As Long
    LanguageID = 0&
    ErrorText = String$(FORMAT_MESSAGE_TEXT_LEN, " ")
    ErrorTextLength = Len(ErrorText)
    FormatMessageResult = 0&
    #If VBA7 Then
        Dim FormatMessageAPILastArgument As LongPtr
        FormatMessageAPILastArgument = 0
    #Else
        Dim FormatMessageAPILastArgument As Long
        FormatMessageAPILastArgument = 0
    #End If
    FormatMessageResult = FormatMessage( _
                          dwFlags:=FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
                          lpSource:=0&, _
                          dwMessageId:=ErrorNumber, _
                          dwLanguageId:=0&, _
                          lpBuffer:=ErrorText, _
                          nSize:=ErrorTextLength, _
                          Arguments:=FormatMessageAPILastArgument)
    If FormatMessageResult > 0 Then
        ErrorText = TrimToNull(ErrorText)
        GetSystemErrorMessageText = ErrorText
    Else
        GetSystemErrorMessageText = "NO ERROR DESCRIPTION AVAILABLE"
    End If
End Function

Function TrimToNull(TEXT As String) As String
    Rem  Returns all the text in Text to the left of the vbNullChar
    Dim NullCharIndex As Integer
    NullCharIndex = InStr(1, TEXT, vbNullChar, vbTextCompare)
    If NullCharIndex > 0 Then
        TrimToNull = left(TEXT, NullCharIndex - 1)
    Else
        TrimToNull = TEXT
    End If
End Function

Rem UserformMinimize
Sub AddMinimizeButtonToUserform(form As Object)
    Dim UserFormCaption As String
    UserFormCaption = form.Caption
    Dim hWnd            As Long
    Dim exLong          As Long
    hWnd = FindWindowA(vbNullString, UserFormCaption)
    exLong = GetWindowLongA(hWnd, -16)
    If (exLong And &H20000) = 0 Then
        SetWindowLongA hWnd, -16, exLong Or &H20000
    Else
    End If
End Sub

Sub UserformSetHandCursor(Optional form As Object)
    '#INCLUDE SetHandCursor
    '#INCLUDE ActiveModule
    If form Is Nothing Then
        Dim Module As VBComponent
        Set Module = ActiveModule
        If Module.Type = vbext_ct_MSForm Then
            Dim ctr As MSForms.control
            For Each ctr In Module.Designer.Controls
                SetHandCursor ctr
            Next
        End If
    End If
End Sub

Sub UserformSelectedControlsSetHandCursor()
    '#INCLUDE SetHandCursor
    '#INCLUDE SelectedControls
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    If Module.Type = vbext_ct_MSForm Then
        Dim ctr As MSForms.control
        For Each ctr In SelectedControls
            SetHandCursor ctr
        Next
    End If
End Sub

Sub SetHandCursor(control As MSForms.control)
    On Error GoTo catch
    With control
        .MouseIcon = LoadPicture("C:\Users\acer\Dropbox\SOFTWARE\EXCEL\0 Alex\icons\Hand Cursor Pointer.ico")
        .MousePointer = fmMousePointerCustom
    End With
catch:
End Sub

Sub SwitchControlNames()
    '#INCLUDE SelectedControls
    Dim ctrls As Collection
    Set ctrls = SelectedControls
    If ctrls.count <> 2 Then Exit Sub
    Dim tmp1 As String
    tmp1 = ctrls(1).Name
    Dim tmp2 As String
    tmp2 = ctrls(2).Name
    ctrls(1).Name = "tmp1"
    ctrls(2).Name = "tmp2"
    ctrls(1).Name = tmp2
    ctrls(2).Name = tmp1
End Sub

Sub SwitchControlPositions()
    '#INCLUDE SelectedControls
    Dim ctrls As Collection
    Set ctrls = SelectedControls
    If ctrls.count <> 2 Then Exit Sub
    Dim left1 As Long, left2 As Long
    Dim top1 As Long, top2 As Long
    left1 = ctrls(1).left
    top1 = ctrls(1).top
    left2 = ctrls(2).left
    top2 = ctrls(2).top
    ctrls(1).left = left2
    ctrls(1).top = top2
    ctrls(2).left = left1
    ctrls(2).top = top1
End Sub

Public Sub Reframe(form As Object, control As MSForms.control)
    Dim c As MSForms.control
    For Each c In form.Controls
        If TypeName(c) = "Frame" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                If c.Name <> control.parent.parent.Name Then c.visible = False
            End If
        End If
    Next
    form.Controls(control.Caption).visible = True
    For Each c In form.Controls
        If TypeName(c) = "Label" Then
            If Not InStr(1, c.Tag, "skip", vbTextCompare) > 0 Then
                c.BackColor = &H534848
            End If
        End If
    Next
    control.BackColor = &H80B91E
End Sub

Sub SaveUserformOptions(form As Object, _
                        Optional includeCheckbox As Boolean = True, _
                        Optional includeOptionButton As Boolean = True, _
                        Optional includeTextBox As Boolean = True, _
                        Optional includeListbox As Boolean = True, _
                        Optional includeToggleButton As Boolean = True)
    '#INCLUDE ListboxSelectedIndexes
    '#INCLUDE CreateOrSetSheet
    '#INCLUDE CollectionToArray
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(form.Name & "_Settings", ThisWorkbook)
    ws.Cells.clear
    Dim coll As New Collection
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.control
    For Each c In form.Controls
        If TypeName(c) Like "CheckBox" Then
            If Not includeCheckbox Then GoTo Skip
        ElseIf TypeName(c) Like "OptionButton" Then
            If Not includeOptionButton Then GoTo Skip
        ElseIf TypeName(c) Like "TextBox" Then
            If Not includeTextBox Then GoTo Skip
        ElseIf TypeName(c) = "ListBox" Then
            If Not includeListbox Then GoTo Skip
        ElseIf TypeName(c) Like "ToggleButton" Then
            If Not includeToggleButton Then GoTo Skip
        Else
            GoTo Skip
        End If
        cell = c.Name
        Select Case TypeName(c)
            Case "TextBox", "CheckBox", "OptionButton", "ToggleButton"
                cell.OFFSET(0, 1) = c.Value
            Case "ListBox"
                Set coll = ListboxSelectedIndexes(c)
                If coll.count > 0 Then
                    cell.OFFSET(0, 1) = Join(CollectionToArray(coll), ",")
                Else
                    cell.OFFSET(0, 1) = -1
                End If
        End Select
        Set cell = cell.OFFSET(1, 0)
Skip:
    Next
End Sub

Sub ListboxToRangeSelect(lBox As MSForms.ListBox)
    '#INCLUDE ListboxSelectedValues
    '#INCLUDE GetInputRange
    '#INCLUDE CollectionsToArrayTable
    Dim rng As Range
    If GetInputRange(rng, "Range picker", "Select range to output listbox' list") = False Then Exit Sub
    rng.RESIZE(lBox.ListCount, lBox.columnCount) = CollectionsToArrayTable(ListboxSelectedValues(lBox))
End Sub

Sub LoadUserformOptions(form As Object, Optional ExcludeThese As Variant)
    '#INCLUDE SelectListboxItems
    '#INCLUDE IsInArray
    '#INCLUDE CreateOrSetSheet
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet(form.Name & "_Settings", ThisWorkbook)
    If ws.Range("A1") = "" Then Exit Sub
    Dim cell As Range
    Set cell = ws.Cells(1, 1)
    Dim c As MSForms.control
    Dim v
    On Error Resume Next
    Do While cell <> ""
        Set c = form.Controls(cell.TEXT)
        If Not TypeName(c) = "Nothing " Then
            If Not IsInArray(cell, ExcludeThese) Then
                Select Case TypeName(c)
                    Case "TextBox", "CheckBox", "OptionButton", "ToggleButton"
                        c.Value = cell.OFFSET(0, 1)
                    Case "ListBox"
                        If InStr(1, cell.OFFSET(0, 1), ",") > 0 Then
                            SelectListboxItems c, Split(cell.OFFSET(0, 1), ","), True
                        Else
                            c.SELECTED(CInt(cell.OFFSET(0, 1))) = True
                        End If
                End Select
            End If
        End If
        Set cell = cell.OFFSET(1, 0)
    Loop
End Sub

Sub AddFormControls(controlID As String, _
                    CountOrArrayOfNames As Variant, _
                    Optional Captions As Variant = 0, _
                    Optional Vertical As Boolean = True, _
                    Optional OFFSET As Long = 0, _
                    Optional form As Object)
    '#INCLUDE ActiveModule
    If IsNumeric(CountOrArrayOfNames) And IsArray(Captions) Then
        If UBound(Captions) + 1 <> CLng(CountOrArrayOfNames) Then Exit Sub
    End If
    Dim Module As VBComponent
    If form Is Nothing Then
        Set Module = ActiveModule
        If Module.Type <> vbext_ct_MSForm Then Exit Sub
    End If
    Dim c As MSForms.control
    Dim i As Long
    If IsNumeric(CountOrArrayOfNames) Then
        For i = 1 To CLng(CountOrArrayOfNames)
            If form Is Nothing Then
                Set c = Module.Designer.Controls.Add(controlID)
            Else
                Set c = form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.top = i * c.Height + i * 5 - c.Height
                c.left = OFFSET
            Else
                c.top = OFFSET
                c.left = i * c.Width + i * 5 - c.Width
            End If
            If IsArray(Captions) Then
                c.Caption = Captions(i - 1)
            Else
                On Error Resume Next
                c.Caption = CountOrArrayOfNames(i - 1)
                If c.Caption = "" Then c.Caption = c.Name
                On Error GoTo 0
            End If
        Next
    Else
        For i = 1 To UBound(CountOrArrayOfNames) + 1
            If form Is Nothing Then
                Set c = Module.Designer.Controls.Add(controlID)
            Else
                Set c = form.Controls.Add(controlID)
            End If
            If Vertical Then
                c.top = i * c.Height + i * 5 - c.Height
                c.left = OFFSET
            Else
                c.top = OFFSET
                c.left = i * c.Width + i * 5 - c.Width
            End If
            c.Name = CountOrArrayOfNames(i - 1)
            If IsArray(Captions) Then
                c.Caption = Captions(i - 1)
            Else
                On Error Resume Next
                c.Caption = CountOrArrayOfNames(i - 1)
                If c.Caption = "" Then c.Caption = c.Name
                On Error GoTo 0
            End If
        Next
    End If
End Sub

Sub AddMultipleControls(ControlTypes As Variant, count As Long, Optional Vertical As Boolean = True, Optional form As Object = Nothing)
    '#INCLUDE AddFormControls
    '#INCLUDE ActiveModule
    Dim i As Long
    For i = 1 To UBound(ControlTypes) + 1
        If Vertical Then
            AddFormControls CStr(ControlTypes(i - 1)), count, , Vertical, i * 60 - 50, form
        Else
            AddFormControls CStr(ControlTypes(i - 1)), count, , Vertical, i * 20 - 20, form
        End If
    Next
    Dim c As MSForms.control
    On Error Resume Next
    If form Is Nothing Then
        For Each c In ActiveModule.Designer.Controls
            If Not TypeName(c) Like "TextBox" Then c.AutoSize = True
        Next
    Else
        For Each c In form.Controls
            If Not TypeName(c) Like "TextBox" Then c.AutoSize = True
        Next
    End If
End Sub

Sub EditObjectProperties(obj As Variant, PropertyArguement As Variant)
    Rem EditObjectProperties SelectedControl, Array("left",0,"top",0)
    Rem For Each c In SelectedControls: EditObjectProperties c, Array("left",0,"top",0): Next
Rem for i=1 to SelectedControls.count: EditObjectProperties activemodule.Designer.controls(SelectedControls(i).name),Array("left",0,"top",0): next
'#INCLUDE SelectedControl
'#INCLUDE SelectedControls
'#INCLUDE ActiveModule
Dim i As Long
Do While i < UBound(PropertyArguement)
CallByName obj, PropertyArguement(i), VbLet, _
        IIf(IsNumeric(PropertyArguement(i + 1)), _
            CLng(PropertyArguement(i + 1)), _
            PropertyArguement(i + 1))
i = i + 2
Loop
End Sub

Sub EditObjectsProperty(obj As Collection, objProperty As String, Args As Variant)
    If obj.count <> UBound(Args) + 1 Then
        MsgBox "Not selected controls count <> arguements count"
        Exit Sub
    End If
    Dim ArgItem
    Dim i As Long
    i = obj.count
    Dim element As Variant
    For Each element In obj
        CallByName element, objProperty, VbLet, _
                   IIf(IsNumeric(Args(i - 1)), _
                       CLng(Args(i - 1)), _
                       Args(i - 1))
        i = i - 1
    Next
End Sub

Sub RenameControlAndCode(Optional ctr As MSForms.control)
    '#INCLUDE InStrExact
    '#INCLUDE SelectedControl
    '#INCLUDE SelectedControls
    '#INCLUDE ActiveModule
    '#INCLUDE InputboxString
    If ctr Is Nothing Then
        If SelectedControls.count = 1 Then Set ctr = SelectedControl
        If ctr Is Nothing Then
            MsgBox "No control passed as arguement or no 1 control selected in designer"
            Exit Sub
        End If
    End If
    Dim Module As VBComponent: Set Module = ActiveModule
    If Module.Type <> vbext_ct_MSForm Then Exit Sub
    Dim OldName As String: OldName = ctr.Name
    Dim NewName As String: NewName = InputboxString
    If NewName = "" Then Exit Sub
    ctr.Name = NewName
    Dim CountOfLines As Long: CountOfLines = Module.CodeModule.CountOfLines
    If CountOfLines = 0 Then Exit Sub
    Dim strLine As String
    Dim i As Long
    Rem @TODO this part is wrong
    For i = 1 To CountOfLines
        strLine = Module.CodeModule.Lines(i, 1)
        If InStr(1, strLine, " " & OldName & "_") > 0 Then
            If InStrExact(1, strLine, OldName & "_") > 0 Then
                Module.CodeModule.ReplaceLine (i), Replace(strLine, OldName, NewName & "_")
            End If
        End If
    Next
End Sub

Sub SortControlsHorizontally()
    '#INCLUDE SortControls
    SortControls False
End Sub

Sub SortControlsVertivally()
    '#INCLUDE SortControls
    SortControls True
End Sub

Sub SortControls(Optional SortVertically As Boolean = True)
    Rem call from immediate window while looking at userform
    '#INCLUDE SelectedControls
    '#INCLUDE ActiveModule
    '#INCLUDE SortCollection
    Dim Module As VBComponent
    Set Module = ActiveModule
    If Module.Type <> vbext_ct_MSForm Then Exit Sub
    Dim ctr As MSForms.control
    Dim coll As New Collection
    Dim lastTop As Long
    Dim lastLeft As Long
    Dim element As Variant
    For Each element In SelectedControls
        coll.Add element.Name
    Next
    Set coll = SortCollection(coll)
    lastTop = 2000
    For Each element In coll
        If Module.Designer.Controls(element).top < lastTop Then lastTop = Module.Designer.Controls(element).top
        If Module.Designer.Controls(element).left < lastLeft Then lastLeft = Module.Designer.Controls(element).left
    Next
    For Each element In coll
        If SortVertically = True Then
            lastTop = lastTop + Module.Designer.Controls(element).Height + 6
        Else
            lastLeft = lastLeft + Module.Designer.Controls(element).Width + 6
        End If
        Module.Designer.Controls(element).top = lastTop
        Module.Designer.Controls(element).left = lastLeft
    Next
End Sub

Sub CopyControlProperties(Optional control As MSForms.control)
    '#INCLUDE SelectedControl
    '#INCLUDE CreateOrSetSheet
    '#INCLUDE Min
    If control Is Nothing Then Set control = SelectedControl
    Dim ws As Worksheet: Set ws = CreateOrSetSheet("CopyControlProperties", ThisWorkbook)
    Dim PropertiesArray As Variant
    PropertiesArray = Array("Accelerator", "Alignment", "AutoSize", "AutoTab", "BackColor", "BackStyle", "BorderColor", "BorderStyle", "BoundColumn", _
                            "Caption", "Children", "columnCount", "ColumnHeads", "ColumnWidths", "ControlSource", "ControlTipText", "Cycle", "DrawBuffer", "Enabled", "EnterKeyBehavior", "Expanded", _
                            "FirstSibling", "FontBold", "FontSize", "ForeColor", "FullPath", "GroupName", "Height", "HelpContextID", "KeepScrollBarsVisible", "LargeChange", "LastSibling", "LineStyle", "ListRows", "Locked", _
                            "Max", "MaxLength", "Min", "MouseIcon", "MousePointer", "MultiLine", "MultiSelect", "Next", "Nodes", "Orientation", _
                            "Parent", "PasswordChar", "PathSeparator", "Picture", "PictureAlignment", "PictureSizeMode", "PictureTiling", "Previous", "RightToLeft", "Root", "RowSource", _
                            "ScrollBars", "ScrollHeight", "ScrollLeft", "ScrollTop", "ScrollWidth", "Selected", "SelectedItem", "ShowModal", "SmallChange", "Sorted", "SpecialEffect", "StartUpPosition", _
                            "Style", "Tag", "Text", "TextColumn", "TripleState", "WhatsThisHelp", "Width", "Zoom")
    If ws.Range("A1") = "" Then ws.Range("A1").RESIZE(UBound(PropertiesArray) + 1) = WorksheetFunction.Transpose(PropertiesArray)
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.RESIZE(, 1)
    Dim property As Range
    On Error Resume Next
    For Each property In PropertiesRange
        property.OFFSET(0, 1) = CallByName(control, property.Value, VbGet)
    Next
End Sub

Sub PasteControlProperties(Optional Controls As Collection)
    '#INCLUDE CopyControlProperties
    '#INCLUDE SelectedControls
    Dim control As MSForms.control
    If Controls Is Nothing Then Set Controls = SelectedControls
    Dim ws As Worksheet: Set ws = ThisWorkbook.SHEETS("CopyControlProperties")
    If ws.Columns(2).SpecialCells(xlCellTypeConstants).count = 0 Then
        MsgBox "You haven't saved properties before"
        Exit Sub
    End If
    Dim PropertiesRange As Range: Set PropertiesRange = ws.Range("A1").CurrentRegion.RESIZE(, 1)
    Dim property As Range
    On Error Resume Next
    For Each control In Controls
        For Each property In PropertiesRange
            CallByName control, property.Value, VbLet, property.OFFSET(0, 1).Value
        Next
    Next
End Sub

Function SelectedControl() As MSForms.control
    '#INCLUDE SelectedControls
    '#INCLUDE ActiveModule
    Dim Module As VBComponent
    Set Module = ActiveModule
    If SelectedControls.count = 1 Then
        Dim ctl    As control
        For Each ctl In ActiveModule.Designer.SELECTED
            Set SelectedControl = ctl
            Exit Function
        Next ctl
    End If
End Function

Function SelectedControls() As Collection
    '#INCLUDE ActiveModule
    Dim ctl    As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    For Each ctl In Module.Designer.SELECTED
        out.Add ctl
    Next ctl
    Set SelectedControls = out
    Set out = Nothing
End Function

Sub RemoveControlsCaptions()
    '#INCLUDE SelectedControls
    Dim c As MSForms.control
    For Each c In SelectedControls
        c.Caption = ""
    Next
End Sub

Function SelectedFrameControls() As Collection
    '#INCLUDE ActiveModule
    Dim ctl    As control, c As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    For Each ctl In Module.Designer.SELECTED
        For Each c In ctl.Controls
            out.Add c
        Next
    Next ctl
    Set SelectedFrameControls = out
    Set out = Nothing
End Function

Function SelectedFrameControl() As MSForms.control
    '#INCLUDE ActiveModule
    Dim ctl    As control, c As control
    Dim out As New Collection
    Dim Module As VBComponent
    Set Module = ActiveModule
    For Each ctl In Module.Designer.SELECTED
        For Each c In ctl.Controls
            out.Add c
        Next
    Next ctl
    If out.count = 0 Then Exit Function
    Set SelectedFrameControl = out(1)
End Function

Sub SelectListboxItems(lBox As MSForms.ListBox, FindMe As Variant, Optional ByIndex As Boolean)
    Dim i As Long
    Select Case TypeName(FindMe)
        Case Is = "String", "Long", "Integer"
            For i = 0 To lBox.ListCount - 1
                If lBox.list(i) = CStr(FindMe) Then
                    lBox.SELECTED(i) = True
                    DoEvents
                    If lBox.multiSelect = fmMultiSelectSingle Then Exit Sub
                End If
            Next
        Case Else
            Dim el As Variant
            If ByIndex Then
                For Each el In FindMe
                    lBox.SELECTED(el) = True
                Next
            Else
                For Each el In FindMe
                    For i = 0 To lBox.ListCount - 1
                        If lBox.list(i) = el Then
                            lBox.SELECTED(i) = True
                            DoEvents
                        End If
                    Next
                Next
            End If
    End Select
End Sub

Sub CreateListboxHeader(body As MSForms.ListBox, header As MSForms.ListBox, arrHeaders)
    '#INCLUDE ArrayDimensions
    header.Width = body.Width
    Dim i As Long
    header.columnCount = body.columnCount
    header.ColumnWidths = body.ColumnWidths
    header.clear
    header.AddItem
    If ArrayDimensions(arrHeaders) = 1 Then
        For i = 0 To UBound(arrHeaders)
            header.list(0, i) = arrHeaders(i)
        Next i
    Else
        For i = 1 To UBound(arrHeaders, 2)
            header.list(0, i - 1) = arrHeaders(1, i)
        Next i
    End If
    body.ZOrder (1)
    header.ZOrder (0)
    header.SpecialEffect = fmSpecialEffectFlat
    header.BackColor = RGB(200, 200, 200)
    header.Height = 15
    header.Width = body.Width
    header.left = body.left
    header.top = body.top - header.Height - 1
    header.Font.Bold = True
    header.Font.Name = "Comic Sans MS"
    header.Font.Size = 9
End Sub

Sub SavePosition(form As Object)
    SaveSetting "My Settings Folder", form.Name, "Left Position", form.left
    SaveSetting "My Settings Folder", form.Name, "Top Position", form.top
End Sub

Sub LoadPosition(form As Object)
    If GetSetting("My Settings Folder", form.Name, "Left Position") = "" _
                                                                      And GetSetting("My Settings Folder", form.Name, "Top Position") = "" Then
        form.StartUpPosition = 1
    Else
        form.left = GetSetting("My Settings Folder", form.Name, "Left Position")
        form.top = GetSetting("My Settings Folder", form.Name, "Top Position")
    End If
End Sub

Rem
Sub ResizeUserformToFitControls(form As Object)
    form.Width = 0
    form.Height = 0
    Dim ctr As MSForms.control
    Dim myWidth
    myWidth = form.InsideWidth
    Dim myHeight
    myHeight = form.InsideHeight
    For Each ctr In form.Controls
        If ctr.visible = True Then
            If ctr.left + ctr.Width > myWidth Then myWidth = ctr.left + ctr.Width
            If ctr.top + ctr.Height > myHeight Then myHeight = ctr.top + ctr.Height
        End If
    Next
    form.Width = myWidth + form.Width - form.InsideWidth + 10
    form.Height = myHeight + form.Height - form.InsideHeight + 10
End Sub

Function whichOption(Frame As Variant, controlType As String) As Variant
    Dim subControl As MSForms.control
    Dim out As New Collection
    Dim control As MSForms.control
    For Each control In Frame.Controls
        If UCase(TypeName(control)) = UCase("Frame") Then
            If UCase(TypeName(control)) = UCase(controlType) Then
                If control.Value = True Then
                    out.Add control
                End If
            End If
        End If
    Next
    If out.count = 1 Then
        whichOption = out(1)
    ElseIf out.count > 1 Then
        Set whichOption = out
    End If
End Function

Rem Control
Public Sub flashControl(ctr As MSForms.control, blinkCount As Integer)
    Rem if blinkCount = odd then the control will become hidden
    Dim lngTime As Long
    Dim i As Integer
    If blinkCount Mod 2 <> 0 Then blinkCount = blinkCount + 1
    For i = 1 To blinkCount * 2
        lngTime = getTickCount
        If ctr.visible = True Then
            ctr.visible = False
        Else
            ctr.visible = True
        End If
        DoEvents
        Do While getTickCount - lngTime < 200
        Loop
    Next
End Sub

Public Function TextOfControl(c As control) As Variant
    Rem Text of Textbox, Selection of Combobox, Selected items (2d) of Listbox
    '#INCLUDE ListboxSelectedValues
    '#INCLUDE CollectionToArray
    Dim out As New Collection
    If TypeName(c) = "TextBox" Then
        If c.SelLength = 0 Then
            TextOfControl = c.TEXT
        Else
            TextOfControl = c.SelText
        End If
    ElseIf TypeName(c) = "ComboBox" Then
        If c.Style < 2 Then
            TextOfControl = c.TEXT
        Else
            TextOfControl = ""
        End If
    ElseIf TypeName(c) = "ListBox" Then
        Set out = ListboxSelectedValues(c)
        If out.count > 0 Then
            TextOfControl = CollectionToArray(out)
        Else
            TextOfControl = ""
        End If
    End If
End Function

Rem Listbox
Public Function ListboxContains(lBox As MSForms.ListBox, str As String, _
                                Optional ColumnIndexZeroBased As Long = -1, _
                                Optional CaseSensitive As Boolean = False) As Boolean
    Dim i      As Long
    Dim n      As Long
    Dim sTemp  As String
    If ColumnIndexZeroBased > lBox.columnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = lBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = lBox.list(i, ColumnIndexZeroBased)
            Else
                str = LCase(str)
                sTemp = LCase(lBox.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, str) > 0 Then
                ListboxContains = True
                Exit Function
            End If
        Next i
    Else
        Dim columnCount As Long
        n = lBox.ListCount
        For i = n - 1 To 0 Step -1
            For columnCount = 0 To lBox.columnCount - 1
                If CaseSensitive = True Then
                    sTemp = lBox.list(i, columnCount)
                Else
                    str = LCase(str)
                    sTemp = LCase(lBox.list(i, columnCount))
                End If
                If InStr(1, sTemp, str) > 0 Then
                    ListboxContains = True
                    Exit Function
                End If
            Next columnCount
        Next i
    End If
End Function

Public Sub FilterListboxByColumn(lBox As MSForms.ListBox, str As String, _
                                 Optional ColumnIndexZeroBased As Long = -1, Optional CaseSensitive As Boolean = False)
    Dim i               As Long
    Dim n               As Long
    Dim sTemp           As String
    If ColumnIndexZeroBased > lBox.columnCount - 1 Or ColumnIndexZeroBased < 0 Then
        ColumnIndexZeroBased = -1
    End If
    n = lBox.ListCount
    If ColumnIndexZeroBased <> -1 Then
        For i = n - 1 To 0 Step -1
            If CaseSensitive = True Then
                sTemp = lBox.list(i, ColumnIndexZeroBased)
            Else
                str = LCase(str)
                sTemp = LCase(lBox.list(i, ColumnIndexZeroBased))
            End If
            If InStr(1, sTemp, str) = 0 Then
                lBox.RemoveItem (i)
            End If
        Next i
    Else
        Dim columnCount As Long
        n = lBox.ListCount
        For i = n - 1 To 0 Step -1
            For columnCount = 0 To lBox.columnCount - 1
                If CaseSensitive = True Then
                    sTemp = lBox.list(i, columnCount)
                Else
                    str = LCase(str)
                    sTemp = LCase(lBox.list(i, columnCount))
                End If
                If InStr(1, sTemp, str) > 0 Then
                Else
                    If columnCount = lBox.columnCount - 1 Then
                        lBox.RemoveItem (i)
                    End If
                End If
            Next columnCount
        Next i
    End If
End Sub

Public Sub SortListboxOnColumn(lBox As MSForms.ListBox, OnColumn As Long)
    Dim vntData As Variant
    Dim vntTempItem As Variant
    Dim lngOuterIndex As Long
    Dim lngInnerIndex As Long
    Dim lngSubItemIndex As Long
    vntData = lBox.list
    For lngOuterIndex = LBound(vntData, 1) To UBound(vntData, 1) - 1
        For lngInnerIndex = lngOuterIndex + 1 To UBound(vntData, 1)
            If vntData(lngOuterIndex, OnColumn) > vntData(lngInnerIndex, OnColumn) Then
                For lngSubItemIndex = 0 To lBox.columnCount - 1
                    vntTempItem = vntData(lngOuterIndex, lngSubItemIndex)
                    vntData(lngOuterIndex, lngSubItemIndex) = vntData(lngInnerIndex, lngSubItemIndex)
                    vntData(lngInnerIndex, lngSubItemIndex) = vntTempItem
                Next
            End If
        Next lngInnerIndex
    Next lngOuterIndex
    lBox.clear
    lBox.list = vntData
End Sub

Function ListboxSelectedIndexes(lBox As MSForms.ListBox) As Collection
    Dim i As Long
    Dim SelectedIndexes As Collection
    Set SelectedIndexes = New Collection
    If lBox.ListCount > 0 Then
        For i = 0 To lBox.ListCount - 1
            If lBox.SELECTED(i) Then SelectedIndexes.Add i
        Next i
    End If
    Set ListboxSelectedIndexes = SelectedIndexes
End Function

Function ListboxSelectedValues(listboxCollection As Variant) As Collection
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    If TypeName(listboxCollection) = "Collection" Then
        For listboxCount = 1 To listboxCollection.count
            If listboxCollection(listboxCount).ListCount > 0 Then
                For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                    If listboxCollection(listboxCount).SELECTED(listItem) Then
                        selectedCollection.Add CStr(listboxCollection(listboxCount).list(listItem, listboxCollection(listboxCount).BoundColumn - 1))
                    End If
                Next listItem
            End If
        Next listboxCount
    Else
        If listboxCollection.ListCount > 0 Then
            For i = 0 To listboxCollection.ListCount - 1
                If listboxCollection.SELECTED(i) Then
                    selectedCollection.Add listboxCollection.list(i, listboxCollection.BoundColumn - 1)
                End If
            Next i
        End If
    End If
    Set ListboxSelectedValues = selectedCollection
End Function

Function ListboxSelectedCount(listboxCollection As Variant) As Long
    Dim i As Long
    Dim listItem As Long
    Dim selectedCollection As Collection
    Set selectedCollection = New Collection
    Dim listboxCount As Long
    Dim SelectedCount As Long
    If TypeName(listboxCollection) = "Collection" Then
        For listboxCount = 1 To listboxCollection.count
            If listboxCollection(listboxCount).ListCount > 0 Then
                For listItem = 0 To listboxCollection(listboxCount).ListCount - 1
                    If listboxCollection(listboxCount).SELECTED(listItem) = True Then
                        SelectedCount = SelectedCount + 1
                    End If
                Next listItem
            End If
        Next listboxCount
    Else
        If listboxCollection.ListCount > 0 Then
            For i = 0 To listboxCollection.ListCount - 1
                If listboxCollection.SELECTED(i) = True Then
                    SelectedCount = SelectedCount + 1
                End If
            Next i
        End If
    End If
    ListboxSelectedCount = SelectedCount
End Function

Rem var
Public Sub ShowUserform(FormName As String)
    '#INCLUDE IsLoaded
    Dim frm As Object
    If IsLoaded(FormName) = True Then
        For Each frm In VBA.UserForms
            If frm.Name = FormName Then
                frm.Show
                Exit Sub
            End If
        Next frm
    Else
        Dim oUserForm As Object
        On Error GoTo err
        Set oUserForm = UserForms.Add(FormName)
        oUserForm.Show (vbModeless)
        Exit Sub
    End If
err:
    Select Case err.Number
        Case 424:
            MsgBox "The Userform with the name " & FormName & " was not found.", vbExclamation, "Load userforn by name"
        Case Else:
            MsgBox err.Number & ": " & err.Description, vbCritical, "Load userforn by name"
    End Select
End Sub

Sub ResizeControlColumns(ListboxOrCombobox As MSForms.control, Optional ResizeControl As Boolean, Optional ResizeListbox As Boolean)
    '#INCLUDE CreateOrSetSheet
    If ListboxOrCombobox.ListCount = 0 Then Exit Sub
    Application.ScreenUpdating = False
    Dim ws As Worksheet
    Set ws = CreateOrSetSheet("ListboxColumnwidth", ThisWorkbook)
    Dim rng As Range
    Set rng = ws.Range("A1")
    Set rng = rng.RESIZE(UBound(ListboxOrCombobox.list) + 1, ListboxOrCombobox.columnCount)
    rng = ListboxOrCombobox.list
    rng.Font.Name = ListboxOrCombobox.Font.Name
    rng.Font.Size = ListboxOrCombobox.Font.Size + 2
    rng.Columns.AutoFit
    Dim sWidth As String
    Dim vR() As Variant
    Dim n As Integer
    Dim cell As Range
    For Each cell In rng.RESIZE(1)
        n = n + 1
        ReDim Preserve vR(1 To n)
        vR(n) = cell.EntireColumn.Width
    Next cell
    sWidth = Join(vR, ";")
    With ListboxOrCombobox
        .ColumnWidths = sWidth
        .BorderStyle = fmBorderStyleSingle
    End With
    Application.DisplayAlerts = False
    ws.Delete
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    If ResizeListbox = False Then Exit Sub
    Dim w As Long
    Dim i As Long
    For i = LBound(vR) To UBound(vR)
        w = w + vR(i)
    Next
    DoEvents
    ListboxOrCombobox.Width = w + 10
End Sub

Sub DeselectListbox(lBox As MSForms.ListBox)
    If lBox.ListCount <> 0 Then
        Dim i As Long
        For i = 0 To lBox.ListCount - 1
            lBox.SELECTED(i) = False
        Next i
    End If
End Sub

Public Sub SelectDeselectAll(lBox As MSForms.ListBox, Optional toSelect As Boolean)
    If lBox.ListCount = 0 Then Exit Sub
    Dim i As Long
    For i = 0 To lBox.ListCount - 1
        lBox.SELECTED(i) = toSelect
    Next
End Sub

Sub SelectControItemsByFilter(lBox As MSForms.ListBox, criteria As String)
    '#INCLUDE SelectDeselectAll
    SelectDeselectAll lBox
    If criteria = "" Then Exit Sub
    Dim i As Long
    For i = 0 To lBox.ListCount - 1
        If UCase(lBox.list(i, 1)) Like "*" & UCase(criteria) & "*" Then
            lBox.SELECTED(i) = True
        End If
    Next i
End Sub


