Attribute VB_Name = "F_Modifier"
Rem @Folder Modifier References

'Declared elsewhere
'Declare PtrSafe Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Long

Public Enum VirtualKey
    VK_LBUTTON = &H1
    VK_SHIFT = &H10
    VK_CONTROL = &H11
    VK_MENU = &H12
    VK_PAUSE = &H13
    VK_CAPITAL = &H14
    VK_ESCAPE = &H1B
    VK_RBUTTON = &H2
    VK_SPACE = &H20
    VK_PRIOR = &H21
    VK_NEXT = &H22
    VK_END = &H23
    VK_HOME = &H24
    VK_LEFT = &H25
    VK_UP = &H26
    VK_RIGHT = &H27
    VK_DOWN = &H28
    VK_SELECT = &H29
    VK_PRINT = &H2A
    VK_EXECUTE = &H2B
    VK_SNAPSHOT = &H2C
    VK_INSERT = &H2D
    VK_DELETE = &H2E
    VK_HELP = &H2F
    VK_CANCEL = &H3
    VK_MBUTTON = &H4        ' NOT contiguous with L RBUTTON
    VK_NUMPAD0 = &H60
    VK_NUMPAD1 = &H61
    VK_NUMPAD2 = &H62
    VK_NUMPAD3 = &H63
    VK_NUMPAD4 = &H64
    VK_NUMPAD5 = &H65
    VK_NUMPAD6 = &H66
    VK_NUMPAD7 = &H67
    VK_NUMPAD8 = &H68
    VK_NUMPAD9 = &H69
    VK_MULTIPLY = &H6A
    VK_ADD = &H6B
    VK_SEPARATOR = &H6C
    VK_SUBTRACT = &H6D
    VK_DECIMAL = &H6E
    VK_DIVIDE = &H6F
    VK_F1 = &H70
    VK_F2 = &H71
    VK_F3 = &H72
    VK_F4 = &H73
    VK_F5 = &H74
    VK_F6 = &H75
    VK_F7 = &H76
    VK_F8 = &H77
    VK_F9 = &H78
    VK_F10 = &H79
    VK_F11 = &H7A
    VK_F12 = &H7B
    VK_F13 = &H7C
    VK_F14 = &H7D
    VK_F15 = &H7E
    VK_F16 = &H7F
    VK_BACK = &H8
    VK_F17 = &H80
    VK_F18 = &H81
    VK_F19 = &H82
    VK_F20 = &H83
    VK_F21 = &H84
    VK_F22 = &H85
    VK_F23 = &H86
    VK_F24 = &H87
    VK_TAB = &H9
    VK_NUMLOCK = &H90
    VK_SCROLL = &H91
    VK_LSHIFT = &HA0
    VK_RSHIFT = &HA1
    VK_LCONTROL = &HA2
    VK_RCONTROL = &HA3
    VK_LMENU = &HA4
    VK_RMENU = &HA5
    VK_CLEAR = &HC
    VK_RETURN = &HD
    VK_PROCESSKEY = &HE5
    VK_ATTN = &HF6
    VK_CRSEL = &HF7
    VK_EXSEL = &HF8
    VK_EREOF = &HF9
    VK_PLAY = &HFA
    VK_ZOOM = &HFB
    VK_NONAME = &HFC
    VK_PA1 = &HFD
    VK_OEM_CLEAR = &HFE
End Enum

Rem @Folder Modifier

Sub is_correct_key_pressed()
    If GetAsyncKeyState(VirtualKey.VK_SHIFT) <> 0 Then
        MsgBox "Shift was pressed"
    Else
        MsgBox "Shift was NOT pressed"
    End If
End Sub

