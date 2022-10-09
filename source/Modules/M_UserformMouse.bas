Attribute VB_Name = "M_UserformMouse"

Rem @Folder FormMouse
Option Explicit
Option Compare Text
Rem in userform:    ShowFormAtCursor me
Public Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function GetCursorPos Lib "user32" (p As tCursor) As Long
Public Declare PtrSafe Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Public Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Public Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare PtrSafe Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As tCursor) As Long
Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Const LOGPIXELSX = 88
Public Const LOGPIXELSY = 90
Public Type tCursor
    left As Long
    top As Long
End Type

Sub LeftClick()
    mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
    mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
End Sub

Rem in userform activated event
Rem CenterMouseOver Me, ComboBox1
Public Sub CenterMouseOver(f As Object, c As Object)
    '#INCLUDE PointsPerPixelX
    '#INCLUDE PointsPerPixelY
    Dim p As tCursor
    Dim lngHwnd As Long
    lngHwnd = CLng(FindWindow(vbNullString, f.Caption))
    p.left = (c.left + (c.Width \ 2)) / PointsPerPixelX
    p.top = (c.top + (c.Height \ 2)) / PointsPerPixelY
    ClientToScreen lngHwnd, p
    SetCursorPos p.left, p.top
End Sub

Public Function PointsPerPixelX() As Double
    Dim hdc As Long
    hdc = GetDC(0)
    PointsPerPixelX = 72 / GetDeviceCaps(hdc, LOGPIXELSX)
    ReleaseDC 0, hdc
End Function

Public Function PointsPerPixelY() As Double
    Dim hdc As Long
    hdc = GetDC(0)
    PointsPerPixelY = 72 / GetDeviceCaps(hdc, LOGPIXELSY)
    ReleaseDC 0, hdc
End Function

Public Function WhereIsTheMouseAt() As tCursor
    Dim mPos As tCursor
    GetCursorPos mPos
    WhereIsTheMouseAt = mPos
End Function

Public Function convertMouseToForm() As tCursor
    '#INCLUDE PointsPerPixelX
    '#INCLUDE PointsPerPixelY
    '#INCLUDE WhereIsTheMouseAt
    Dim mPos As tCursor
    mPos = WhereIsTheMouseAt
    mPos.left = PointsPerPixelY * mPos.left
    mPos.top = PointsPerPixelX * mPos.top
    convertMouseToForm = mPos
End Function

Sub ShowFormAtCursor(form As Object)
    '#INCLUDE convertMouseToForm
    form.left = convertMouseToForm.left
    form.top = convertMouseToForm.top
End Sub

