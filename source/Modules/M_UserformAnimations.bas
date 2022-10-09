Attribute VB_Name = "M_UserformAnimations"
Rem @Folder UserformAnimations Declarations
Rem Author:Todar
Option Explicit
Option Compare Text
Option Private Module

#If VBA7 And Win64 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If
#If VBA7 Then

    Private Declare PtrSafe Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare PtrSafe Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#Else
    Private Declare Function getFrequency Lib "kernel32" Alias "QueryPerformanceFrequency" (cyFrequency As Currency) As Long
    Private Declare Function getTickCount Lib "kernel32" Alias "QueryPerformanceCounter" (cyTickCount As Currency) As Long
#End If
Rem Effect(box, "Top", Me.InsideHeight - box.Height, 1000) _
, Effect(box2, "Top", 0, 100) _
, Effect(GoButton, "fontsize", 12, 1000) _
, Effect(Me, "Top", 20, 2000)

Rem @Folder UserformAnimations
Public Sub Transition(ParamArray Elements() As Variant)
    '#INCLUDE MicroTimer
    '#INCLUDE AllTransitionsComplete
    '#INCLUDE IncrementElement
    If IsArray(Elements(LBound(Elements, 1))) Then
        Dim temp As Variant
        temp = Elements(LBound(Elements, 1))
        Elements = temp
    End If
    Dim form As MSForms.UserForm
    Set form = Elements(LBound(Elements, 1))("form")
    MicroTimer True
    Do
        Dim index As Integer
        For index = LBound(Elements, 1) To UBound(Elements, 1)
            IncrementElement Elements(index), MicroTimer
        Next index
        Sleep 40
        form.Repaint
    Loop Until AllTransitionsComplete(CVar(Elements))
End Sub

Public Function Effect(obj As Object, property As String, Destination As Double, MilSecs As Double) As Scripting.Dictionary
    Dim temp As New Scripting.Dictionary
    Set temp("obj") = obj
    temp("property") = property
    temp("startValue") = CallByName(obj, property, VbGet)
    temp("destination") = Destination
    temp("travel") = Destination - temp("startValue")
    temp("milSec") = MilSecs
    temp("complete") = False
    On Error GoTo catch:
    Set temp("form") = obj.parent
    Set Effect = temp
    Exit Function
catch:
    Set temp("form") = obj
    Resume Next
End Function

Public Function MicroTimer(Optional StartTime As Boolean = False) As Double
    Static dTime As Double
    Dim cyTicks1 As Currency
    Dim cyTicks2 As Currency
    Static cyFrequency As Currency
    MicroTimer = 0
    If cyFrequency = 0 Then getFrequency cyFrequency
    getTickCount cyTicks1
    getTickCount cyTicks2
    If cyTicks2 < cyTicks1 Then cyTicks2 = cyTicks1
    If cyFrequency Then MicroTimer = cyTicks2 / cyFrequency
    If StartTime = True Then
        dTime = MicroTimer
        MicroTimer = 0
    Else
        MicroTimer = (MicroTimer - dTime) * 1000
    End If
End Function

Private Function AllTransitionsComplete(Elements As Variant) As Boolean
    '#INCLUDE TransitionComplete
    Dim el As Object
    Dim index As Integer
    For index = LBound(Elements, 1) To UBound(Elements, 1)
        Set el = Elements(index)
        If Not TransitionComplete(el) Then
            AllTransitionsComplete = False
            Exit Function
        End If
    Next index
    AllTransitionsComplete = True
End Function

Private Function TransitionComplete(ByVal el As Scripting.Dictionary) As Boolean
    If Math.Round(el("destination")) = Math.Round(CallByName(el("obj"), el("property"), VbGet)) Then
        TransitionComplete = True
    End If
End Function

Private Function IncrementElement(ByVal el As Scripting.Dictionary, CurrentTime As Double) As Boolean
    '#INCLUDE TransitionComplete
    '#INCLUDE easeInAndOut
    Dim IncrementValue As Double
    Dim CurrentValue As Double
    If TransitionComplete(el) Then
        Exit Function
    End If
    Dim o As Object
    Dim p As String
    Dim Value As Double
    Dim d As Double
    IncrementValue = easeInAndOut(CurrentTime, el("startValue"), el("travel"), el("milSec"))
    If el("travel") < 0 Then
        If Math.Round(IncrementValue, 4) < el("destination") Then
            CallByName el("obj"), el("property"), VbLet, el("destination")
        Else
            CallByName el("obj"), el("property"), VbLet, IncrementValue
        End If
    Else
        If Math.Round(IncrementValue, 4) > el("destination") Then
            CallByName el("obj"), el("property"), VbLet, el("destination")
        Else
            CallByName el("obj"), el("property"), VbLet, IncrementValue
        End If
    End If
End Function

Private Function easeInAndOut(ByVal t As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double) As Double
    d = d / 2
    t = t / d
    If (t < 1) Then
        easeInAndOut = c / 2 * t * t * t + b
    Else
        t = t - 2
        easeInAndOut = c / 2 * (t * t * t + 2) + b
    End If
End Function


