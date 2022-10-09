VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStackItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Class module cStackItem
'User-defined type for a stack item

Option Explicit

Public baseCell As Range
Public rowOffset As Long
#If VBA7 Then
    Public hwndParent As LongPtr
#Else
    Public hwndParent As Long
#End If
