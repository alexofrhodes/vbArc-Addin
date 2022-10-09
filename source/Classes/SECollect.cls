VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SECollect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private LBCol As Collection
Private SuperExtendLB As SExtendLB

Private Sub Class_Initialize()
    Set LBCol = New Collection
End Sub

Public Sub Add(ByVal NewLB As MSForms.ListBox)
    Set SuperExtendLB = New SExtendLB
    SuperExtendLB.Add NewLB
    LBCol.Add SuperExtendLB
End Sub

