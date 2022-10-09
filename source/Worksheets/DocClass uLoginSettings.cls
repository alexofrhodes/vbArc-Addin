VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet27"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    If Not Intersect(Target, Range("H5:Z24")) Is Nothing Then
        If Target.Value = [B1] Then
            Target.Value = [b2]
            Cancel = True
            Exit Sub
        End If
        If Target.Value = [b3] Then Target.Value = [B1]
        If Target.Value = [b2] Then Target.Value = [b3]
        If Target.Value = Empty Then Target.Value = [B1]
    End If
    If Not Intersect(Target, Range("G5:G24")) Is Nothing Then
        If Target.Value = [B1] Then
            Range("G" & Target.row & ":M" & Target.row).Value = [b2]
            Cancel = True
            Exit Sub
        End If
        If Target.Value = [b3] Then Range("G" & Target.row & ":M" & Target.row).Value = [B1]
        If Target.Value = [b2] Then Range("G" & Target.row & ":M" & Target.row).Value = [b3]
        If Target.Value = Empty Then Range("G" & Target.row & ":M" & Target.row).Value = [B1]
    End If
    Cancel = True
End Sub

