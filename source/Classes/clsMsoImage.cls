VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsMsoImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public WithEvents LabelEvents As MSForms.label
Attribute LabelEvents.VB_VarHelpID = -1

Private Sub LabelEvents_Click()
    uImageMso.TextBox1.TEXT = LabelEvents.Caption
    uImageMso.TextBox1.SelStart = 0
    uImageMso.TextBox1.SelLength = Len(uImageMso.TextBox1.TEXT)
End Sub

