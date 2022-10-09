Attribute VB_Name = "U_Pop"
Rem @Folder UserformPop
Sub Pop(Optional TextOrArray As Variant = "vbArc", _
        Optional SecondsPerMessage As Long = 5, _
        Optional ImagePath As String, _
        Optional TextSize As Long = 12, _
        Optional FontBold As Boolean = True, _
        Optional TextColor As Long = vbBlack, _
        Optional CounterColor As Long = vbBlack)
    Rem pop array("This is my home town, Rhodes","Have you been here?"),300,"C:\Users\acer\Pictures\sdafs.jpg",24,true,vbwhite
    uPop.Init TextOrArray, SecondsPerMessage, ImagePath, TextSize, FontBold, TextColor, CounterColor
End Sub


