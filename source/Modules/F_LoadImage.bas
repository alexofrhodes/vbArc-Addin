Attribute VB_Name = "F_LoadImage"
Rem This module provides a LoadPictureGDI function, which can
Rem be used instead of VBAs LoadPicture, to load a wide variety
Rem of image types from disk - including png.
Rem
Rem The png format is used in Office 2007-2010 to provide images that
Rem include an alpha channel for each pixels transparency
Rem
Rem Author:    Stephen Bullen
Rem Date:      31 October, 2006
Rem Email:     stephen@oaltd.co.uk
Rem
Rem Updated :  30 December, 2010
Rem By :       Rob Bovey
Rem Reason :   Also working now in the 64 bit version of Office 2010
Option Explicit
Rem Declare a UDT to store a GUID for the IPicture OLE Interface
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

#If VBA7 Then
    Rem Declare a UDT to store the bitmap information
Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As LongPtr
    hPal As LongPtr
End Type

Rem Declare a UDT to store the GDI+ Startup information
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Rem Windows API calls into the GDI+ library
Private Declare PtrSafe Function GdiplusStartup Lib "GDIPlus" (token As LongPtr, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As LongPtr = 0) As Long
Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal fileName As LongPtr, BITMAP As LongPtr) As Long
Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal BITMAP As LongPtr, hbmReturn As LongPtr, ByVal background As LongPtr) As Long
Private Declare PtrSafe Function GdipDisposeImage Lib "GDIPlus" (ByVal image As LongPtr) As Long
Private Declare PtrSafe Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As LongPtr)
Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, iPic As IPicture) As Long
#Else
Rem Declare a UDT to store the bitmap information
Private Type PICTDESC
    Size As Long
    Type As Long
    hPic As Long
    hPal As Long
End Type

Rem Declare a UDT to store the GDI+ Startup information
Private Type GDIPlusStartupInput
    GdiPlusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Rem Windows API calls into the GDI+ library
Private Declare Function GdiplusStartup Lib "GDIPlus" (token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Function GdipCreateBitmapFromFile Lib "GDIPlus" (ByVal filename As Long, bitmap As Long) As Long
Private Declare Function GdipCreateHBITMAPFromBitmap Lib "GDIPlus" (ByVal bitmap As Long, hbmReturn As Long, ByVal background As Long) As Long
Private Declare Function GdipDisposeImage Lib "GDIPlus" (ByVal image As Long) As Long
Private Declare Sub GdiplusShutdown Lib "GDIPlus" (ByVal token As Long)
Private Declare Function OleCreatePictureIndirect Lib "oleaut32" (PicDesc As PICTDESC, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
#End If

Rem  Procedure:    LoadPictureGDI
Rem  Purpose:      Loads an image using GDI+
Rem  Returns:      The image as an IPicture Object
Public Function LoadPictureGDI(ByVal sFilename As String) As IPicture
    '#INCLUDE CreateIPicture
    Dim uGdiInput As GDIPlusStartupInput
    Dim lResult As Long
    #If VBA7 Then
        Dim hGdiPlus As LongPtr
        Dim hGdiImage As LongPtr
        Dim hBitmap As LongPtr
    #Else
        Dim hGdiPlus As Long
        Dim hGdiImage As Long
        Dim hBitmap As Long
    #End If
    Rem Initialize GDI+
    uGdiInput.GdiPlusVersion = 1
    lResult = GdiplusStartup(hGdiPlus, uGdiInput)
    If lResult = 0 Then
        Rem Load the image
        lResult = GdipCreateBitmapFromFile(StrPtr(sFilename), hGdiImage)
        If lResult = 0 Then
            Rem Create a bitmap handle from the GDI image
            lResult = GdipCreateHBITMAPFromBitmap(hGdiImage, hBitmap, 0)
            Rem Create the IPicture object from the bitmap handle
            Set LoadPictureGDI = CreateIPicture(hBitmap)
            Rem Tidy up
            GdipDisposeImage hGdiImage
        End If
        Rem Shutdown GDI+
        GdiplusShutdown hGdiPlus
    End If
End Function

Rem  Procedure:    CreateIPicture
Rem  Purpose:      Converts a image handle into an IPicture object.
Rem  Returns:      The IPicture object
#If VBA7 Then
Private Function CreateIPicture(ByVal hPic As LongPtr) As IPicture
#Else
Private Function CreateIPicture(ByVal hPic As Long) As IPicture
#End If
Dim lResult As Long
Dim uPicinfo As PICTDESC
Dim IID_IDispatch As GUID
Dim iPic As IPicture
Rem OLE Picture types
Const PICTYPE_BITMAP = 1
Rem  Create the Interface GUID (for the IPicture interface)
With IID_IDispatch
    .Data1 = &H7BF80980
    .Data2 = &HBF32
    .Data3 = &H101A
    .Data4(0) = &H8B
    .Data4(1) = &HBB
    .Data4(2) = &H0
    .Data4(3) = &HAA
    .Data4(4) = &H0
    .Data4(5) = &H30
    .Data4(6) = &HC
    .Data4(7) = &HAB
End With
Rem  Fill uPicInfo with necessary parts.
With uPicinfo
    .Size = Len(uPicinfo)
    .Type = PICTYPE_BITMAP
    .hPic = hPic
    .hPal = 0
End With
Rem  Create the Picture object.
lResult = OleCreatePictureIndirect(uPicinfo, IID_IDispatch, True, iPic)
Rem  Return the new Picture object.
Set CreateIPicture = iPic
End Function

