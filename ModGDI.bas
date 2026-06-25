Attribute VB_Name = "ModGDI"
Option Explicit

Private Type GdiplusStartupInput
  GdiplusVersion As Long
  DebugEventCallback As Long
  SuppressBackgroundThread As Long
  SuppressExternalCodecs As Long
End Type



Private Enum ImageLockMode
  ImageLockModeRead = &H1
  ImageLockModeWrite = &H2
  ImageLockModeUserInputBuf = &H4
End Enum

Private Type BitmapData
  Width As Long
  Height As Long
  Stride As Long
  PixelFormat As Long
  scan0 As Long
  reserved As Long
End Type



Private Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type



Private Const UnitPixel As Long = &H2&



Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipLoadImageFromFile Lib "gdiplus" (ByVal FileName As Long, hImage As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "gdiplus" (ByVal hdc As Long, Graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "gdiplus" (ByVal Graphics As Long) As Long
Private Declare Function GdipDrawImageRectRectI Lib "gdiplus" (ByVal Graphics As Long, ByVal hImage As Long, ByVal dstX As Long, ByVal dstY As Long, ByVal dstWidth As Long, ByVal dstHeight As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, ByVal srcUnit As Long, Optional ByVal imageAttributes As Long = 0, Optional ByVal callback As Long = 0, Optional ByVal callbackData As Long = 0) As Long
Private Declare Function GdipGetImageWidth Lib "gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "gdiplus" (ByVal Image As Long, Height As Long) As Long



Public Sub ShowPNG(PicSource As PictureBox, PNGFile As String)
  Dim Image As Long
  Dim Graphics As Long
  Dim Token As Long
  Dim GdipInput As GdiplusStartupInput
  Dim Width As Long, Height As Long
  GdipInput.GdiplusVersion = 1
  GdiplusStartup Token, GdipInput
  GdipLoadImageFromFile StrPtr(PNGFile), Image
  GdipGetImageWidth Image, Width
  GdipGetImageHeight Image, Height
  GdipCreateFromHDC PicSource.hdc, Graphics
  GdipDrawImageRectRectI Graphics, Image, 0, 0, Width, Height, 0, 0, Width, Height, UnitPixel, 0, 0, 0
  GdipDeleteGraphics Graphics
  GdipDisposeImage Image
  PicSource.Refresh
  PicSource.Picture = PicSource.Image
  GdiplusShutdown Token
End Sub
