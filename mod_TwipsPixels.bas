Attribute VB_Name = "mod_TwipsPixels"
Option Explicit

Public Enum Direction
    Horizontal
    Vertical
End Enum

Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Const WU_LOGPIXELSX = 88
Const WU_LOGPIXELSY = 90
Const TWIPSPERINCH = 1440

Public Function CPixel(ByVal Twips As Long, ByDirection As Direction) As Long
   Dim DCLong As Long: DCLong = GetDC(0)
   Dim PixelsPerInch As Long
   If (ByDirection = 0) Then PixelsPerInch = GetDeviceCaps(DCLong, WU_LOGPIXELSX)
   If (ByDirection = 1) Then PixelsPerInch = GetDeviceCaps(DCLong, WU_LOGPIXELSY)
   DCLong = ReleaseDC(0, DCLong)
   CPixel = (Twips / TWIPSPERINCH) * PixelsPerInch
End Function

Public Function CTwips(ByVal Pixels As Long, ByDirection As Direction) As Long
   Dim DCLong As Long: DCLong = GetDC(0)
   Dim PixelsPerInch As Long
   If (ByDirection = 0) Then PixelsPerInch = GetDeviceCaps(DCLong, WU_LOGPIXELSX)
   If (ByDirection = 1) Then PixelsPerInch = GetDeviceCaps(DCLong, WU_LOGPIXELSY)
   DCLong = ReleaseDC(0, DCLong)
   CTwips = Pixels * (TWIPSPERINCH / PixelsPerInch)
End Function
