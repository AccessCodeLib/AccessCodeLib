Attribute VB_Name = "modWinAPI_Layout"
Attribute VB_Description = "WinAPI-Funktionen zur Layoutgestaltung"
'---------------------------------------------------------------------------------------
' Package: api.winapi.modWinAPI_Layout
'---------------------------------------------------------------------------------------
'
' WinAPI functions for layout
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Layout.bas</file>
'  <license>_codelib/license.bas</license>
'  <description>Standard module version: set window back color, convert twips/pixels, get scrollbar width; prefer WinApiLayoutTools.cls</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Declare PtrSafe Function CreateSolidBrush Lib "gdi32.dll" ( _
      ByVal CrColor As Long _
) As LongPtr

Private Declare PtrSafe Function RedrawWindow Lib "user32" ( _
      ByVal Hwnd As LongPtr, _
      LprcUpdate As Any, _
      ByVal HrgnUpdate As Long, _
      ByVal FuRedraw As Long _
) As Long

Private Declare PtrSafe Function SetClassLong _
      Lib "user32.dll" _
      Alias "SetClassLongA" ( _
      ByVal Hwnd As LongPtr, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As LongPtr _
      ) As Long

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

'--------------------------------------
Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare PtrSafe Function GetDC Lib "user32" (ByVal Hwnd As LongPtr) As LongPtr
Private Declare PtrSafe Function ReleaseDC Lib "user32" (ByVal Hwnd As LongPtr, ByVal hDC As LongPtr) As Long
Private Declare PtrSafe Function GetDeviceCaps Lib "gdi32" (ByVal hDC As LongPtr, ByVal nIndex As Long) As LongPtr
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXVSCROLL As Long = 2


'---------------------------------------------------------------------------------------
' Sub: SetBackColor
'---------------------------------------------------------------------------------------
'
' Set the background color of a window
'
' Parameters:
'     Hwnd  - Window handle
'     Color - Color number
'
'---------------------------------------------------------------------------------------
Public Sub SetBackColor(ByVal Hwnd As LongPtr, ByVal Color As Long)
  
   Dim NewBrush As LongPtr
   
   'Brush erzeugen
   NewBrush = CreateSolidBrush(Color)
   'Brush zuweisen
   SetClassLong Hwnd, GCL_HBRBACKGROUND, NewBrush
   'Fenster neuzeichnen (gesamtes Fenster inkl. Background)
   RedrawWindow Hwnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

End Sub

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelX
'---------------------------------------------------------------------------------------
'
' Width of a pixel in twips
'
' Returns:
'     Single
'
' Remarks:
'     http://support.microsoft.com/kb/94927/de
'
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelX() As Single
   Dim lngDC As LongPtr
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelY
'---------------------------------------------------------------------------------------
'
' Height of a pixel in twips
'
' Returns:
'     Single
'
' Remarks:
'     http://support.microsoft.com/kb/94927/de
'
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelY() As Single
   Dim lngDC As LongPtr
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: GetScrollbarWidth
'---------------------------------------------------------------------------------------
'
' Width of the scroll bar
'
' Returns:
'     Single
'
'---------------------------------------------------------------------------------------
Public Function GetScrollbarWidth() As Single
   GetScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * TwipsPerPixelX
End Function

'---------------------------------------------------------------------------------------
' Function: GetTwipsFromPixel
'---------------------------------------------------------------------------------------
'
' Rechnet Pixel in Twips um
'
' Parameters:
'     pixel - Number of pixels
'
' Returns:
'     Long
'
' Remarks:
'     GetTwipsFromPixel = TwipsPerPixelX * pixel
'
'---------------------------------------------------------------------------------------
Public Function GetTwipsFromPixel(ByVal Pixel As Long) As Long
   GetTwipsFromPixel = TwipsPerPixelX * Pixel
End Function

'---------------------------------------------------------------------------------------
' Function: GetPixelFromTwips
'---------------------------------------------------------------------------------------
'
' Rechnet twips in Pixel um
'
' Parameters:
'     twips - Number of twips
'
' Returns:
'     Long
'
' Remarks:
'     GetPixelFromTwips = twips / TwipsPerPixelX
'
'---------------------------------------------------------------------------------------
Public Function GetPixelFromTwips(ByVal Twips As Long) As Long
        GetPixelFromTwips = Twips / TwipsPerPixelX
End Function
