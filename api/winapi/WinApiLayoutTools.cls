VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinApiLayoutTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module: modWinAPI_Layout
'---------------------------------------------------------------------------------------
'/**
' <summary>
' WinAPI-Funktionen zur Layoutgestaltung
' </summary>
' <remarks>
' </remarks>
'\ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/WinApiLayoutTools.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Declare Function CreateSolidBrush _
      Lib "gdi32.dll" ( _
      ByVal crColor As Long _
      ) As Long

Private Declare Function RedrawWindow _
      Lib "user32" ( _
      ByVal Hwnd As Long, _
      lprcUpdate As Any, _
      ByVal hrgnUpdate As Long, _
      ByVal fuRedraw As Long _
      ) As Long

Private Declare Function SetClassLong _
      Lib "USER32.DLL" _
      Alias "SetClassLongA" ( _
      ByVal Hwnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long _
      ) As Long

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

'--------------------------------------
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal Hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal Hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXVSCROLL As Long = 2


'---------------------------------------------------------------------------------------
' Sub: SetBackColor (Josef Pötzl, 2010-04-19)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hintergrundfarbe eines Fensters einstellen
' </summary>
' <param name="Hwnd">Fenster-Handle</param>
' <param name="Color">Farbnummer</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetBackColor(ByVal Hwnd As Long, ByVal Color As Long)
  
   Dim NewBrush As Long
   
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
'/**
' <summary>
' Breite eines Pixels in twips
' </summary>
' <param name="Param"></param>
' <returns>Single</returns>
' <remarks>
' http://support.microsoft.com/kb/94927/de
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelX() As Single
   Dim lngDC As Long
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: TwipsPerPixelY
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Höhe eines Pixels in twips
' </summary>
' <returns>Single</returns>
' <remarks>
' http://support.microsoft.com/kb/94927/de
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TwipsPerPixelY() As Single
   Dim lngDC As Long
   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, lngDC
End Function

'---------------------------------------------------------------------------------------
' Function: GetScrollbarWidth
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Breite der Bildlaufleiste
' </summary>
' <param name="Param"></param>
' <returns>Single</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetScrollbarWidth() As Single
   GetScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * TwipsPerPixelX
End Function

'---------------------------------------------------------------------------------------
' Function: GetTwipsFromPixel
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Rechnet Pixel in Twips um
' </summary>
' <param name="pixel">Anzahl der Pixel</param>
' <returns>Long</returns>
' <remarks>
' GetTwipsFromPixel = TwipsPerPixelX * pixel
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetTwipsFromPixel(ByVal pixel As Long) As Long
   GetTwipsFromPixel = TwipsPerPixelX * pixel
End Function

'---------------------------------------------------------------------------------------
' Function: GetPixelFromTwips
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Rechnet twips in Pixel um
' </summary>
' <param name="twips">Anzahl twips</param>
' <returns>Long</returns>
' <remarks>
'  GetPixelFromTwips = twips / TwipsPerPixelX
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetPixelFromTwips(ByVal twips As Long) As Long
        GetPixelFromTwips = twips / TwipsPerPixelX
End Function
