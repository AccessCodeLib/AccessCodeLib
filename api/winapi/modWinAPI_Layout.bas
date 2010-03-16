Attribute VB_Name = "modWinAPI_Layout"
Attribute VB_Description = "WinAPI-Funktionen zur Layoutgestaltung"
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
'  <file>api/winapi/modWinAPI_Layout.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'/** \addtogroup WinAPI
'@{ **/

Private Declare Function CreateSolidBrush _
      Lib "gdi32.dll" ( _
      ByVal crColor As Long _
      ) As Long

Private Declare Function RedrawWindow _
      Lib "user32" ( _
      ByVal hWnd As Long, _
      lprcUpdate As Any, _
      ByVal hrgnUpdate As Long, _
      ByVal fuRedraw As Long _
      ) As Long

Private Declare Function SetClassLong _
      Lib "USER32.DLL" _
      Alias "SetClassLongA" ( _
      ByVal hWnd As Long, _
      ByVal nIndex As Long, _
      ByVal dwNewLong As Long _
      ) As Long

Private Const GCL_HBRBACKGROUND As Long = -10
Private Const RDW_INVALIDATE As Long = &H1
Private Const RDW_ERASE As Long = &H4

'--------------------------------------
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const HWND_DESKTOP As Long = 0
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90

Private Const SM_CXVSCROLL As Long = 2

'----------------------------------------------------------------------------------------
Public Sub SetBackColor(ByVal H As Long, ByVal Color As Long)
  
   Dim NewBrush As Long
   
   'Brush erzeugen
On Error GoTo HandleErr

   NewBrush = CreateSolidBrush(Color)
   'Brush zuweisen
   SetClassLong H, GCL_HBRBACKGROUND, NewBrush
   'Fenster neuzeichnen (gesamtes Fenster inkl. Background)
   RedrawWindow H, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_ERASE

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "SetBackColor", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

'----------------------------------------------------------------------------------------
' http://support.microsoft.com/kb/94927/de
'
'--------------------------------------------------
Public Function TwipsPerPixelX() As Single
'--------------------------------------------------
'Returns the width of a pixel, in twips.
'--------------------------------------------------

   Dim lngDC As Long
   
On Error GoTo HandleErr

   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelX = 1440& / GetDeviceCaps(lngDC, LOGPIXELSX)
   ReleaseDC HWND_DESKTOP, lngDC

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "TwipsPerPixelX", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Function

'--------------------------------------------------
Public Function TwipsPerPixelY() As Single
'--------------------------------------------------
'Returns the height of a pixel, in twips.
'--------------------------------------------------

   Dim lngDC As Long
On Error GoTo HandleErr

   lngDC = GetDC(HWND_DESKTOP)
   TwipsPerPixelY = 1440& / GetDeviceCaps(lngDC, LOGPIXELSY)
   ReleaseDC HWND_DESKTOP, lngDC

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "TwipsPerPixelY", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function
      

Public Function GetScrollbarWidth() As Single

On Error GoTo HandleErr

   GetScrollbarWidth = GetSystemMetrics(SM_CXVSCROLL) * TwipsPerPixelX

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetScrollbarWidth", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function



Public Function GetTwipsFromPixel(ByVal pixel As Long) As Long

On Error GoTo HandleErr

   GetTwipsFromPixel = TwipsPerPixelX * pixel

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetTwipsFromPixel", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function GetPixelFromTwips(ByVal twips As Long) As Long

On Error GoTo HandleErr

GetPixelFromTwips = twips / TwipsPerPixelX

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetPixelFromTwips", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
