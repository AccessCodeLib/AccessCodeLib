Attribute VB_Name = "modWinApi_Mouse"
'---------------------------------------------------------------------------------------
' Module: modWinApi_Mouse (Josef Pötzl, 2010-03-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Maus-Zeiger einstellen
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Mouse.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Enum IDC_MouseCursor
   IDC_HAND = 32649&
   IDC_APPSTARTING = 32650&
   IDC_ARROW = 32512&
   IDC_CROSS = 32515&
   IDC_IBEAM = 32513&
   IDC_ICON = 32641&
   IDC_SIZE = 32640&
   IDC_SIZEALL = 32646&
   IDC_SIZENESW = 32643&
   IDC_SIZENS = 32645&
   IDC_SIZENWSE = 32642&
   IDC_SIZEWE = 32644&
   IDC_UPARROW = 32516&
   IDC_WAIT = 32514&
   IDC_NO = 32648&
End Enum

Private Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

'---------------------------------------------------------------------------------------
' Sub: MouseCursor (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Mauszeiger einstellen
' </summary>
' <param name="CursorType">Gewünschter Mauszeiger</param>
' <returns>Variant</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub MouseCursor(ByVal CursorType As IDC_MouseCursor)
  Dim lngRet As Long
  lngRet = LoadCursorBynum(0&, CursorType)
  lngRet = SetCursor(lngRet)
End Sub
