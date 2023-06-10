Attribute VB_Name = "modWinApi_Mouse"
'---------------------------------------------------------------------------------------
' Package: api.winapi.modWinApi_Mouse
'---------------------------------------------------------------------------------------
'
' Set mouse cursor
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Mouse.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit
Option Private Module

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

#If VBA7 Then
   Private Declare PtrSafe Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As LongPtr, ByVal LpCursorName As Long) As LongPtr
   Private Declare PtrSafe Function SetCursor Lib "user32" (ByVal hCursor As LongPtr) As LongPtr
#Else
   Private Declare Function LoadCursorBynum Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal LpCursorName As Long) As Long
   Private Declare Function SetCursor Lib "user32" (ByVal Cursor As Long) As Long
#End If

'---------------------------------------------------------------------------------------
' Sub: MouseCursor
'---------------------------------------------------------------------------------------
'
' Set mouse cursor
'
' Parameters:
'     CursorType  - Desired mouse cursor
'
'---------------------------------------------------------------------------------------
Public Sub MouseCursor(ByVal CursorType As IDC_MouseCursor)
  
  Dim CursorPtr As LongPtr
  
  CursorPtr = LoadCursorBynum(0&, CursorType)
  SetCursor CursorPtr
  
End Sub
