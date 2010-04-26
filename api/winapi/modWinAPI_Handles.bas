Attribute VB_Name = "modWinAPI_Handles"
Attribute VB_Description = "Win-API-Funktionen: Handles"
'---------------------------------------------------------------------------------------
' Module: modWinAPI_Handles
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Win-API-Funktionen: Handles
' </summary>
' <remarks>
' </remarks>
' \ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Handles.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Die Prozeudren (GetMDI, GetHeaderSection, GetDetailSection, GetFooterSection und GetControl
' stammen aus dem AEK10-Vortrag von Jörg Ostendorp
'
'----------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Type POINTAPI
   x As Long
   Y As Long
End Type

Private Declare Function ClientToScreen Lib "USER32.DLL" ( _
         ByVal Hwnd As Long, _
         ByRef lpPoint As POINTAPI _
      ) As Long

Private Declare Function FindWindowEx Lib "USER32.DLL" Alias "FindWindowExA" ( _
         ByVal hWnd1 As Long, _
         ByVal hWnd2 As Long, _
         ByVal lpsz1 As String, _
         ByVal lpsz2 As String _
      ) As Long

'---------------------------------------------------------------------------------------
' Function: GetMDI
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle des MDI-Client-Fensters.
' </summary>
' <returns>Handle (Long)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetMDI() As Long

   'Ermittelt den Handle des MDI-Client-Fensters.

   Dim H As Long
On Error GoTo HandleErr

   H = Application.hWndAccessApp

   'Erstes (und einziges) "MDIClient"-Kindfenster des Applikationsfensters suchen
   GetMDI = FindWindowEx(H, 0&, "MDIClient", vbNullString)


ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetMDI", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetHeaderSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle für den Kopfbereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetHeaderSection(ByVal fHwnd As Long) As Long

   'Ermittelt den Handle für den Kopfbereich eines Formulares

   Dim H As Long

   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln
On Error GoTo HandleErr

   H = FindWindowEx(fHwnd, 0&, "OformSub", vbNullString)
   GetHeaderSection = H

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetHeaderSection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetDetailSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle für den Detailbereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDetailSection(ByVal fHwnd As Long) As Long

   'Ermittelt den Handle für den Detailbereich eines Formulares

   Dim H As Long

   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Kopfbereich
On Error GoTo HandleErr

   H = GetHeaderSection(fHwnd)
   H = FindWindowEx(fHwnd, H, "OformSub", vbNullString)
   GetDetailSection = H

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetDetailSection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetFooterSection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle für den Fußbereich eines Formulares
' </summary>
' <param name="fHwnd">Handle des Formulars (Form.Hwnd)</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFooterSection(ByVal fHwnd As Long) As Long

   'Ermittelt den Handle für den Fußbereich eines Formulares

   Dim H As Long

   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Detailbereich
On Error GoTo HandleErr

   H = GetDetailSection(fHwnd)
   H = FindWindowEx(fHwnd, H, "OformSub", vbNullString)
   GetFooterSection = H

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetFooterSection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetControl
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt den Handle eines beliebigen Controls
' </summary>
' <param name="frm">Formular-Referenz</param>
' <param name="sHwnd">Handle des Bereichs, auf dem sich das Control befindet (Header, Detail, Footer)</param>
' <param name="ClassName">Name der Fensterklasse des Controls</param>
' <param name="ControlName">Name des Controls</param>
' <returns>Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetControl(ByRef frm As Access.Form, ByVal sHwnd As Long, ByVal ClassName As String, ByVal ControlName As String) As Long

   'Ermittelt den Handle eines beliebigen Controls

   'Parameter:
   ' frm - Formular
   ' Handle des Bereichs, auf dem sich das Control befindet (Header, Detail, Footer)
   ' ControName - Name der Fensterklasse des Controls
   ' ControlName - Name des Controls


   'Exitieren mehrere Controls der gleichen Klasse auf einem Formular, z.B. TabControls, besteht das Problem, daß
   'deren Reihenfolge nicht definiert ist (anders also als bei den Sektionsfenstern)
   'In diesem Fall kann man alle Kindfenster dieser Klasse in einer Schleife durchlaufen
   'und z.B. prüfen, ob die Position des Fensters des zurückgegebenen Handles
   'mit der des Access-Steuerelementes übereinstimmt.
   'Nachfolgend wird hierfür die undokumentierte Funktion accHittest verwendet.
   'Dieser werden als Parameter die Screenkoordinaten der linken oberen Ecke eines
   'Steuerelementes übergeben. Befindet sich dort ein Objekt, erhält man dieses als Rückgabewert.
   'Ist der Name des Objektes identisch mit dem übergebenen Steuerelementnamen, so
   'hat man das Handle ermittelt:

On Error Resume Next

   Dim H As Long
   Dim obj As Object
   Dim pt As POINTAPI

   H = 0

   Do
      'Erstes (h=0)/nächstes (h<>0) Control auf dem Sektionsfenster ermitteln
      H = FindWindowEx(sHwnd, H, ClassName, vbNullString)

      'Bildschirmkoordinaten dieses Controls ermitteln
      'dafür die Punktkoordinaten aus dem letzten Durchlauf zurücksetzen, sonst wird addiert!
      pt.x = 0
      pt.Y = 0
      ClientToScreen H, pt

      'Objekt bei den Koordinaten ermitteln
      Set obj = frm.accHitTest(pt.x, pt.Y)

      'Wenn Objektname = Tabname Ausstieg aus der Schleife
      If obj.Name = ControlName Then
         Exit Do
      End If
   Loop While H <> 0

   'Handle zurückgeben
   GetControl = H

End Function
