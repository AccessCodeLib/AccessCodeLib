Attribute VB_Name = "modWinAPI_Handles"
Attribute VB_Description = "Win-API-Funktionen: Handles"
'---------------------------------------------------------------------------------------
' Package: api.winapi.modWinAPI_Handles
'---------------------------------------------------------------------------------------
'
' Win API functions: handles
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Handles.bas</file>
'  <license>_codelib/license.bas</license>
'  <description>Standard module version: window handles of the MDI client, form sections and controls; prefer WinApiWindowTools.cls</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Die Prozeudren (GetMDI, GetHeaderSection, GetDetailSection, GetFooterSection und GetControl
' stammen aus dem AEK10-Vortrag von Jörg Ostendorp
'
'----------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Private Declare PtrSafe Function ClientToScreen Lib "user32.dll" ( _
         ByVal Hwnd As LongPtr, _
         ByRef lpPoint As POINTAPI _
      ) As Long

Private Declare PtrSafe Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" ( _
         ByVal HWnd1 As LongPtr, _
         ByVal HWnd2 As LongPtr, _
         ByVal Lpsz1 As String, _
         ByVal Lpsz2 As String _
      ) As LongPtr

'---------------------------------------------------------------------------------------
' Function: GetMDI
'---------------------------------------------------------------------------------------
'
' Determines the handle of the MDI client window.
'
' Returns:
'     Handle (Long)
'
'---------------------------------------------------------------------------------------
Public Function GetMDI() As LongPtr
   Dim h As LongPtr
   h = Application.hWndAccessApp
   'Erstes (und einziges) "MDIClient"-Kindfenster des Applikationsfensters suchen
   GetMDI = FindWindowEx(h, 0&, "MDIClient", vbNullString)
End Function

'---------------------------------------------------------------------------------------
' Function: GetHeaderSection
'---------------------------------------------------------------------------------------
'
' Determines the handle of the header section of a form
'
' Parameters:
'     fHwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function GetHeaderSection(ByVal fHwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln
   h = FindWindowEx(fHwnd, 0&, "OformSub", vbNullString)
   GetHeaderSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetDetailSection
'---------------------------------------------------------------------------------------
'
' Determines the handle of the detail section of a form
'
' Parameters:
'     fHwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function GetDetailSection(ByVal fHwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Kopfbereich
   h = GetHeaderSection(fHwnd)
   h = FindWindowEx(fHwnd, h, "OformSub", vbNullString)
   GetDetailSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetFooterSection
'---------------------------------------------------------------------------------------
'
' Determines the handle of the footer section of a form
'
' Parameters:
'     fHwnd - Handle of the form (Form.Hwnd)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function GetFooterSection(ByVal fHwnd As LongPtr) As LongPtr
   Dim h As LongPtr
   'Erstes "OFormsub"-Kindfenster des Formulares (fhwnd) ermitteln, beginnend
   'nach dem Detailbereich
   h = GetDetailSection(fHwnd)
   h = FindWindowEx(fHwnd, h, "OformSub", vbNullString)
   GetFooterSection = h
End Function

'---------------------------------------------------------------------------------------
' Function: GetControl
'---------------------------------------------------------------------------------------
'
' Determines the handle of any control
'
' Parameters:
'     frm         - Form reference
'     sHwnd       - Handle of the section containing the control (header, detail, footer)
'     sClassName  - Name of the window class of the control
'     ControlName - Name of the control
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function GetControl(ByRef frm As Access.Form, ByVal sHwnd As LongPtr, _
                           ByVal sClassName As String, ByVal ControlName As String) As LongPtr

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

   Dim h As LongPtr
   Dim obj As Object
   Dim pt As POINTAPI

   h = 0

   Do
      'Erstes (h=0)/nächstes (h<>0) Control auf dem Sektionsfenster ermitteln
      h = FindWindowEx(sHwnd, h, sClassName, vbNullString)

      'Bildschirmkoordinaten dieses Controls ermitteln
      'dafür die Punktkoordinaten aus dem letzten Durchlauf zurücksetzen, sonst wird addiert!
      pt.X = 0
      pt.Y = 0
      ClientToScreen h, pt

      'Objekt bei den Koordinaten ermitteln
      Set obj = frm.accHitTest(pt.X, pt.Y)

      'Wenn Objektname = Tabname Ausstieg aus der Schleife
      If obj.Name = ControlName Then
         Exit Do
      End If
   Loop While h <> 0

   'Handle zurückgeben
   GetControl = h

End Function
