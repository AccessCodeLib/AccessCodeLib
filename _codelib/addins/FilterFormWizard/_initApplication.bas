Attribute VB_Name = "_initApplication"
'---------------------------------------------------------------------------------------
' Modul: _initApplication
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Initialisierungsaufruf der Anwendung
' </summary>
' <remarks>
' </remarks>
' \ingroup base
' @todo StartApplication-Prozedur für allgemeine Verwendung umschreiben => in Klasse verlagern
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/_initApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>_codelib/addins/FilterFormWizard/defGlobal_FilterFormWizard.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'-------------------------
' Anwendungseinstellungen
'-------------------------
'
' => siehe _config_Application
'
'-------------------------

'---------------------------------------------------------------------------------------
' Function: StartApplication (Josef Pötzl, 2009-12-14)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prozedur für den Anwendungsstart
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function StartApplication(Optional ByRef param As Variant) As Boolean

On Error GoTo HandleErr

   StartApplication = CurrentApplication.Start

ExitHere:
   Exit Function

HandleErr:
   StartApplication = False
   MsgBox "Anwendung kann nicht gestartet werden.", vbCritical, CurrentApplicationName
   Application.Quit acQuitSaveNone
   Resume ExitHere

End Function


Public Sub RestoreApplicationDefaultSettings()
   On Error Resume Next
   CurrentApplication.ApplicationTitle = CurrentApplication.ApplicationFullName
End Sub
