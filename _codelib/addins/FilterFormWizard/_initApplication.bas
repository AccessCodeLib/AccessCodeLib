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


Private Sub DateienEinstellen()
   SaveModulesInTable
End Sub

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



Private Sub SaveModulesInTable()

   Dim x As Variant
   Dim i As Long
   
   x = Array("SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager")
   For i = 0 To UBound(x)
      SaveCodeModulInTable acModule, x(i)
   Next
   
End Sub

Private Sub SaveCodeModulInTable(ByVal ObjType As AcObjectType, ByVal sModulName As String)
   
   Dim strFileName As String

   strFileName = FileTools.GetNewTempFileName
   Application.SaveAsText ObjType, sModulName, strFileName
   CurrentApplication.SaveAppFile sModulName, strFileName, True
   Kill strFileName
   
End Sub
