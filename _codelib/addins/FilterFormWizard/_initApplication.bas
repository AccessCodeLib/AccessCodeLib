Attribute VB_Name = "_initApplication"
'---------------------------------------------------------------------------------------
' Package: _codelib.addins.FilterFormWizard._initApplication
'---------------------------------------------------------------------------------------
'
' Initialization call of the application
'
' Remarks:
'     TODO: Rewrite the StartApplication procedure for general use => move it into a class
'
'---------------------------------------------------------------------------------------

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
' Function: StartApplication
'---------------------------------------------------------------------------------------
'
' Procedure for the application start
'
' Parameters:
'     param - (optional) start parameter
'
' Returns:
'     Boolean
'
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
