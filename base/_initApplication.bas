Attribute VB_Name = "_initApplication"
'---------------------------------------------------------------------------------------
' Modul: _initApplication
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Initialising the application
' </summary>
' <remarks>
' </remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/_initApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'-------------------------
' Anwendungseinstellungen
'-------------------------
'
' => see _config_Application
'
'-------------------------

'---------------------------------------------------------------------------------------
' Function: StartApplication
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Procedure for application start-up
' </summary>
' <returns>Boolean (sucess = true)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function StartApplication() As Boolean

On Error GoTo HandleErr

   StartApplication = CurrentApplication.Start

ExitHere:
   Exit Function

HandleErr:
   StartApplication = False
   MsgBox "Application can not be started.", vbCritical, CurrentApplicationName
   Application.Quit acQuitSaveNone
   Resume ExitHere

End Function

Public Sub RestoreApplicationDefaultSettings()
   On Error Resume Next
   CurrentApplication.ApplicationTitle = CurrentApplication.ApplicationFullName
End Sub
