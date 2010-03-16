Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _config_Application (Beispiel)
'---------------------------------------------------------------------------------------
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/samples/_config_Application_DirTextbox.bas</file>
'  <replace>base/_config_Application.bas</replace>
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>usability/ApplicationHandler_DirTextbox.cls</use>
'  <use>_test/usability/TEST_DirTextbox.frm</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Das Modul _config_Application wird vom Import-Assistenden nicht überschrieben.
' Sollte eine neues _config_Application-Modul geladen werden,
' ist das alte zuvor umzubennen oder zu löschen.
'
'
Option Compare Database
Option Explicit

Private Const m_ApplicationName As String = "Access Code Library: DirTextbox"
Private Const m_ApplicationFullName As String = m_ApplicationName
Private Const m_ApplicationTitle As String = m_ApplicationName

Private Const m_DefaultErrorHandlerMode As Long = ACLibErrorHandlerMode.aclibErrMsgBox

'---------------------------------------------------------------------------------------
' Sub: InitConfig
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen initialisieren
' </summary>
' <param name="oCurrentAppHandler">Möglichkeit einer Referenzübergabe, damit nicht CurrentApplication genutzt werden muss</param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub InitConfig(Optional ByRef oCurrentAppHandler As ApplicationHandler = Nothing)

'----------------------------------------------------------------------------
' Fehlerbehandlung
'
On Error GoTo HandleErr

   modErrorHandler.DefaultErrorHandlerMode = m_DefaultErrorHandlerMode

'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   
   
'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If oCurrentAppHandler Is Nothing Then
      Set oCurrentAppHandler = CurrentApplication
   End If

   With oCurrentAppHandler
   
      'Anwendungsname
      .ApplicationName = m_ApplicationName
      .ApplicationFullName = m_ApplicationFullName
      
      'Titelleiste der Anwendung
      .ApplicationTitle = m_ApplicationTitle
      
      ' Formular, das am Ende von StartApplication aufgerufen wird
      '.ApplicationStartFormName =

   End With
   
   
'----------------------------------------------------------------------------
' Erweiterung: ...
'
   AddApplicationHandlerExtension New ApplicationHandler_DirTextbox

'----------------------------------------------------------------------------
' Konfiguration nach Initialisierung der Erweiterungen
'

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "InitConfig", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
