Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/ImportWizard/_config_Application.bas</file>
'  <replace>base/_config_Application.bas</replace> 'dieses Modul ersetzt base/_config_Application.bas
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/ImportWizard/defGlobal_ACLibImportWizard.bas</use>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>_codelib/addins/shared/ACLibConfiguration.cls</use>
'  <use>_codelib/addins/ImportWizard/ACLibFileManager.cls</use>
'  <use>_codelib/addins/ImportWizard/ACLibImportWizardForm.frm</use>
'  <use>usability/ApplicationHandler_DirTextbox.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Versionsnummer
Private Const APPLICATION_VERSION As String = "1.0.10" '2015-06-15

#Const USE_CLASS_ApplicationHandler_AppFile = 1
#Const USE_CLASS_ApplicationHandler_DirTextbox = 1

Private Const APPLICATION_NAME As String = "ACLib Import Wizard"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - Import Wizard"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"

Private Const DEFAULT_ERRORHANDLERMODE As Long = ACLibErrorHandlerMode.aclibErrMsgBox

Private Const APPLICATION_STARTFORMNAME As String = "ACLibImportWizardForm"

'---------------------------------------------------------------------------------------
' Sub: InitConfig (Josef Pötzl, 2009-12-11)
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
Public Sub InitConfig(Optional ByRef CurrentAppHandlerRef As ApplicationHandler = Nothing)

On Error GoTo HandleErr

'----------------------------------------------------------------------------
' Fehlerbehandlung
'

   modErrorHandler.DefaultErrorHandlerMode = DEFAULT_ERRORHANDLERMODE

   
'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   defGlobal_ACLibImportWizard.ACLibIconFileName = APPLICATION_ICONFILE

'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If CurrentAppHandlerRef Is Nothing Then
      Set CurrentAppHandlerRef = CurrentApplication
   End If

   With CurrentAppHandlerRef
   
      'Zur Sicherheit AccDb einstellen
      Set .AppDb = CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird
   
      'Anwendungsname
      .ApplicationName = APPLICATION_NAME
      .ApplicationFullName = APPLICATION_FULLNAME
      .ApplicationTitle = APPLICATION_TITLE
      
      'Version
      .Version = APPLICATION_VERSION
      
      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = APPLICATION_STARTFORMNAME
   
   End With
   
'----------------------------------------------------------------------------
' Erweiterung: AppFile
'
#If USE_CLASS_ApplicationHandler_AppFile = 1 Then
   modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile
#End If

'Dateiauswahl in Textbox
#If USE_CLASS_ApplicationHandler_DirTextbox = 1 Then
   modApplication.AddApplicationHandlerExtension New ApplicationHandler_DirTextbox
#End If

'----------------------------------------------------------------------------
' Erweiterungen für Add-In laden
'
   'Konfiguration/Add-In-Einstellungen
   modApplication.AddApplicationHandlerExtension New ACLibConfiguration
   
   'Import/Export von Dateien bzw. Access-Objekten
   modApplication.AddApplicationHandlerExtension New ACLibFileManager
   
   

'----------------------------------------------------------------------------
' Konfiguration nach Erweiterungen
'

   'AppIcon
   'oCurrentAppHandler.SetAppIcon CodeProject.Path & "\" & m_ApplicationIconFile, True

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
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
Private Sub SetAppFiles()

   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)

End Sub
