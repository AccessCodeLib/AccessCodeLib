Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/_config_Application.bas</file>
'  <replace>base/_config_Application.bas</replace> 'dieses Modul ersetzt base/_config_Application.bas
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/FilterFormWizard/defGlobal_ACLibFilterFormWizard.bas</use>
'  <use>base/_initApplication.bas</use>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>_codelib/addins/shared/AppFileCodeModulTransfer.cls</use>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Versionsnummer
Private Const APPLICATION_VERSION As String = "1.3.5" '2016-08-06
Private Const APPLICATION_CODEMODULES_SVNREV As Long = 365

#Const USE_CLASS_APPLICATIONHANDLER_APPFILE = 1
#Const USE_CLASS_APPLICATIONHANDLER_VERSION = 1

Private Const APPLICATION_NAME As String = "ACLib FilterForm Wizard"
Private Const APPLICATION_FULLNAME As String = "Access Code Library - FilterForm Wizard"
Private Const APPLICATION_TITLE As String = APPLICATION_FULLNAME
Private Const APPLICATION_ICONFILE As String = "ACLib.ico"

Public Const DownloadSource As String = "http://wiki.access-codelib.net/ACLib-FilterForm-Wizard"
Private Const APPLICATION_DOWNLOAD_FOLDER As String = "http://access-codelib.net/download/addins/"
Private Const APPLICATION_DOWNLOAD_VERSIONXMLFILE As String = APPLICATION_DOWNLOAD_FOLDER & "ACLibFilterFormWizard.xml"

Private Const DefaultErrorHandlerMode As Long = ACLibErrorHandlerMode.aclibErrMsgBox

Private Const ApplicationStartFormName As String = "frmFilterFormWizard"

'---------------------------------------------------------------------------------------
' Sub: InitConfig (Josef Pötzl)
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

   modErrorHandler.DefaultErrorHandlerMode = DefaultErrorHandlerMode

   
'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   'defGlobal_FilterFormWizard.ACLibIconFileName = m_ApplicationIconFile


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
      .ApplicationStartFormName = ApplicationStartFormName
   
    
   End With

   
'----------------------------------------------------------------------------
' Erweiterung: AppFile
'
#If USE_CLASS_APPLICATIONHANDLER_APPFILE = 1 Then
   modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile
#End If


'----------------------------------------------------------------------------
' Erweiterung: AppFile
'
#If USE_CLASS_APPLICATIONHANDLER_VERSION = 1 Then
   Dim AppHdlVersion As ApplicationHandler_Version
   
   Set AppHdlVersion = New ApplicationHandler_Version
   modApplication.AddApplicationHandlerExtension AppHdlVersion
   AppHdlVersion.XmlVersionCheckFile = APPLICATION_DOWNLOAD_VERSIONXMLFILE
   
#End If


'----------------------------------------------------------------------------
' Erweiterungen für Add-In laden
'


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
On Error GoTo HandleErr

   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & APPLICATION_ICONFILE)
   SaveCodeModulesToTable

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "setAppFiles", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub

Private Sub SaveCodeModulesToTable()

   Dim CodeModuleArray As Variant
   Dim i As Long
   
   CodeModuleArray = Array("SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager")
   For i = 0 To UBound(CodeModuleArray)
      SaveCodeModulToTable acModule, CodeModuleArray(i), APPLICATION_CODEMODULES_SVNREV
   Next
   
End Sub

Private Sub SaveCodeModulToTable(ByVal ObjType As AcObjectType, ByVal ModulName As String, ByVal SvnRev As Long)
   
   Dim FileName As String

   FileName = FileTools.GetNewTempFileName
   Application.SaveAsText ObjType, ModulName, FileName
   CurrentApplication.SaveAppFile ModulName, FileName, True, "SvnRev", SvnRev
   Kill FileName
   
End Sub
