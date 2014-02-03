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
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Versionsnummer
Private Const m_ApplicationVersion As String = "0.9.3"

#Const USE_CLASS_ApplicationHandler_AppFile = 1

Private Const m_ApplicationName As String = "ACLib FilterForm Wizard"
Private Const m_ApplicationFullName As String = "Access Code Library - FilterForm Wizard"
Private Const m_ApplicationTitle As String = m_ApplicationFullName
Private Const m_ApplicationIconFile As String = "ACLib.ico"

Private Const m_DefaultErrorHandlerMode As Long = ACLibErrorHandlerMode.aclibErrMsgBox

Private Const m_ApplicationStartFormName As String = "frmFilterFormWizard"

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
Public Sub InitConfig(Optional ByRef oCurrentAppHandler As ApplicationHandler = Nothing)

On Error GoTo HandleErr

'----------------------------------------------------------------------------
' Fehlerbehandlung
'

   modErrorHandler.DefaultErrorHandlerMode = m_DefaultErrorHandlerMode

   
'----------------------------------------------------------------------------
' Globale Variablen einstellen
'
   'defGlobal_FilterFormWizard.ACLibIconFileName = m_ApplicationIconFile


'----------------------------------------------------------------------------
' Anwendungsinstanz einstellen
'
   If oCurrentAppHandler Is Nothing Then
      Set oCurrentAppHandler = CurrentApplication
   End If

   With oCurrentAppHandler
   
      'Zur Sicherheit AccDb einstellen
      Set .AppDb = CodeDb 'muss auf CodeDb zeigen,
                          'da diese Anwendung als Add-In verwendet wird
   
      'Anwendungsname
      .ApplicationName = m_ApplicationName
      .ApplicationFullName = m_ApplicationFullName
      .ApplicationTitle = m_ApplicationTitle
      
      'Version
      .Version = m_ApplicationVersion
      
      ' Formular, das am Ende von CurrentApplication.Start aufgerufen wird
      .ApplicationStartFormName = m_ApplicationStartFormName
   
    
   End With

   
'----------------------------------------------------------------------------
' Erweiterung: AppFile
'
#If USE_CLASS_ApplicationHandler_AppFile = 1 Then
   modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile
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

   Call CurrentApplication.Extensions("AppFile").SaveAppFile("AppIcon", CodeProject.Path & "\" & m_ApplicationIconFile)
   SaveModulesInTable

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
