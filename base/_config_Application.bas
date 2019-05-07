Attribute VB_Name = "_config_Application"
'
'############################################################################
'##                                                                        ##
'##  Individuell gestaltete Config-Module nicht in das Repositiory laden!  ##
'##                                                                        ##
'############################################################################
'
'---------------------------------------------------------------------------------------
' Modul: _config_Application (Beispiel)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beispiel für Anwendungskonfiguration
' </summary>
' <remarks>
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
' </remarks>
'\ingroup base
'**/
'<codelib>
'  <file>base/_config_Application.bas</file> <-- umschreiben bzw. löschen!!!
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Das Modul _config_Application wird vom Import-Assistenden nicht überschrieben.
' Sollte eine neues _config_Application-Modul geladen werden,
' ist das alte zuvor umzubennen oder zu löschen.
'
'
Option Compare Text
Option Explicit

'/** \addtogroup base
'@{ **/

Private Const m_ApplicationVersion As String = "0.0.0"

Private Const m_ApplicationName As String = "Anwendungsnamen eingeben"
Private Const m_ApplicationFullName As String = m_ApplicationName
Private Const m_ApplicationTitle As String = m_ApplicationName
Private Const m_ApplicationIconFile As String = m_ApplicationName & ".ico"

#If USE_GLOBAL_ERRORHANDLER Then
Const m_DefaultErrorHandlerMode = ACLibErrorHandlerMode.aclibErrMsgBox
#End If

Private Const m_ApplicationStartFormName As String = ""

#USE_EXTENSIONS = True
#If USE_EXTENSIONS = True
Private m_Extensions As ApplicationHandler_ExtensionCollection
#End If

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
Public Sub InitConfig(Optional oCurrentAppHandler As ApplicationHandler = Nothing)

'----------------------------------------------------------------------------
' Fehlerbehandlung
'
#If USE_GLOBAL_ERRORHANDLER Then
   modErrorHandler.DefaultErrorHandlerMode = m_DefaultErrorHandlerMode
#End If

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

      'Version
      .Version = m_ApplicationVersion
      
      ' Formular, das am Ende von StartApplication aufgerufen wird
      '.ApplicationStartFormName = m_ApplicationStartFormName

   End With
   
   
'----------------------------------------------------------------------------
' Erweiterung: ...
'
#If USE_EXTENSIONS = True

   Set m_Extensions = New ApplicationHandler_ExtensionCollection
   With m_Extensions
      Set .ApplicationHandler = oCurrentAppHandler
	  
	  ' Erweiterungen laden
	  ' z. B.:
      '.Add New ApplicationHandler_AppFile
	  
   End With

#End If
   
'----------------------------------------------------------------------------
' Konfiguration nach Initialisierung der Erweiterungen
'
   'Icon der Anwendung und Fenster - erst nach AppFile-Initialisierung laden,
   '                                 falls Icon in AppFile-Tabelle enthalten ist.
   'oCurrentAppHandler.SetAppIcon CurrentProject.Path & "\" & m_ApplicationIconFile, True
   
   
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
   Call CurrentApplication.SaveAppFile("AppIcon", CurrentProject.Path & "\TestApp.ico")
End Sub


'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
