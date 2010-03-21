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
' Beispiel für Anwendungskonfiguration mit Lizenzprüfung
' </summary>
' <remarks>
' Beispieldaten:
'   Name: Access Code Library
'   Key:  118F6-F9A9E-0EE83-5A792
' </remarks>
'\ingroup base
'**/
'<codelib>
'  <file>_codelib/samples/_config_Application_LicenseVerifier.bas</file>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
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

'/** \addtogroup base
'@{ **/

Private Const m_ApplicationName As String = "Access Code Library: Lizenzprüfung"
Private Const m_ApplicationFullName As String = m_ApplicationName
Private Const m_ApplicationTitle As String = m_ApplicationName
Private Const m_ApplicationIconFile As String = m_ApplicationName & ".ico"

Private Const m_DefaultErrorHandlerMode = ACLibErrorHandlerMode.aclibErrMsgBox

Private Const m_EXTENSIONKEY_LicenseVerifier As String = "LicenseVerifier"
Private Const m_ApplicationKey As String = "966B11C0A07E9E9AC5B22C8D579CAC17"
Private Const m_LicenseKey_KeyLen As Long = 20
Private Const m_LicenseKey_Prefix As String = "C8E37DE8E"
Private Const m_LicenseKey_Suffix As String = "C51A"
Private Const m_LicenseKey_Loops As Long = 2


'
' Farben
'
Public Enum ApplicationColors
   MdiBackColor = 8421504         ' = RGB(128,128,128)
   MdiBackColorAppStart = 5263440 ' = RGB(80,80,80)
End Enum

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
' Erweiterung LicenseVerifier:
'
   AddApplicationHandlerExtension New ApplicationHandler_LicenseVerifier
   oCurrentAppHandler.Extensions(m_EXTENSIONKEY_LicenseVerifier).Init _
                           m_ApplicationKey, m_LicenseKey_KeyLen, _
                           m_LicenseKey_Prefix, m_LicenseKey_Suffix, _
                           m_LicenseKey_Loops
                           
'----------------------------------------------------------------------------
' Erweiterung AppLogin:
'
   AddApplicationHandlerExtension New ApplicationHandler_AppLogin_Win

'----------------------------------------------------------------------------
' Konfiguration nach Initialisierung der Erweiterungen
'


End Sub

'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------


'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
