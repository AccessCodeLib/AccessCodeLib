Attribute VB_Name = "_config_Application"
'---------------------------------------------------------------------------------------
' Modul: _config_Application (Beispiel)
'---------------------------------------------------------------------------------------
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_dotnetlib/samples/_config_Application.bas</file> <-- für Exportsperre entfernen bzw. umschreiben!!!
'  <replace>base/_config_Application.bas</replace> 'dieses Modul mit <file> ersetzen ... es darf nur ein Konfig-Datei geben
'  <license>_codelib/license.bas</license>
'  <use>base/modApplication.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>_dotnetlib/integration/DotNetLibsSetup.bas</use>
'  <use>_dotnetlib/integration/DotNetLibRepair.frm</use>
'  <use>/data/SqlTools_DotNetLib.bas</use>
'  <use>/data/SqlTools_DotNetLib_Examples.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_ApplicationName As String = "ACL DotNetLib Integration Example"
Private Const m_ApplicationFullName As String = m_ApplicationName
Private Const m_ApplicationTitle As String = m_ApplicationName
Private Const m_ApplicationIconFile As String = m_ApplicationName & ".ico"

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
      .ApplicationStartFormName = "DotNetLibRepair"

   End With
   
   
'----------------------------------------------------------------------------
' Erweiterung: ...
'


'----------------------------------------------------------------------------
' Konfiguration nach Initialisierung der Erweiterungen
'
   'Icon der Anwendung und Fenster - erst nach AppFile-Initialisierung laden,
   '                                 falls Icon in AppFile-Tabelle enthalten ist.
   oCurrentAppHandler.SetAppIcon CurrentProject.Path & "\" & m_ApplicationIconFile, True
   
   
End Sub

'############################################################################
'
' Funktionen für die Anwendungswartung
' (werden nur im Anwendungsentwurf benötigt)
'
'----------------------------------------------------------------------------
' Hilfsfunktion zum Speichern von Dateien in die lokale AppFile-Tabelle
'----------------------------------------------------------------------------
Private Sub setAppFiles()
   Call CurrentApplication.SaveAppFile("AppIcon", CurrentProject.Path & "\TestApp.ico")
End Sub
