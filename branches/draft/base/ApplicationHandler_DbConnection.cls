VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_DbConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_DbConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse: Einbindung DbConnectionManager
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_DbConnection</code>
' oder mittels ExtensionCollection:\n
' <code>ExtensionsCollectionReferenz.Add New ApplicationHandler_DbConnection</code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_DbConnection.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/DbConnectionManager.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "DbConnection"

Private m_Disposed As Boolean
Private m_DbConnectionManager As DbConnectionManager

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef ObjRef As ApplicationHandler)
   Set m_ApplicationHandler = ObjRef
   m_Disposed = False
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

Private Sub m_ApplicationHandler_AfterDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   Set m_ApplicationHandler = Nothing
End Sub

Private Sub m_ApplicationHandler_BeforeDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef Exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then Exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Public Sub Dispose()

On Error Resume Next

   If m_Disposed Then Exit Sub
   
   DisposeConnectionManager
   m_Disposed = (0 = Err.Number)

End Sub

Public Sub DisposeConnectionManager()
   
   If m_DbConnectionManager Is Nothing Then
      Exit Sub
   End If
   
On Error Resume Next
   m_DbConnectionManager.Dispose
   Set m_DbConnectionManager = Nothing
   
End Sub

'---------------------------------------------------------------------------------------
' Property: ConnectionManager
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zugriff auf DbConnectionManager-Instanz
' </summary>
' <returns>DbConnectionManager</returns>
' <remarks>
' Beim ersten Zugriff wird die Instanz der DbConnectionManager-Klasse erstellt, falls diese noch nicht vorhanden ist.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ConnectionManager() As DbConnectionManager

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set ConnectionManager = m_DbConnectionManager

End Property

'---------------------------------------------------------------------------------------
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Testen, ob Verbindungsaufbau möglich ist
' </summary>
' <returns>Boolean: True = Verbindungsaufbau war erfolgreich</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean
   CheckConnectionStatus = ConnectionManager.CheckConnectionStatus
End Function

'---------------------------------------------------------------------------------------
' Property: ConnectionInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DbConnectionInfo-Instanz der Anwendung.
' </summary>
' <returns>DbConnectionInfo</returns>
' <remarks>
' Dient zum Ausgeben von Informationen zur Datenbankverbindung
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ConnectionInfo() As DbConnectionInfo
   Set ConnectionInfo = ConnectionManager.ConnectionInfo
End Property

'---------------------------------------------------------------------------------------
' Property: DbConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DbConnectionHandler-Instanz = oberste Ebene der Zugriffsstruktur
' </summary>
' <returns>DbConnectionHandler</returns>
' <remarks>
' Kapselt ADODB- und DAO-Methoden sowie ODBC-Zugriffe
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DbConnection() As DbConnectionHandler
   Set DbConnection = ConnectionManager.DbConnection
End Property
