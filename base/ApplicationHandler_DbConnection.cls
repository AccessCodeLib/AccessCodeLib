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
' Class: base.ApplicationHandler_DbConnection
'---------------------------------------------------------------------------------------
'
' Extension for ApplicationHandler class: integration of DbConnectionManager
'
' Remarks:
'     Activation in the _config_Application module:
'
'     (start code)
'     modApplication.AddApplicationHandlerExtension New ApplicationHandler_DbConnection
'     (end)
'
'     or via ExtensionCollection:
'
'     (start code)
'     ExtensionsCollectionReferenz.Add New ApplicationHandler_DbConnection
'     (end)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_DbConnection.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/DbConnectionManager.cls</use>
'  <description>ApplicationHandler extension: provides the DbConnectionManager (backend connection, connection info and status)</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "DbConnection"

Private Const SETTINGS_EXTENSION_KEY As String = "Settings"

Private m_Disposed As Boolean
Private WithEvents m_DbConnectionManager As DbConnectionManager
Attribute m_DbConnectionManager.VB_VarHelpID = -1

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
'
' Access to the DbConnectionManager instance
'
' Returns:
'     DbConnectionManager
'
' Remarks:
'     The instance of the DbConnectionManager class is created on first access if it does not exist yet.
'
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
'
' Test whether a connection can be established
'
' Returns:
'     Boolean: True = connection was established successfully
'
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean
   CheckConnectionStatus = ConnectionManager.CheckConnectionStatus
End Function

'---------------------------------------------------------------------------------------
' Property: ConnectionInfo
'---------------------------------------------------------------------------------------
'
' DbConnectionInfo instance of the application
'
' Returns:
'     DbConnectionInfo
'
' Remarks:
'     Used to output information about the database connection
'
'---------------------------------------------------------------------------------------
Public Property Get ConnectionInfo() As DbConnectionInfo
   Set ConnectionInfo = ConnectionManager.ConnectionInfo
End Property

'---------------------------------------------------------------------------------------
' Property: DbConnection
'---------------------------------------------------------------------------------------
'
' DbConnectionHandler instance = top level of the data access structure
'
' Returns:
'     DbConnectionHandler
'
' Remarks:
'     Encapsulates ADODB and DAO methods as well as ODBC access
'
'---------------------------------------------------------------------------------------
Public Property Get DbConnection() As DbConnectionHandler
   Set DbConnection = ConnectionManager.DbConnection
End Property

Private Sub m_DbConnectionManager_BeforeInitDbConnection(OdbcConnectionString As String, OledbConnectionString As String, Settings As DbConnectionSettings)

   With Settings
   
      .SqlBooleanTrueString = m_ApplicationHandler.GetExtensionProperty(SETTINGS_EXTENSION_KEY, "SqlBooleanTrueString", .SqlBooleanTrueString)
      .SqlDateFormat = m_ApplicationHandler.GetExtensionProperty(SETTINGS_EXTENSION_KEY, "SqlDateFormat", .SqlDateFormat)
      .SqlWildCardString = m_ApplicationHandler.GetExtensionProperty(SETTINGS_EXTENSION_KEY, "SqlWildCardString", .SqlWildCardString)
      
      .DefaultAdodbCursorLocation = m_ApplicationHandler.GetExtensionProperty(SETTINGS_EXTENSION_KEY, "DefaultAdodbCursorLocation", .DefaultAdodbCursorLocation)
      
      .OpenPermanentBackendRecordset = m_ApplicationHandler.GetExtensionProperty(SETTINGS_EXTENSION_KEY, "OpenPermanentBackendRecordset", .OpenPermanentBackendRecordset)
      
   End With

End Sub
