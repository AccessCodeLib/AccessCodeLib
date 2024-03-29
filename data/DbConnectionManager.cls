VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: data.DbConnectionManager
'---------------------------------------------------------------------------------------
'
' Class for instantiation, termination and basic control of data access classes
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     This class reacts to the events of the <data.DbConnectionInfo> and <data.DbConnectionHandler> instance.
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionManager.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionInfo.cls</use>
'  <use>data/DbConnectionHandler.cls</use>
'  <use>data/DbConnectionSettings.cls</use>
'  <use>data/SqlTools.cls</use>
'  <ref><name>DAO</name><major>12</major><minor>0</minor><guid>{4AC9E1DA-5BAD-4AC7-86E3-24F4CDCECA28}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Connection parameters
Private WithEvents m_DbConnectionInfo As DbConnectionInfo
Attribute m_DbConnectionInfo.VB_VarHelpID = -1

'DB connection
Private WithEvents m_DbConnection As DbConnectionHandler
Attribute m_DbConnection.VB_VarHelpID = -1

Public Event BeforeInitDbConnection(ByRef OdbcConnectionString As String, ByRef OledbConnectionString As String, ByRef Settings As DbConnectionSettings)
Public Event DbConnectionInitialized()

Private m_DefaultDbConnecionSettings As DbConnectionSettings
Private m_PermanentBackendRst As DAO.Recordset
Private m_Disposed As Boolean

Private m_DbmsUserName As String
Private m_DbmsUserPassword As String

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Sub: Dispose
'---------------------------------------------------------------------------------------
'
' Remove object references
'
'---------------------------------------------------------------------------------------
Public Sub Dispose()

On Error Resume Next
   
   If m_Disposed Then Exit Sub

   DisposePermanentBackendRst
   
   Set m_DbConnection = Nothing
   Set m_DbConnectionInfo = Nothing

   m_Disposed = (0 = Err.Number)

End Sub

Public Sub SetUserData(ByVal UserName As String, ByVal Password As String)
   m_DbmsUserName = UserName
   m_DbmsUserPassword = Password
End Sub

'---------------------------------------------------------------------------------------
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'
' Test whether a connection can be established
'
' Returns:
'     Boolean - True if success
'
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean

On Error Resume Next

   CheckConnectionStatus = ConnectionInfo.CheckConnection

End Function

'---------------------------------------------------------------------------------------
' Property: DefaultSettings
'---------------------------------------------------------------------------------------
'
' Test whether a connection can be established
'
' Returns:
'     <data.DbConnectionSettings>
'
'---------------------------------------------------------------------------------------
Public Property Get DefaultSettings() As DbConnectionSettings
   If m_DefaultDbConnecionSettings Is Nothing Then
      Set m_DefaultDbConnecionSettings = New DbConnectionSettings
      m_DefaultDbConnecionSettings.DefaultAdodbCursorLocation = adUseServer
      m_DefaultDbConnecionSettings.OpenPermanentBackendRecordset = True
   End If
   Set DefaultSettings = m_DefaultDbConnecionSettings
End Property

Private Sub DisposePermanentBackendRst()
On Error Resume Next
   If Not (m_PermanentBackendRst Is Nothing) Then
      m_PermanentBackendRst.Close
      Set m_PermanentBackendRst = Nothing
   End If
End Sub

'---------------------------------------------------------------------------------------
' Property: ConnectionInfo
'---------------------------------------------------------------------------------------
'
' DbConnectionInfo instance
'
' Returns:
'     <data.DbConnectionInfo>
'
' Remarks:
'     Used to output information on the database connection
'
'---------------------------------------------------------------------------------------
Public Property Get ConnectionInfo() As DbConnectionInfo

   If m_DbConnectionInfo Is Nothing Then
      Set m_DbConnectionInfo = New DbConnectionInfo
      m_Disposed = False
   End If
   Set ConnectionInfo = m_DbConnectionInfo
   
End Property

'---------------------------------------------------------------------------------------
' Property: DbConnection
'---------------------------------------------------------------------------------------
'
' DbConnectionHandler instance = Top level of the data access structure
'
' Returns:
'     <data.DbConnectionHandler>
'
' Remarks:
'     Encapsulates ADODB and DAO methods as well as ODBC accesses
'
'---------------------------------------------------------------------------------------
Public Property Get DbConnection() As DbConnectionHandler

   If m_DbConnection Is Nothing Then
      InitDbConnection
   End If
   Set DbConnection = m_DbConnection
   
End Property

Private Sub InitDbConnection()
   
   Dim OdbcConnectionString As String
   Dim OledbConnectionString As String
   Dim NewSettings As DbConnectionSettings
   
   If m_DbConnection Is Nothing Then
      Set m_DbConnection = New DbConnectionHandler
   End If
   
   Set NewSettings = DefaultSettings.Clone
   With Me.ConnectionInfo
      OdbcConnectionString = .OdbcConnectionString
      OledbConnectionString = .OledbConnectionString
      NewSettings.SqlDateFormat = .DateFormat
      NewSettings.SqlWildCardString = .WildCardString
      NewSettings.SqlBooleanTrueString = .BooleanTrueString
   End With
   
   RaiseEvent BeforeInitDbConnection(OdbcConnectionString, OledbConnectionString, NewSettings)
   
   ' DefaultSettings von NewSettings �bernehmen:
   ' Kopieren und nicht Referenz zuweisen, damit Daten passen, falls an anderer Stelle bereits eine Referenz auf DefaultSettings gesetzt wurde
   DefaultSettings.CopyFrom NewSettings
   
   Set m_DbConnection.CurrentDb = Application.CurrentDb ' ... nicht unbedingt erforderlich, wird bei Bedarf �ber Event angefordert
   m_DbConnection.OdbcConnectionString = OdbcConnectionString
   m_DbConnection.OledbConnectionString = OledbConnectionString
   
   With DefaultSettings

      SqlTools.SqlDateFormat = .SqlDateFormat
      SqlTools.SqlWildCardString = .SqlWildCardString
      SqlTools.SqlBooleanTrueString = .SqlBooleanTrueString
      
       m_DbConnection.ADODB.DefaultCursorLocation = .DefaultAdodbCursorLocation
      
      If .OpenPermanentBackendRecordset Then
         OpenPermanentBackendRecordset
      End If
      
   End With

   RaiseEvent DbConnectionInitialized

End Sub

Private Sub OpenPermanentBackendRecordset()

   If Not (m_PermanentBackendRst Is Nothing) Then
   On Error Resume Next
      m_PermanentBackendRst.Close
   On Error GoTo 0
      Set m_PermanentBackendRst = Nothing
   End If
   
   If Len(ConnectionInfo.DbmsName) = 0 Then
      Exit Sub
   End If
   
   If Len(Me.ConnectionInfo.OdbcConnectionString) > 0 Then
      Set m_PermanentBackendRst = Me.DbConnection.ODBC.OpenRecordsetPT(Me.ConnectionInfo.PermanentBackendRstSqlText, dbOpenSnapshot)
   End If
   
End Sub


'########################################################################################
'
' Event handling
'
' DbConnectioninfo
'
Private Sub m_DbConnectionInfo_DbmsConnectionChanged(ByVal newCID As String)
   'DbConnection neu initialisieren lassen
   
On Error Resume Next

   'ADO-Verbindung schlie�en
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schlie�en
   Me.DbConnection.ODBC.Dispose
   
   'Verbindung zum BE �ffnen
   If DefaultSettings.OpenPermanentBackendRecordset Then
      OpenPermanentBackendRecordset
   End If
   
End Sub

Private Sub m_DbConnectionInfo_DbmsConnectionInfoCleared()

'=> Verbindungsinstanzen entfernen

On Error Resume Next

   'ADO-Verbindung schlie�en
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schlie�en
   Me.DbConnection.ODBC.Dispose
   
End Sub

Private Sub m_DbConnectionInfo_Disposed()
On Error Resume Next
   Set m_DbConnectionInfo = Nothing
End Sub

Private Sub m_DbConnectionInfo_LoginDataRequired(UserName As String, Password As String)
   If Len(m_DbmsUserName) > 0 Then
      UserName = m_DbmsUserName
      Password = m_DbmsUserPassword
   End If
End Sub

Private Sub m_DbConnectionInfo_PasswordChanged()

   MsgBox "A restart is required for the ODBC connection of the linked tables!" & vbNewLine & _
          "The connection via ADODB/OLEDB responds to the new password from now on."

   'DbConnection neu einstellen
   InitDbConnection

End Sub

'
' DbConnection
'
Private Sub m_DbConnection_Disposed()
On Error Resume Next
   Set m_DbConnection = Nothing
End Sub

Private Sub m_DbConnection_ErrorEmptyOdbcConnectionString(ByVal Msg As String, NewConnectionString As String)
   NewConnectionString = Me.ConnectionInfo.OdbcConnectionString
End Sub

Private Sub m_DbConnection_ErrorEmptyOledbConnectionString(ByVal Msg As String, NewConnectionString As String)
   NewConnectionString = Me.ConnectionInfo.OledbConnectionString
End Sub

Private Sub m_DbConnection_ErrorMissingCurrentConnection(ByVal Msg As String, NewCurrentConnection As ADODB.Connection)
' No reset required, connection is set via OledbConnectionString.
' Should this also be missing, ErrorEmptyOledbConnectionString is triggered.
End Sub

Private Sub m_DbConnection_ErrorMissingCurrentDb(ByVal Msg As String, NewCurrentDbRef As DAO.Database)
   Set NewCurrentDbRef = Application.CurrentDb
End Sub
