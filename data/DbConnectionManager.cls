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
' Klassenmodul: DbConnectionManager
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Klasse zur Instanzierung, Terminierung und Basissteuerung von Hilfsklassen
' </summary>
' <remarks>In dieser Klasse wird u. a. auf die Ereignisse der <see cref="DbConnectionInfo">DbConnectionInfo</see>- u. <see cref="DbConnectionHandler">DbConnectionHandler</see>-Instanz reagiert</remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionManager.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionInfo.cls</use>
'  <use>data/DbConnectionHandler.cls</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Verbindungsparameter
Private WithEvents m_DbConnectionInfo As DbConnectionInfo
Attribute m_DbConnectionInfo.VB_VarHelpID = -1

'DB-Verbindung
Private WithEvents m_DbConnection As DbConnectionHandler
Attribute m_DbConnection.VB_VarHelpID = -1

Private m_PermanentBackendRst As DAO.Recordset
Private m_Disposed As Boolean

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Public Sub Dispose()

On Error Resume Next
   
   If m_Disposed Then Exit Sub

   DisposePermanentBackendRst
   
   Set m_DbConnection = Nothing
   Set m_DbConnectionInfo = Nothing

   m_Disposed = (0 = Err.Number)

End Sub

Private Sub DisposePermanentBackendRst()
On Error Resume Next
   If Not (m_PermanentBackendRst Is Nothing) Then
      m_PermanentBackendRst.Close
      Set m_PermanentBackendRst = Nothing
   End If
End Sub

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

On Error Resume Next

   CheckConnectionStatus = ConnectionInfo.CheckConnection

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

   If m_DbConnectionInfo Is Nothing Then
      Set m_DbConnectionInfo = New DbConnectionInfo
      m_Disposed = False
   End If
   Set ConnectionInfo = m_DbConnectionInfo
   
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

   If m_DbConnection Is Nothing Then
      Set m_DbConnection = New DbConnectionHandler
      InitDbConnection
   End If
   Set DbConnection = m_DbConnection
   
End Property

Private Sub InitDbConnection()
   'DbConnection neu einstellen
   Set Me.DbConnection.CurrentDb = Application.CurrentDb ' ... nicht unbedingt erforderlich, wird bei Bedarf über Event angefordert
   With Me.ConnectionInfo
      Me.DbConnection.OdbcConnectionString = .OdbcConnectionString
      Me.DbConnection.OledbConnectionString = .OledbConnectionString
      SqlTools.SqlDateFormat = .DateFormat
   End With
   OpenPermanentBackendRecordset
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


' ########################################################################################
'
' Ereignisbehandlung
'
' DbConnectioninfo
'
Private Sub m_DbConnectionInfo_DbmsConnectionChanged(ByVal newCID As String)
   'DbConnection neu initialisieren lassen
   
On Error Resume Next

   'ADO-Verbindung schließen
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schließen
   Me.DbConnection.ODBC.Dispose
   
   'Verbindung zum BE öffnen
   OpenPermanentBackendRecordset
   
End Sub

Private Sub m_DbConnectionInfo_DbmsConnectionInfoCleared()

'=> Verbindungsinstanzen entfernen

On Error Resume Next

   'ADO-Verbindung schließen
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schließen
   Me.DbConnection.ODBC.Dispose
   
End Sub

Private Sub m_DbConnectionInfo_Disposed()
On Error Resume Next
   Set m_DbConnectionInfo = Nothing
End Sub

Private Sub m_DbConnectionInfo_PasswordChanged()

   MsgBox "Für die ODBC-Verbindung der verknüpften Tabellen ist ein Neustart erforderlich!" & vbNewLine & _
          "Die Verbindung über ADODB/OLEDB reagiert ab sofort auf das neue Passwort."

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
' kein Neusetzen erforderlich, Connection wird über OledbConnectionString eingestellt
' sollte dieser auch fehlen, wird ErrorEmptyOledbConnectionString ausgelöst.
End Sub

Private Sub m_DbConnection_ErrorMissingCurrentDb(ByVal Msg As String, NewCurrentDbRef As DAO.Database)
   Set NewCurrentDbRef = Application.CurrentDb
End Sub
