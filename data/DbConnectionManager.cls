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
'  <use>data/DbConnectionSettings.cls</use>
'  <use>data/SqlTools.cls</use>
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

Public Event BeforeInitDbConnection(ByRef OdbcConnectionString As String, ByRef OledbConnectionString As String, ByRef Settings As DbConnectionSettings)
Public Event DbConnectionInitialized()

Private m_DefaultDbConnecionSettings As DbConnectionSettings
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
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Testen, ob Verbindungsaufbau m?glich ist
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
   
   ' DefaultSettings von NewSettings ?bernehmen:
   ' Kopieren und nicht Referenz zuweisen, damit Daten passen, falls an anderer Stelle bereits eine Referenz auf DefaultSettings gesetzt wurde
   DefaultSettings.CopyFrom NewSettings
   
   Set m_DbConnection.CurrentDb = Application.CurrentDb ' ... nicht unbedingt erforderlich, wird bei Bedarf ?ber Event angefordert
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


' ########################################################################################
'
' Ereignisbehandlung
'
' DbConnectioninfo
'
Private Sub m_DbConnectionInfo_DbmsConnectionChanged(ByVal newCID As String)
   'DbConnection neu initialisieren lassen
   
On Error Resume Next

   'ADO-Verbindung schlie?en
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schlie?en
   Me.DbConnection.ODBC.Dispose
   
   'Verbindung zum BE ?ffnen
   OpenPermanentBackendRecordset
   
End Sub

Private Sub m_DbConnectionInfo_DbmsConnectionInfoCleared()

'=> Verbindungsinstanzen entfernen

On Error Resume Next

   'ADO-Verbindung schlie?en
   Me.DbConnection.ADODB.Dispose
   
   'ODBC-Verbindung schlie?en
   Me.DbConnection.ODBC.Dispose
   
End Sub

Private Sub m_DbConnectionInfo_Disposed()
On Error Resume Next
   Set m_DbConnectionInfo = Nothing
End Sub

Private Sub m_DbConnectionInfo_PasswordChanged()

   MsgBox "F?r die ODBC-Verbindung der verkn?pften Tabellen ist ein Neustart erforderlich!" & vbNewLine & _
          "Die Verbindung ?ber ADODB/OLEDB reagiert ab sofort auf das neue Passwort."

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
' kein Neusetzen erforderlich, Connection wird ?ber OledbConnectionString eingestellt
' sollte dieser auch fehlen, wird ErrorEmptyOledbConnectionString ausgel?st.
End Sub

Private Sub m_DbConnection_ErrorMissingCurrentDb(ByVal Msg As String, NewCurrentDbRef As DAO.Database)
   Set NewCurrentDbRef = Application.CurrentDb
End Sub
