VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdodbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: AdodbHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' ADODB-Zugriffsmethoden
' </summary>
' <remarks></remarks>
'\ingroup data_adodb
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/ado/AdodbHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'  <test>_test/data/ado/AdodbHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'ADODB-Connection
Private m_AdoConnection As ADODB.Connection

'ADO-Connectionstring
Private m_conADO_ConnectionString As String

'Error-Events für mögliche Fehlerbehandlung über die FE-Anwendung
'/**
' <summary>
' Ereignis für fehlende ADODB-Connection-Referenz
' </summary>
' <param name="Msg">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="NewCurrentConnection">ADODB.Connection-Referenz, die für CurrentConnection verwendet werden soll</param>
' <remarks>
' Über den ByRef-Parameter NewCurrentConnection kann die passende ADODB.Database-Referenz weitergegeben werden.
' Eine Referenz-Weitergabe ist allerdings nicht zwingend erforderlich, da z. B. in der Methode #CurrentConnection
' die Connection dann aus dem #ConnectionString erstellt wird.
' </remarks>
'**/
Public Event ErrorMissingCurrentConnection(ByVal Msg As String, ByRef NewCurrentConnection As ADODB.Connection)

'/**
' <summary>
' Ereignis für fehlenden OLEDB-Connectionstring
' </summary>
' <param name="Msg">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="NewConnectionString">OLEDB-Connectionstring</param>
' <remarks>über den ByRef-Parameter NewConnectionString kann der passende OLEDB-Connectionstring weitergegeben werden.</remarks>
'**/
Public Event ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)


Public Event HookInsertIdentityReturn(ByRef InsertSQL As String, ByRef IdentityTable As String, _
                                      ByRef ActiveConnection As ADODB.Connection, ByRef IdentitySelectString As String, _
                                      ByRef HookResumeMode As Long, ByRef HookResumeMessage As Variant)

Private m_HooksEnabled As Boolean

'Private Type AdodbCommandParamDef
'   ParamName As String
'   ParamValue As Variant
'End Type

'
' Init / Terminate
' ----------------
'
'Private Sub Class_Initialize()
''
'End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Public Sub Dispose()
   Set m_AdoConnection = Nothing
End Sub

Public Property Get HooksEnabled() As Boolean
   HooksEnabled = m_HooksEnabled
End Property

Public Property Let HooksEnabled(ByVal bEnable As Boolean)
   m_HooksEnabled = bEnable
End Property


'---------------------------------------------------------------------------------------
' Property: CurrentConnection (2009-08-06)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Aktuell eingestellte ADODB-Connction
' </summary>
' <returns>ADODB.Connection</returns>
' <remarks>
' Falls noch keine ADODB-Connection vorhanden ist, wird eine neue ADODB-Verbindung erstellt.
' Ablauf der Instanzieren:
' 1. RaiseEvent #ErrorMissingCurrentConnection ... falls keine Connection übergeben wurde:
' 2. Neue ADODB-Connection auf Basis des OLEDB-Connectionstring erzeugen (Einsatz der Eigenschaft #ConnectionString)
'
' Falls bereits eine Referenz zu einer ADODB-Connection vorhanden ist, wird geprüft, ob diese geschlossen ist und im Bedarfsfall geöffnet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnection() As ADODB.Connection

   Dim strConnString As String

   If (m_AdoConnection Is Nothing) Then
      'Connection über Event anfordern.
      RaiseEvent ErrorMissingCurrentConnection("Unavailable CurrentConnection", m_AdoConnection)
      If (m_AdoConnection Is Nothing) Then
         'Neue Verbindung einstellen, falls durch Event keine Connection übergeben wurde
         strConnString = ConnectionString '... siehe Property ConnectionString
         If Len(strConnString) > 0 Then
            Set m_AdoConnection = New ADODB.Connection
            m_AdoConnection.Open strConnString
         End If
      End If
   ElseIf (m_AdoConnection.State = ObjectStateEnum.adStateClosed) Then
      If Len(m_AdoConnection.ConnectionString) > 0 Then
         m_AdoConnection.Open
      Else 'neu verbinden
         m_AdoConnection.Open ConnectionString
      End If
   End If
   Set CurrentConnection = m_AdoConnection

End Property

'---------------------------------------------------------------------------------------
' Property: CurrentConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Übergabe einer ADODB-Connection
' </summary>
' <param name="NewConnection">Neue ADODB-Connection</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentConnection(ByRef NewConnection As ADODB.Connection)
   If Not (m_AdoConnection Is Nothing) Then
      If m_AdoConnection.State <> adStateClosed Then
         m_AdoConnection.Close
      End If
      Set m_AdoConnection = Nothing
   End If
   Set m_AdoConnection = NewConnection
End Property

'---------------------------------------------------------------------------------------
' Property: ConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OLEDB-Connectionstring
' </summary>
' <returns>OLEDB-Connectionstring</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ConnectionString() As String
   If Len(m_conADO_ConnectionString) = 0 Then
      If Not (m_AdoConnection Is Nothing) Then
         m_conADO_ConnectionString = m_AdoConnection.ConnectionString
      End If
      If Len(m_conADO_ConnectionString) = 0 Then
         'Connectiostring mittels Event abfragen
         RaiseEvent ErrorEmptyConnectionString("DbConnection.ConnectionString", m_conADO_ConnectionString)
      End If
   End If
   ConnectionString = m_conADO_ConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: ConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OLEDB-Connectionstring
' </summary>
' <param name="Value">OLEDB-Connectionstring für den Verbindungaufbau</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ConnectionString(ByVal Value As String)
   Set CurrentConnection = Nothing
   m_conADO_ConnectionString = Value
End Property

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sql-Anweisung ausführen
' </summary>
' <param name="CommandText">SQL-Anweisung</param>
' <param name="RecordsAffected">ByRef-Parameter zur Rückgabe der Anzahl der von der Execute-Anweisung betroffenen Datensätze</param>
' <param name="Options">zusätztliche Optionen der ADODB.Connection-Execute-Methode</param>
' <returns>ADODB.Recordset</returns>
' <remarks>Kapselung der ADODB.Connection.Execute-Methode (Parameter werden 1:1 weitergereicht)</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal CommandText As String, _
               Optional ByRef RecordsAffected As Long, _
               Optional ByVal Options As ADODB.ExecuteOptionEnum = -1) As ADODB.Recordset
   Set Execute = Me.CurrentConnection.Execute(CommandText, RecordsAffected, Options)
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteCommand
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sql-Anweisung mittels ADODB.Command ausführen
' </summary>
' <param name="CmdText">ADODB.CommandText-Wert</param>
' <param name="CmdType">ADODB.CommandType-Wert</param>
' <param name="commandParamDefs">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <param name="RecordsAffected">ByRef-Parameter zur Rückgabe der Anzahl der von der Execute-Anweisung betroffenen Datensätze</param>
' <param name="Options">zusätztliche Optionen der ADODB.Connection-Execute-Methode</param>
' <returns>ADODB.Recordset</returns>
' <remarks>Mit dieser Variante kann sehr einfach SQL-Injection vermieden werden, da die Parameter nicht direkt in der SQL-Anweisung eingebaut sind und somit ein Parameterwert (der z. B. von einer Benutzereingabe stammt) keine SQL-Anweisung einschleusen kann.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteCommand(ByVal CmdText As String, _
                               ByVal CmdType As ADODB.CommandTypeEnum, _
                      Optional ByVal CmdParamDefs As Variant, _
                      Optional ByRef RecordsAffected As Long, _
                      Optional ByVal Options As Long = -1) As ADODB.Recordset
'gibt geöffnetes ADODB-Recordset zurück

   Dim cmd As ADODB.Command
   Dim ParamDefsArraySize As Long
   Dim i As Long
   Dim strParamName As String

   Set cmd = New ADODB.Command
   With cmd
   
      Set .ActiveConnection = Me.CurrentConnection
      .CommandText = CmdText
      .CommandType = CmdType
   
      If Not IsMissing(CmdParamDefs) Then
         ParamDefsArraySize = UBound(CmdParamDefs, 1)
         For i = 0 To ParamDefsArraySize
            strParamName = Nz(CmdParamDefs(i, 0), vbNullString)
            If Len(strParamName) > 0 Then
               .Parameters(strParamName).Value = CmdParamDefs(i, 1)
            End If
         Next i
      End If
      
      Set ExecuteCommand = .Execute(RecordsAffected, , Options)
      
   End With
   Set cmd = Nothing

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordset (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OpenRecordset
' </summary>
' <param name="Source">Datenquelle (z.B. SQL-Anweisung)</param>
' <param name="CursorType">ADODB.CursorTypeEnum-Wert</param>
' <param name="LockType">ADODB.LockTypeEnum-Wert</param>
' <param name="CursorLocation">ADODB.CursorLocationEnum-Wert</param>
' <param name="DisconnectedRecordset">True: Verbindung nach Öffnen des Recordset kappen ... verbindungsloses Recordset</param>
' <returns>ADODB.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal CursorType As ADODB.CursorTypeEnum = ADODB.CursorTypeEnum.adOpenForwardOnly, _
                     Optional ByVal LockType As ADODB.LockTypeEnum = ADODB.LockTypeEnum.adLockReadOnly, _
                     Optional ByVal CursorLocation As ADODB.CursorLocationEnum = ADODB.CursorLocationEnum.adUseServer, _
                     Optional ByVal DisconnectedRecordset As Boolean = False) As ADODB.Recordset
' gibt geöffnetes ADO-Recordset zurück
' Source = SQL-Anweisung

   Dim rst As ADODB.Recordset

   Set rst = New ADODB.Recordset
   With rst
      .CursorLocation = CursorLocation
      .Open Source, Me.CurrentConnection, CursorType, LockType
      
      If DisconnectedRecordset Then
         Set .ActiveConnection = Nothing
      End If
      
   End With
   Set OpenRecordset = rst
   
End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für Übergabe einer vollständigen SQL-Anweisung
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgeliefert werden soll. Die Index-Übergabe erfolgt analaog ADODB.Recordset.Fields(index)</param>
' <param name="ValueIfNull">Rückgabewert bei fehlendem Datensatz oder NULL-Feldwert (Standard: Null</param>
' <returns>Variant (NULL, falls SQL-Anweisung keinen Datensatz liefert.)</returns>
' <remarks>Die SQL-Anweisung ist im SQL-Dialekt des betroffenen DBMS zu schreiben.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQL(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim rst As ADODB.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(SqlText, adOpenForwardOnly, adLockReadOnly, adUseServer)
   With rst
      If .EOF Then
         LookupSQL = ValueIfNull
      Else
         LookupSQL = Nz(.Fields(Index), ValueIfNull)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
   Exit Function
   
HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "LookupSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: Lookup
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Lookup(ByVal Expr As String, ByVal Domain As String, _
              Optional ByVal Criteria As Variant, _
              Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim strSQL As String

   strSQL = "SELECT " & Expr & " FROM " & Domain
   If Not (VarType(Criteria) = vbError) Then
      If Len(Criteria) > 0 Then
         strSQL = strSQL & " WHERE " & Criteria
      End If
   End If
   Lookup = Me.LookupSQL(strSQL, , ValueIfNull)

End Function

'---------------------------------------------------------------------------------------
' Function: DCount
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DCount-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Long</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Long
   Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)
End Function

'---------------------------------------------------------------------------------------
' Function: Max
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DMax-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Max(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Min
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DMin-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Sum
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DSum-Recordset-Ersatzfunktion
' </summary>
' <param name="Expr">Feldname oder berechneter Ausdruck (Select-Teil)</param>
' <param name="Domain">Tabelle, Abfrage oder SQL-Ausdruck für Sub-Abfrage (FROM-Teil)</param>
' <param name="Criteria">Filter (WHERE-Teil)</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Insert-SQL-Anweisung ausführen und letzten Identity-Wert (Autowert) zurückgeben
' </summary>
' <param name="InsertSQL">Insert-SQL-Anweisung (INSERT ...)</param>
' <returns>Identity-Wert oder NULL falls kein Datensatz angefügt wurde</returns>
' <remarks>Funktioniert für Jet erst ab Jet 4.0 (Access 2000), IdentityTable ist nur für MSSQL anwendbar, unter MySQL gibt es die Funktion IDENT_CURRENT nicht.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSQL As String, Optional ByVal IdentityTable As String = vbNullString) As Variant

   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim lngRecordsAffected As Long
   Dim strIdentitySQL As String
   
   Dim lngHookResumeMode As Long
   Dim varHookResumeMessage As Variant

   Set cnn = Me.CurrentConnection
   
   If m_HooksEnabled Then
      RaiseEvent HookInsertIdentityReturn(InsertSQL, IdentityTable, cnn, strIdentitySQL, lngHookResumeMode, varHookResumeMessage)
      If lngHookResumeMode <> 0 Then 'Hook ermittelte bereits das Ergebnis
         InsertIdentityReturn = varHookResumeMessage
         Exit Function
      End If
   End If
   
   cnn.Execute InsertSQL, lngRecordsAffected
   
   If lngRecordsAffected = 0 Then
      InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
      Exit Function
   End If

   If Len(strIdentitySQL) = 0 Then
      If Len(IdentityTable) > 0 Then
         strIdentitySQL = "SELECT IDENT_CURRENT('" & IdentityTable & "')" 'wegen trigger .. funktioniert nicht bei allen DBMS!!!
      Else
         strIdentitySQL = "SELECT @@Identity" 'zuletzt zugewiesener Autowert
      End If
   End If

On Error GoTo HandleErr

   Set rst = cnn.Execute(strIdentitySQL)
   With rst
      If .EOF Then
         InsertIdentityReturn = Null
      Else
         InsertIdentityReturn = .Fields(0)
      End If
      .Close
   End With
   Set rst = Nothing
   Set cnn = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function
