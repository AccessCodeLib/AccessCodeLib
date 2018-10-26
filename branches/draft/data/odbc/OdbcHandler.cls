VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdbcHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: OdbcHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' ODBC-Zugriffsmethoden
' </summary>
' <remarks></remarks>
'\ingroup data_odbc
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/odbc/OdbcHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/odbc/OdbcHandlerHookTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ERRNR_NOCONFIG As Long = vbObjectError + 1002

'ODBC-Connectionstring
Private m_OdbcConnectionString As String

'DAO-Database
Private m_DaoDb As DAO.Database    'lokale Database-Instanz ... z. B. zum Erstellen von Temp-QueryDef
Private m_DaoDbBE As DAO.Database  'BE (Database-Instanz über ODBC-Connection)

'Error-Events für mögliche Fehlerbehandlung über die FE-Anwendung
'/**
' <summary>
' Ereignis für fehlende CurrentDb-Referenz
' </summary>
' <param name="Msg">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="NewCurrentDbRef">DAO.Database-Referenz, die für CurrentDb verwendet werden soll</param>
' <remarks>über den ByRef-Parameter NewCurrentDbRef kann die passende DAO.Database-Referenz weitergegeben werden.</remarks>
'**/
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

'/**
' <summary>
' Ereignis für fehlenden ODBC-Connectionstring
' </summary>
' <param name="Msg">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="NewConnectionString">ODBC-Connectionstring</param>
' <remarks>
' Über den ByRef-Parameter NewConnectionString kann der passende ODBC-Connectionstring weitergegeben werden.
' <remarks>
'**/
Public Event ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'Informations-Ereignisse
'/**
' <summary>
' Info über das zuletzt verknüpfte Element
' </summary>
' <param name="LastDbObjectName">Meldung über Ursache der Ereignisauslösung. (Wird noch nicht genutzt)</param>
' <param name="lCounter">Anzahl der verknüpften Elemente seit Aufruf der Methode</param>
' <remarks>Wird in der Methode #RelinkTablesAndQueries genutzt.</remarks>
'**/
Public Event DbObjectRelinked(ByVal LastDbObjectName As String, ByVal lCounter As Long)

'/**
' <summary>
' Info über das zuletzt gelöschte Element
' </summary>
' <param name="LastDbObjectName">Name des zuletzt gelöschten Elements</param>
' <param name="lCounter">Anzahl der gelöschten Elemente seit Aufruf der Methode</param>
' <remarks>Wird in der Methode #DeleteOdbcTableDefs genutzt.</remarks>
'**/
Public Event DbObjectDeleted(ByVal LastDbObjectName As String, ByVal lCounter As Long)

'/**
' <summary>
' Event zum Anfagen von InsertIdentityReturn um die Sql-Anweisung zu ändern bzw. Ergebnis von außen einfließen zu lassen
' </summary>
' <param name="InsertSQL">Insert-SQL-Anweisung (kann mittels Hook geändert werden)</param>
' <param name="ActiveDb">Aktuell genutzte Datenbank</param>
' <param name="IdentitySelectString">SQL-Anweisung zum Auswerten des letzen Identity-Wertes (kann mittels Hook geändert werden)</param>
' <param name="ResponseMode">Anweisung für die Weiterarbeit in der Prozedur</param>
' <param name="ResponseMessage">Rückgabewert</param>
' <remarks>Wird in der Methode #DeleteOdbcTableDefs genutzt.</remarks>
'**/
Public Event HookInsertIdentityReturn(ByRef InsertSql As String, _
                                      ByRef ActiveDb As DAO.Database, _
                                      ByRef IdentitySelectString As String, _
                                      ByRef ResponseMode As Long, ByRef ResponseMessage As Variant)

Private m_HooksEnabled As Boolean

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

On Error Resume Next

   Set m_DaoDb = Nothing
   
   If Not m_DaoDbBE Is Nothing Then
      m_DaoDbBE.Close
      Set m_DaoDbBE = Nothing
   End If

End Sub

Friend Property Get HooksEnabled() As Boolean
   HooksEnabled = m_HooksEnabled
End Property

Friend Property Let HooksEnabled(ByVal NewValue As Boolean)
   m_HooksEnabled = NewValue
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO.Database-Instanz des Frontends bzw. jener Jet-DB in der die Pass-Through-Abfragen erstellt werden sollen
' </summary>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
   If (m_DaoDb Is Nothing) Then
      'Event auslösen, damit neue Referenz geliefert werden kann
      RaiseEvent ErrorMissingCurrentDb("OdbcHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
   End If
   Set CurrentDb = m_DaoDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO.Database-Instanz des Frontends bzw. jener Jet-DB in der die Pass-Through-Abfragen erstellt werden sollen
' </summary>
' <param name="NewCurrentDb">DAO.Database-Referenz die für CurrentDb verwendet werden soll</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
   Set m_DaoDb = NewCurrentDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Database-Referenz zum Backend
' </summary>
' <returns>DAO.Database</returns>
' <remarks>DAO.Database wird über DBEngine.OpenDatabase("", dbDriverComplete, False, <see cref=OdbcHandler#ConnectionString>ConnectionString</see>) geöffnet.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDbBE() As DAO.Database
   Dim tempOdbConnectionstring As String
   If (m_DaoDbBE Is Nothing) Then
      tempOdbConnectionstring = ConnectionString
      If Len(tempOdbConnectionstring) > 0 Then
         Set m_DaoDbBE = DBEngine.OpenDatabase("", dbDriverComplete, False, tempOdbConnectionstring)
      End If
   End If
   Set CurrentDbBE = m_DaoDbBE
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Database-Referenz zum Backend
' </summary>
' <param name="NewCurrentDbBE">DAO.Database-Referenz auf eine bereits geöffnete Datanbank</param>
' <returns>DAO.Database</returns>
' <remarks>Diese Set-Property-Prozedur ist normalerweise nicht erforderlich, da die Datanbank überlicherweise mittels angegebenen ODBC-Connectionstring geöffnet wird.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDbBE(ByRef NewCurrentDbBE As DAO.Database)
   CloseDaoDbBE
   Set m_DaoDbBE = NewCurrentDbBE
   If Not (NewCurrentDbBE Is Nothing) Then
      m_OdbcConnectionString = NewCurrentDbBE.Connect
   End If
End Property


'---------------------------------------------------------------------------------------
' Property: ConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-ConnectionString
' </summary>
' <returns>String</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ConnectionString() As String
   If Len(m_OdbcConnectionString) = 0 Then
      If Not (m_DaoDbBE Is Nothing) Then
On Error Resume Next
         m_OdbcConnectionString = m_DaoDbBE.Connect
On Error GoTo 0
      End If
      If Len(m_OdbcConnectionString) = 0 Then
         RaiseEvent ErrorEmptyConnectionString("OdbcHandler.ConnectionString", m_OdbcConnectionString)
      End If
   End If
   ConnectionString = m_OdbcConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: ConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-ConnectionString
' </summary>
' <param name="Value">ODBC-Connectionstring für Verbindungsaufbau</param>
' <remarks>Beispiel: "ODBC;DRIVER={SQL Server};DATABASE=testdb;SERVER=(local)\SQL2005;Trusted_Connection=Yes;"</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ConnectionString(ByVal Value As String)
   If m_OdbcConnectionString = Value Then
      Exit Property
   End If
   CloseDaoDbBE
   m_OdbcConnectionString = Value
End Property

'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'/**
' <summary>
' openRecordset
' </summary>
' <param name="Quelle">SQL-Anweisung</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenForwardOnly)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                     Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                     Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset
                                         
   If (RecordsetOptions And dbSeeChanges) = 0 Then
      RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If
   Set OpenRecordset = Me.CurrentDbBE.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)

End Function


'---------------------------------------------------------------------------------------
' Function: OpenRecordsetPT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OpenRecordsetPT
' </summary>
' <param name="Source">SQL-Anweisung</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenForwardOnly)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges+dbSQLPassThrough)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetPT(ByVal Source As String, _
                  Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                  Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges Or DAO.RecordsetOptionEnum.dbSQLPassThrough, _
                  Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset

   If (RecordsetOptions And dbSeeChanges) = 0 Then
     RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If
   If (RecordsetOptions And dbSQLPassThrough) = 0 Then
     RecordsetOptions = RecordsetOptions + dbSQLPassThrough
   End If
   
   Set qdf = Me.CurrentDb.CreateQueryDef("")
   With qdf
     .Connect = Me.ConnectionString
     .ODBCTimeout = 0
     .SQL = Source
     Set rst = .OpenRecordset(RecordsetType, RecordsetOptions, LockEdit)
     .Close
   End With

   Set OpenRecordsetPT = rst

End Function

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO.Execute-Anweisung über #CurrentDbBE
' </summary>
' <param name="CommandText">SQL-Anweisung im Jet-SQL-Dialekt</param>
' <param name="Options">Werte von RecordsetOptionEnum analog DAO.Database.Execute</param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal CommandText As String, _
               Optional ByVal Options As DAO.RecordsetOptionEnum) As Long

   Dim db As DAO.Database
   
   Set db = Me.CurrentDbBE

   If (Options And dbSeeChanges) = 0 Then
      Options = Options + dbSeeChanges
   End If
   db.Execute CommandText, Options
   Execute = db.RecordsAffected

End Function


'---------------------------------------------------------------------------------------
' Sub: ExecutePT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Execute-Anweisung über eine Pass-Through-Abfrage.
' </summary>
' <param name="CommandText">SQL-Anweisung im SQL-Dialekt des DBMS</param>
' <param name="Options">Werte von RecordsetOptionEnum analog DAO.CreateQueryDef.Execute</param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecutePT(ByVal CommandText As String, _
                 Optional ByVal Options As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSQLPassThrough _
                         ) As Long
         
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef

   If (Options And dbSQLPassThrough) = 0 Then
     Options = Options + dbSQLPassThrough
   End If
   
   Set db = Me.CurrentDb
   Set qdf = db.CreateQueryDef("")
   With qdf
     .Connect = Me.ConnectionString
     .ODBCTimeout = 0
     .SQL = CommandText
     .ReturnsRecords = False
     .Execute Options
     .Close
   End With
   
   ExecutePT = db.RecordsAffected

End Function


'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer kompletten SQL-Anweisung
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <param name="ValueIfNull">Rückgabewert bei fehlendem Datensatz oder NULL-Feldwert (Standard: Null</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSql(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(SqlText, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         LookupSql = ValueIfNull
      Else
         LookupSql = Nz(.Fields(Index), ValueIfNull)
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

   Err.Raise Err.Number, "LookupSqlPT:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSqlPT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer kompletten SQL-Anweisung
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <param name="ValueIfNull">Rückgabewert bei fehlendem Datensatz oder NULL-Feldwert (Standard: Null</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSqlPT(ByVal SqlText As String, _
                   Optional ByVal Index As Variant = 0&, _
                   Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordsetPT(SqlText, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         LookupSqlPT = ValueIfNull
      Else
         LookupSqlPT = Nz(.Fields(Index), ValueIfNull)
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

   Err.Raise Err.Number, "LookupSqlPT:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

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

   Dim SqlText As String

   SqlText = "SELECT " & Expr & " FROM " & Domain
   If Not (VarType(Criteria) = vbError) Then
      If Len(Criteria) > 0 Then
         SqlText = SqlText & " WHERE " & Criteria
      End If
   End If
   Lookup = Me.LookupSql(SqlText, , ValueIfNull)

End Function

'---------------------------------------------------------------------------------------
' Function: Count
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
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
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
' <remarks>Funktioniert erst ab Jet 4.0 (Access 2000) und ist vom DBMS abhängig. Bei MySQL-Tabellen wird z. B. 0 zurückgegeben</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSql As String) As Variant

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim RecordsAffected As Long
   Dim IdentitySql As String
   Dim HookResponseMode As Long
   Dim HookResponseMessage As Variant

   Set db = Me.CurrentDbBE
   
   If m_HooksEnabled Then
      RaiseEvent HookInsertIdentityReturn(InsertSql, db, IdentitySql, HookResponseMode, HookResponseMessage)
      If HookResponseMode <> 0 Then 'Hook ermittelte bereits das Ergebnis
         InsertIdentityReturn = HookResponseMessage
         Exit Function
      End If
   End If
   
On Error GoTo HandleErr

   db.Execute InsertSql, dbFailOnError
   RecordsAffected = db.RecordsAffected
   
   If RecordsAffected = 0 Then
      InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
      Exit Function
   End If
   
   If Len(IdentitySql) = 0 Then
      IdentitySql = "SELECT @@Identity"
   End If

   Set rst = db.OpenRecordset(IdentitySql) 'zuletzt zugewiesener Autowert / ... funktioniert nicht bei MySQL!
   With rst
      If .EOF Then
         InsertIdentityReturn = Null
      Else
         InsertIdentityReturn = .Fields(0)
      End If
      .Close
   End With
   Set rst = Nothing
   Set db = Nothing
    
ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Sub: LinkTable
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Backend-Tabelle im Access-Frontend verknüpfen
' </summary>
' <param name="SourceTableName">Tabelle im Backend</param>
' <param name="LinkedTableName">Tabelle im Frontend</param>
' <param name="SavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <param name="RemoveSchemaName">Schema-Kennung entfernen</param>
' <remarks>Sollten eine 2-geteilte Tabellenkennung wie "dbo.Tabelle123" im BE vorliegen und  bRemoveSchemaName auf false gesetzt sein, wird der Punkt durch "_" ersetzt.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub LinkTable(ByVal SourceTableName As String, _
            Optional ByVal LinkedTableName As String = vbNullString, _
            Optional ByVal SavePWD As Boolean = False, _
            Optional ByVal RemoveSchemaName As Boolean = True)

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim Pos As Long
   Dim NewConnectionString As String
   
   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "LinkTable", "ODBC-Connectionstring fehlt!"
   End If
   
   If Len(LinkedTableName) = 0 Then
      LinkedTableName = SourceTableName
      If RemoveSchemaName Then
         Pos = InStr(1, LinkedTableName, ".")
         If Pos > 0 Then
            LinkedTableName = Mid$(LinkedTableName, Pos + 1) ' nur Text nach Punkt verwenden
         End If
      End If
      LinkedTableName = Replace(LinkedTableName, ".", "_")
   End If
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   'Prüfen ob Tabelle bereits vorhanden und verknüpft ist.
   For Each tdf In dbs.TableDefs
      If tdf.Name = LinkedTableName Then
         If tdf.SourceTableName = SourceTableName Then
            RelinkTableDef tdf, NewConnectionString, SavePWD
            Exit Sub
         Else
            If MsgBox("Der SourceTableName der vorhandene Tabelle '" & LinkedTableName & "' ist '" & tdf.SourceTableName & "' und nicht " & SourceTableName & "'" & vbNewLine & _
                      "Soll die vorhandene Tabelle mit der neuen Datenherkunft verknüpft werden?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
               dbs.TableDefs.Delete tdf.Name
            Else
               Exit Sub
            End If
         End If
         Exit For
      End If
   Next tdf

   Set tdf = dbs.CreateTableDef(LinkedTableName)
   tdf.SourceTableName = SourceTableName
   tdf.Connect = NewConnectionString
   dbs.TableDefs.Append tdf
   If SavePWD Then
      AddTableDefAttribut tdf, dbAttachSavePWD
      tdf.RefreshLink
   End If
   
   Set tdf = Nothing
   dbs.TableDefs.Refresh
   Set dbs = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Function: RelinkTables
'---------------------------------------------------------------------------------------
'/**
' <summary>
' RelinkTables
' </summary>
' <param name="SavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <param name="EventPeriod">Hilfsparamter für Wertübergabe: Anzahl der Verknüpften Elemente für DbObjectRelinked-Ereignisaufruf</param>
' <returns>Long (Anzahl der verknüpften Tabellen)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkTables(Optional ByVal SavePWD As Boolean = False, _
                             Optional ByVal EventPeriod As Long = 0) As Long
'EventPeriod = 0 ... Rückgabe ist deaktiviert

   Dim NewConnectionString As String
   Dim RelinkedItemCnt As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "ODBC-Connectionstring fehlt!"
   End If
   
   RelinkTablesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedItemCnt
   
   RelinkTables = RelinkedItemCnt

End Function

Private Sub RelinkTablesWorker(ByVal NewConnectionString As String, _
                               ByVal SavePWD As Boolean, _
                               ByVal EventPeriod As Long, _
                               ByRef RelinkedTableCounter As Long)
'EventPeriod = 0 ... Rückgabe ist deaktiviert

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   
   For Each tdf In dbs.TableDefs
   
      If RelinkTableDef(tdf, NewConnectionString, SavePWD) Then
         RelinkedTableCounter = RelinkedTableCounter + 1
         If EventPeriod > 0 Then
            If RelinkedTableCounter Mod EventPeriod = 0 Then
               RaiseEvent DbObjectRelinked(tdf.Name, RelinkedTableCounter)
            End If
         End If
      End If
   
   Next tdf

End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkTable
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verküpfte Tabelle mit möglicher Änderung der Server-Daten neu verknüpfen
' </summary>
' <param name="LinkedTableName">Tabelle im Frontend</param>
' <param name="SavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <returns>Boolean</returns>
' <remarks>Die TableDef-Eigenschaften SourceTableName wird nicht geändert</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkTable(ByVal LinkedTableName As String, _
                   Optional ByVal SavePWD As Boolean = False) As Boolean

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim NewConnectionString As String
   
   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "ODBC-Connectionstring fehlt!"
   End If

   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   Set tdf = dbs.TableDefs(LinkedTableName)
   RelinkTable = RelinkTableDef(tdf, NewConnectionString, SavePWD)

End Function

Private Function RelinkTableDef(ByVal TdfRef As DAO.TableDef, _
                                ByVal NewConnectionString As String, _
                                ByVal SavePWD As String) As Boolean
   
   If Not IsLinkedTable(TdfRef) Then
      RelinkTableDef = False
      Exit Function
   End If
   
   With TdfRef
   
      If Left(.Connect, 5) <> "ODBC;" Then Exit Function
   
      .Connect = NewConnectionString
      .RefreshLink
      
      If SavePWD Then
         AddTableDefAttribut TdfRef, dbAttachSavePWD
      Else
         RemoveTableDefAttribut TdfRef, dbAttachSavePWD
      End If
   End With
   
   RelinkTableDef = True
   
End Function

'---------------------------------------------------------------------------------------
' Function: RelinkPassThroughQueries
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle Pass-Through-Abfragen neu verknüpfen
' </summary>
' <param name="SavePWD">ODBC-Passwort in QueryDef speichern</param>
' <param name="EventPeriod">Hilfsparamter für Wertübergabe: Anzahl der Verknüpften Elemente für DbObjectRelinked-Ereignisaufruf</param>
' <returns>Long (Anzahl der verknüpften Abfragen)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkPassThroughQueries(Optional ByVal SavePWD As Boolean = False, _
                                         Optional ByVal EventPeriod As Long = 0) As Long
'EventPeriod = 0 ... Rückgabe ist deaktiviert

   Dim NewConnectionString As String
   Dim RelinkedItems As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "ODBC-Connectionstring fehlt!"
   End If
   
   RelinkPassThroughQueriesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedItems
   
   RelinkPassThroughQueries = RelinkedItems

End Function

Private Sub RelinkPassThroughQueriesWorker(ByVal NewConnectionString As String, _
                                          ByVal SavePWD As Boolean, _
                                          ByVal EventPeriod As Long, _
                                          ByRef RelinkCounter As Long)
'EventPeriod = 0 ... Rückgabe ist deaktiviert
   
   Dim dbs As DAO.Database
   Dim qdf As DAO.QueryDef
   
   Set dbs = Me.CurrentDb
   dbs.QueryDefs.Refresh
   
   If Not SavePWD Then
      NewConnectionString = RemovePasswordFromConnectionString(NewConnectionString)
   End If
   
   For Each qdf In dbs.QueryDefs
      If RelinkPassThroughQueryDef(qdf, NewConnectionString) Then
         RelinkCounter = RelinkCounter + 1
         If EventPeriod > 0 Then
            If RelinkCounter Mod EventPeriod = 0 Then
               RaiseEvent DbObjectRelinked(qdf.Name, RelinkCounter)
            End If
         End If
      End If
      qdf.Close
   Next qdf

End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkPassThroughQuery
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Pass-Through-Abfrage neu verknüpfen
' </summary>
' <param name="QueryName">Abfragename im Frontend</param>
' <param name="SavePWD">ODBC-Passwort in QueryDef speichern</param>
' <remarks>Mit ConnectionString kann vermieden werden, dass Benutzer u. Passwort gespeichert werden.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkPassThroughQuery(ByVal QueryName As String, _
                              Optional ByVal SavePWD As Boolean = False) As Boolean

   Dim dbs As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim NewConnectionString As String

   NewConnectionString = Me.ConnectionString
   
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkPassThroughQuery", "ODBC-Connectionstring fehlt!"
   End If
   
   If Not SavePWD Then
      NewConnectionString = RemovePasswordFromConnectionString(NewConnectionString)
   End If
   
   Set dbs = Me.CurrentDb
   dbs.QueryDefs.Refresh
   Set qdf = dbs.QueryDefs(QueryName)
   RelinkPassThroughQuery = RelinkPassThroughQueryDef(qdf, NewConnectionString)
   qdf.Close

End Function

Private Function RelinkPassThroughQueryDef(ByVal QdfRef As DAO.QueryDef, _
                                           ByVal NewConnectionString As String) As Boolean
   
   If Not IsPassThroughQuery(QdfRef) Then
      RelinkPassThroughQueryDef = False
      Exit Function
   End If
   
   QdfRef.Connect = NewConnectionString
   RelinkPassThroughQueryDef = True
   
End Function

'---------------------------------------------------------------------------------------
' Sub: RelinkTablesAndQueries
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle verknüpften Tabellen und Pass-Through-Abfragen neu verknüpfen
' </summary>
' <param name="SavePWD">ODBC-Passwort in Tabellenverknüpfung bzw. QueryDef speichern</param>
' <param name="EventPeriod">Hilfsparamter für Wertübergabe: Anzahl der Verknüpften Elemente für DbObjectRelinked-Ereignisaufruf</param>
' <returns>Long (Anzahl der verknüpften Tabellen und Abfragen)</returns>
' <remarks>
'  Je nach Einstellung von EventPeriod wird das Ereignis #DbObjectRelinked ausgelöst (0 = kein Ereignisaufruf, 1 = nach jedem Element Ereignis auslösen, 5 nach jedem 5. Element Ereignis auslösen,...).
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkTablesAndQueries( _
                        Optional ByVal SavePWD As Boolean = False, _
                        Optional ByVal EventPeriod As Long = 0) As Long
   
   Dim NewConnectionString As String
   Dim RelinkedTableCounter As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTablesAndQueries", "ODBC-Connectionstring fehlt!"
   End If
   
   'Tabellen
   RelinkTablesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedTableCounter
   
   'Pass-Through-Abfragen
   RelinkPassThroughQueriesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedTableCounter

   RelinkTablesAndQueries = RelinkedTableCounter

End Function

'---------------------------------------------------------------------------------------
' Function: DeleteOdbcTableDefs
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle verknüpften Tabelle im Frontend löschen. (Hat keine Auswirkung auf die Backend-Tabellen.)
' </summary>
' <param name="lEventCounter">Hilfsparamter für Wertübergabe: Anzahl der gelöschen Elemente für DbObjectDeleted-Ereignisaufruf</param>
' <returns>Long (Anzahl der gelöschten Tabellen)</returns>
' <remarks>
' Je nach Einstellung von lEventCounter wird das Ereignis #DbObjectDeleted ausgelöst
' (0 = kein Ereignisaufruf, 1 = nach jedem Element Ereignis auslösen, 5 nach jedem 5. Element Ereignis auslösen,...).
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DeleteOdbcTableDefs(Optional ByVal EventPeriod As Long = 0) As Long
   
   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim i As Long
   Dim TdfCount As Long
   Dim DeleteCounter As Long
   Dim LastDbObjectName As String

   Set dbs = Me.CurrentDb
      
   dbs.TableDefs.Refresh
   TdfCount = dbs.TableDefs.Count - 1
   For i = TdfCount To 0 Step -1
      Set tdf = dbs.TableDefs(i)
      If IsLinkedTable(tdf) Then
         LastDbObjectName = tdf.Name
         dbs.TableDefs.Delete LastDbObjectName
         DeleteCounter = DeleteCounter + 1
         If EventPeriod > 0 Then '0 ... Rückgabe ist deaktiviert
            If DeleteCounter Mod EventPeriod = 0 Then
               RaiseEvent DbObjectDeleted(LastDbObjectName, DeleteCounter)
            End If
         End If
      End If
   Next
   
   DeleteOdbcTableDefs = DeleteCounter

End Function

Private Sub CloseDaoDbBE()
On Error Resume Next
   If Not (m_DaoDbBE Is Nothing) Then
      m_DaoDbBE.Close
      Set m_DaoDbBE = Nothing
   End If
End Sub

Private Sub RemoveTableDefAttribut(ByRef TdfRef As DAO.TableDef, ByRef Attr As TableDefAttributeEnum)

   Dim AttributesValue As Long

   AttributesValue = TdfRef.Attributes
   If (AttributesValue And Attr) = Attr Then
   
      If (TdfRef.Attributes And dbAttachedODBC) = dbAttachedODBC Then
         AttributesValue = AttributesValue - dbAttachedODBC
      ElseIf (TdfRef.Attributes And dbAttachedTable) = dbAttachedTable Then
         AttributesValue = AttributesValue - dbAttachedTable
      End If
      
      'Entfernen
      AttributesValue = AttributesValue - Attr
      TdfRef.Attributes = AttributesValue
      
   End If

End Sub

Private Sub AddTableDefAttribut(ByRef TdfRef As DAO.TableDef, ByRef Attr As TableDefAttributeEnum)

   Dim AttributesValue As Long

   AttributesValue = TdfRef.Attributes
   If (AttributesValue And Attr) = 0 Then
   
      If (TdfRef.Attributes And dbAttachedODBC) = dbAttachedODBC Then
         AttributesValue = AttributesValue Xor dbAttachedODBC
      ElseIf (TdfRef.Attributes And dbAttachedTable) = dbAttachedTable Then
         AttributesValue = AttributesValue Xor dbAttachedTable
      End If
   
      'Ergänzen
      AttributesValue = AttributesValue Or Attr
      TdfRef.Attributes = AttributesValue
   End If

End Sub

Public Function IsLinkedTable(ByVal TableToCheck As DAO.TableDef) As Boolean
   IsLinkedTable = (Len(Trim$(TableToCheck.Connect)) > 0)
End Function

Public Function IsPassThroughQuery(ByVal QueryToCheck As DAO.QueryDef) As Boolean
   IsPassThroughQuery = (Len(Trim$(QueryToCheck.Connect)) > 0)
End Function

Private Function RemovePasswordFromConnectionString(ByVal OdbcConnectionString As String) As String
   
   Dim OdbcSettings() As String
   Dim OdbcSetting As String
   Dim i As Long
   
   OdbcSettings = Split(OdbcConnectionString, ";")
   For i = 0 To UBound(OdbcSettings)
      OdbcSetting = Trim(OdbcSettings(i))
      If Left(OdbcSetting, 4) = "PWD=" Then
         OdbcSettings(i) = vbNullString
      ElseIf Left(OdbcSetting, 4) = "UID=" Then
         OdbcSettings(i) = vbNullString
      End If
   Next
   RemovePasswordFromConnectionString = Join(OdbcSettings, ";")
   
End Function
