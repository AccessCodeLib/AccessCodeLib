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
' Class: data.odbc.OdbcHandler
'---------------------------------------------------------------------------------------
'
' ODBC data connection methods
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

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

'ODBC connection string
Private m_OdbcConnectionString As String

'DAO Database
Private m_DaoDb As DAO.Database    'local database instance ... e.g. to create Temp-QueryDef
Private m_DaoDbBE As DAO.Database  'Backend (database instance via ODBC connection)

'#############################################################
' Group: Events

'---------------------------------------------------------------------------------------
' Event: ErrorMissingCurrentDb
'---------------------------------------------------------------------------------------
' Event for missing CurrentDb reference
'
' Parameters:
'     Msg             - Message about the cause of the event triggering. (Not yet used)
'     NewCurrentDbRef - DAO.Database reference to be used for CurrentDb
'
' Remarks:
'     The required DAO.Database reference can be passed on via the ByRef parameter NewCurrentDbRef.
'
'---------------------------------------------------------------------------------------
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

'---------------------------------------------------------------------------------------
' Event: ErrorEmptyConnectionString
'---------------------------------------------------------------------------------------
'
' Event for missing ODBC connection string
'
' Parameters:
'     Msg                  - Message about the cause of the event triggering. (Not yet used)
'     NewConnectionString  - ODBC Connectionstring
'
' Remarks:
'     The appropriate ODBC connection string can be passed on via the ByRef parameter NewConnectionString.
'
'---------------------------------------------------------------------------------------
Public Event ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'---------------------------------------------------------------------------------------
' Event: DbObjectRelinked
'---------------------------------------------------------------------------------------
'
' Info about the last linked element
'
' Parameters:
'     LastDbObjectName  - the last linked table or view
'     Counter           - Number of linked elements since the method was called
'
' Remarks:
'     Used in the RelinkTablesAndQueries method.
'
'---------------------------------------------------------------------------------------
Public Event DbObjectRelinked(ByVal LastDbObjectName As String, ByVal Counter As Long)

'---------------------------------------------------------------------------------------
' Event: DbObjectDeleted
'---------------------------------------------------------------------------------------
'
' Info about the last deleted element
'
' Parameters:
'     LastDbObjectName  - the last deleted table or view
'     Counter           - Number of deleted elements since the method was called
'
' Remarks:
'     Used in the RelinkTablesAndQueries method.
'
'---------------------------------------------------------------------------------------
Public Event DbObjectDeleted(ByVal LastDbObjectName As String, ByVal Counter As Long)

'---------------------------------------------------------------------------------------
' Event: HookInsertIdentityReturn
'---------------------------------------------------------------------------------------
'
' Event for missing OLEDB connection string
'
' Parameters:
'     InsertSql            - SQL text of insert statement
'     ActiveDb             - ADODB Connection to use
'     IdentitySelectString - Input string to use
'     HookResponseMode     - resume action after event call: 0 ... code resume after event, 1 exit function after event
'     HookResponseMessage  - return value of function (e.g. InsertIdentityReturn) when HookResponseMode <> 0
'
' Remarks:
'     New values can be passed on via the ByRef parameters.
'
'---------------------------------------------------------------------------------------
Public Event HookInsertIdentityReturn(ByRef InsertSql As String, _
                                      ByRef ActiveDb As DAO.Database, _
                                      ByRef IdentitySelectString As String, _
                                      ByRef ResponseMode As Long, ByRef HookResponseMessage As Variant)

Private m_HooksEnabled As Boolean

'#############################################################
' Group: Class control

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

'---------------------------------------------------------------------------------------
' Property: HooksEnabled
'---------------------------------------------------------------------------------------
' Active Hooks
'
' Remarks:
'     Used in InsertIdentityReturn
'
'---------------------------------------------------------------------------------------
Friend Property Get HooksEnabled() As Boolean
   HooksEnabled = m_HooksEnabled
End Property

Friend Property Let HooksEnabled(ByVal NewValue As Boolean)
   m_HooksEnabled = NewValue
End Property

'#############################################################
' Group: Data Connection

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'
' DAO.Database-Instanz des Frontends bzw. jener Jet-DB in der die Pass-Through-Abfragen erstellt werden sollen
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
   If (m_DaoDb Is Nothing) Then
      'Event auslösen, damit neue Referenz geliefert werden kann
      RaiseEvent ErrorMissingCurrentDb("OdbcHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
   End If
   Set CurrentDb = m_DaoDb
End Property

Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
   Set m_DaoDb = NewCurrentDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'
' Database-Referenz zum Backend
'
' Remarks
'     DAO.Database wird über DBEngine.OpenDatabase("", dbDriverComplete, False, <ConnectionString>) geöffnet.
'
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
'
' ODBC Connection string
'
' Remarks:
'     Example - "ODBC;DRIVER={SQL Server};DATABASE=testdb;SERVER=(local)\SQL2005;Trusted_Connection=Yes;"
'
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
'
' Open DAO.Recordset
'
' Parameters:
'     Source            - SQL statement or table or query name
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
' Remarks:
'     Uses <CurrentDbBE> to open the recordset.
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
'
' Open Pass Through DAO.Recordset
'
' Parameters:
'     Source            - SQL statement or table or query name
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges+dbSQLPassThrough)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
' Remarks:
'     Uses <CurrentDbBE> to open the recordset.
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
'
' Execute SQL statement (CurrentDbBE.Execute)
'
' Parameters:
'     CommandText - SQL statement (String)
'     Options     - DAO.RecordsetOptionEnum
'
' Returns:
'     RecordsAffected (Long)
'
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
'---------------------------------------------------------------------------------------'
' Execute SQL statement with Pass Trough Query
'
' Parameters:
'     CommandText - SQL statement (String)
'     Options     - DAO.RecordsetOptionEnum (Default: DAO.RecordsetOptionEnum.dbSQLPassThrough)
'
' Returns:
'     RecordsAffected (Long)
'
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
'
' Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement) via CurrentDbBE
'
' Parameters:
'     SqlText     - SQL statement
'     Index       - Field selection if the first field of the recordset is not to be returned.
'                   The index is passed in the same way as ADODB.Recordset.Fields(index).
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant  - Null, if SQL statement does not return a record.
'
' Remarks:
'     The SQL statement must be written in the Jet/ACE SQL dialect.
'
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
'
' Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement) via Pass Through Query
'
' Parameters:
'     SqlText     - SQL statement
'     Index       - Field selection if the first field of the recordset is not to be returned.
'                   The index is passed in the same way as ADODB.Recordset.Fields(index).
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant  - Null, if SQL statement does not return a record.
'
' Remarks:
'     The SQL statement must be written in the SQL dialect of backend (e. g. T-SQL for SQL-Server)
'
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
'
' DAO.Recordset replacement function for DLookup (via CurrentDbBE)
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant
'
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
'
' DAO.Recordset replacement function for DCount (via CurrentDbBE)
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (optional) Criteria string (WHERE part)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)
End Function

'---------------------------------------------------------------------------------------
' Function: Max
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DMax (via CurrentDbBE)
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Max(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Min
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DMin (via CurrentDbBE)
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Sum
'---------------------------------------------------------------------------------------
'
' DAO.Recordset replacement function for DSum (via CurrentDbBE)
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn
'---------------------------------------------------------------------------------------
'
' Execute insert SQL statement and return last identity value (via CurrentDbBE)
'
' Parameters:
'     InsertSQL   - Insert SQL statement
'
' Returns:
'     Identity value or NULL if no record was appended
'
' Remarks:
'     Works for Jet only from Jet 4.0 (Access 2000), IdentityTable is only applicable for MSSQL, under MySQL the IDENT_CURRENT function does not exist.
'     If HooksEnabled: event HookInsertIdentityReturn will be raised
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

'#############################################################
' Group: Linked Table Support

'---------------------------------------------------------------------------------------
' Sub: LinkTable
'---------------------------------------------------------------------------------------
'
' Link backend table in Access frontend
'
' Parameters:
'     SourceTableName   - Table or view in the backend
'     LinkedTableName   - Table in frontend
'     SavePWD           - Save ODBC password in linked table
'     RemoveSchemaName  - Remove schema identifier
'
' Remarks:
'     If there is a 2-part table identifier like "dbo.Table123" in the BE and bRemoveSchemaName is set to false, the dot is replaced by "_".
'
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
              "LinkTable", "Missing ODBC connection string!"
   End If
   
   If Len(LinkedTableName) = 0 Then
      LinkedTableName = SourceTableName
      If RemoveSchemaName Then
         Pos = InStr(1, LinkedTableName, ".")
         If Pos > 0 Then
            LinkedTableName = Mid$(LinkedTableName, Pos + 1) ' use only text after point
         End If
      End If
      LinkedTableName = Replace(LinkedTableName, ".", "_")
   End If
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   'Check if table already exists and is linked.
' TODO: refactor for-each
   For Each tdf In dbs.TableDefs
      If tdf.Name = LinkedTableName Then
         If tdf.SourceTableName = SourceTableName Then
            RelinkTableDef tdf, NewConnectionString, SavePWD
            Exit Sub
         Else
            If MsgBox("The SourceTableName of the existing table '" & LinkedTableName & "' is '" & tdf.SourceTableName & "' and not '" & SourceTableName & "'" & vbNewLine & _
                      "Should the existing table be linked to the new data origin?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
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
'
' Relink existing tables
'
' Parameters:
'     SavePWD     - Save ODBC password in linked table
'     EventPeriod - Number of linked elements when the DbObjectRelinked event is to be called. (Default: 0 ... no event call).
'
' Returns:
'     Long  - Number of linked tables
'
'---------------------------------------------------------------------------------------
Public Function RelinkTables(Optional ByVal SavePWD As Boolean = False, _
                             Optional ByVal EventPeriod As Long = 0) As Long
'EventPeriod = 0 ... Return is disabled

   Dim NewConnectionString As String
   Dim RelinkedItemCnt As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "Missing ODBC connection string!"
   End If
   
   RelinkTablesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedItemCnt
   
   RelinkTables = RelinkedItemCnt

End Function

Private Sub RelinkTablesWorker(ByVal NewConnectionString As String, _
                               ByVal SavePWD As Boolean, _
                               ByVal EventPeriod As Long, _
                               ByRef RelinkedTableCounter As Long)
'EventPeriod = 0 ... Return is disabled

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim LastTdfName As String
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   
   For Each tdf In dbs.TableDefs
   
      If RelinkTableDef(tdf, NewConnectionString, SavePWD) Then
         RelinkedTableCounter = RelinkedTableCounter + 1
         If EventPeriod > 0 Then
            LastTdfName = tdf.Name
            If RelinkedTableCounter Mod EventPeriod = 0 Then
               RaiseEvent DbObjectRelinked(LastTdfName, RelinkedTableCounter)
            End If
         End If
      End If
   
   Next tdf
   
   If EventPeriod > 0 Then
      If RelinkedTableCounter Mod EventPeriod > 0 Then
         RaiseEvent DbObjectRelinked(LastTdfName, RelinkedTableCounter)
      End If
   End If

End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkTable
'---------------------------------------------------------------------------------------
'
' Relink linked table with possible change of server data
'
' Parameters:
'     LinkedTableName   - Table in frontend
'     SavePWD           - Save ODBC password in linked table
'
' Returns:
'     Boolean  - success
'
' Remarks:
'     The TableDef property SourceTableName is not changed.
'
'---------------------------------------------------------------------------------------
Public Function RelinkTable(ByVal LinkedTableName As String, _
                   Optional ByVal SavePWD As Boolean = False) As Boolean

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim NewConnectionString As String
   
   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "Missing ODBC connection string!"
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
'
' Alle Pass-Through-Abfragen neu verknüpfen
'
' Parameters:
'     SavePWD     - Save ODBC password in linked table
'     EventPeriod - Number of linked elements when the DbObjectRelinked event is to be called. (Default: 0 ... no event call).
'
' Returns:
'     Long        - Number of linked pass through queries
'
'---------------------------------------------------------------------------------------
Public Function RelinkPassThroughQueries(Optional ByVal SavePWD As Boolean = False, _
                                         Optional ByVal EventPeriod As Long = 0) As Long
'EventPeriod = 0 ... Return is disabled

   Dim NewConnectionString As String
   Dim RelinkedItems As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTable", "Missing ODBC connection string!"
   End If
   
   RelinkPassThroughQueriesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedItems
   
   RelinkPassThroughQueries = RelinkedItems

End Function

Private Sub RelinkPassThroughQueriesWorker(ByVal NewConnectionString As String, _
                                          ByVal SavePWD As Boolean, _
                                          ByVal EventPeriod As Long, _
                                          ByRef RelinkCounter As Long)
'EventPeriod = 0 ... Return is disabled
   
   Dim dbs As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim LastQdfName As String
   
   Set dbs = Me.CurrentDb
   dbs.QueryDefs.Refresh
   
   If Not SavePWD Then
      NewConnectionString = RemovePasswordFromConnectionString(NewConnectionString)
   End If
   
   For Each qdf In dbs.QueryDefs
      If RelinkPassThroughQueryDef(qdf, NewConnectionString) Then
         RelinkCounter = RelinkCounter + 1
         If EventPeriod > 0 Then
            LastQdfName = qdf.Name
            If RelinkCounter Mod EventPeriod = 0 Then
               RaiseEvent DbObjectRelinked(LastQdfName, RelinkCounter)
            End If
         End If
      End If
      qdf.Close
   Next qdf
   
   If EventPeriod > 0 Then
      If RelinkCounter Mod EventPeriod > 0 Then
         RaiseEvent DbObjectRelinked(LastQdfName, RelinkCounter)
      End If
   End If

End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkPassThroughQuery
'---------------------------------------------------------------------------------------
'
' Relink pass through query
'
' Parameters:
'     QueryName   - Table in frontend
'     SavePWD     - Save ODBC password in pass throug query
'
' Returns:
'     Boolean  - success
'
'---------------------------------------------------------------------------------------
Public Function RelinkPassThroughQuery(ByVal QueryName As String, _
                              Optional ByVal SavePWD As Boolean = False) As Boolean

   Dim dbs As DAO.Database
   Dim qdf As DAO.QueryDef
   Dim NewConnectionString As String

   NewConnectionString = Me.ConnectionString
   
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkPassThroughQuery", "Missing ODBC connection string!"
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
'
' Relink all linked tables and pass-through queries
'
' Parameters:
'     SavePWD     - Save ODBC password in linked table
'     EventPeriod - Number of linked elements when the DbObjectRelinked event is to be called. (Default: 0 ... no event call).
'
' Returns:
'     Long  - Number of linked tables and queries
'
'---------------------------------------------------------------------------------------
Public Function RelinkTablesAndQueries( _
                        Optional ByVal SavePWD As Boolean = False, _
                        Optional ByVal EventPeriod As Long = 0) As Long
   
   Dim NewConnectionString As String
   Dim RelinkedCounter As Long

   NewConnectionString = Me.ConnectionString
   If Len(NewConnectionString) = 0 Then
      Err.Raise ERRNR_NOCONFIG, _
              "RelinkTablesAndQueries", "Missing ODBC connection string!"
   End If
   
   'Tables
   RelinkTablesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedCounter
   
   'Pass through queries
   RelinkPassThroughQueriesWorker NewConnectionString, SavePWD, EventPeriod, RelinkedCounter

   RelinkTablesAndQueries = RelinkedCounter

End Function

'---------------------------------------------------------------------------------------
' Function: DeleteOdbcTableDefs
'---------------------------------------------------------------------------------------
'
' Delete all linked table in the frontend. (Has no effect on the backend tables).
'
' Parameters:
'     EventPeriod - Number of linked elements when the DbObjectRelinked event is to be called. (Default: 0 ... no event call).
'
' Returns:
'     Long  - Number of deleted tables
'
' Remarks:
'     Depending on the setting of EventPeriod the event DbObjectDeleted is triggered
'
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

'---------------------------------------------------------------------------------------
' Function: IsLinkedTable
'---------------------------------------------------------------------------------------
'
' Check if TableDef is a linked table
'
' Parameters:
'     TableToCheck   - TableDef reference to Access table in frontend
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function IsLinkedTable(ByVal TableToCheck As DAO.TableDef) As Boolean
   IsLinkedTable = (Len(Trim$(TableToCheck.Connect)) > 0)
End Function

'---------------------------------------------------------------------------------------
' Function: IsPassThroughQuery
'---------------------------------------------------------------------------------------
'
' Check if QueryDef is a pass through query
'
' Parameters:
'     QueryToCheck   - QueryDef reference to Access query in frontend
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
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
