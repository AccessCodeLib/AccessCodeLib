VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdbcHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: OdbcHandler (2009-08-30)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-Zugriffsmethoden
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/odbc/OdbcHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/odbc/TEST_OdbcHandlerHook.bas</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit


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


Public Event HookInsertIdentityReturn(ByRef InsertSQL As String, _
                                      ByRef ActiveDb As DAO.Database, ByRef IdentitySelectString As String, _
                                      ByRef HookResumeMode As Long, ByRef HookResumeMessage As Variant)

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

Friend Property Let HooksEnabled(ByVal bEnable As Boolean)
   m_HooksEnabled = bEnable
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb (2009-06-15)
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
      'Error-Event auslösen und hoffen, dass neue Referenz geliefert wird
      RaiseEvent ErrorMissingCurrentDb("OdbcHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
   End If
   Set CurrentDb = m_DaoDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb (2009-06-15)
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
' Property: CurrentDbBE (2009-06-15)
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

On Error GoTo HandleErr

   If (m_DaoDbBE Is Nothing) Then
      tempOdbConnectionstring = ConnectionString
      If Len(tempOdbConnectionstring) > 0 Then
         Set m_DaoDbBE = DBEngine.OpenDatabase("", dbDriverComplete, False, tempOdbConnectionstring)
      End If
   End If
   
   Set CurrentDbBE = m_DaoDbBE

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentDbBE", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

   
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE (2009-06-15)
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

On Error GoTo HandleErr

   closeDaoDbBE
   Set m_DaoDbBE = NewCurrentDbBE
   If Not (NewCurrentDbBE Is Nothing) Then
      m_OdbcConnectionString = NewCurrentDbBE.Connect
   End If

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentDbBE", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property


'---------------------------------------------------------------------------------------
' Property: ConnectionString (2009-06-15)
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

On Error GoTo HandleErr

On Error Resume Next

   If Len(m_OdbcConnectionString) = 0 Then
      If Not (m_DaoDbBE Is Nothing) Then
         m_OdbcConnectionString = m_DaoDbBE.Connect
      End If
      If Len(m_OdbcConnectionString) = 0 Then
         RaiseEvent ErrorEmptyConnectionString("OdbcHandler.ConnectionString", m_OdbcConnectionString)
      End If
   End If
   ConnectionString = m_OdbcConnectionString

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ConnectionString", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

'---------------------------------------------------------------------------------------
' Property: ConnectionString (2009-06-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-ConnectionString
' </summary>
' <param name="sConnectionString">ODBC-Connectionstring für Verbidnungsaufbau</param>
' <remarks>Beispiel: "ODBC;DRIVER={SQL Server};DATABASE=testdb;SERVER=(local)\SQL2005;Trusted_Connection=Yes;"</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ConnectionString(ByVal sConnectionString As String)

On Error GoTo HandleErr

   If m_OdbcConnectionString <> sConnectionString Then
      closeDaoDbBE
   End If
   m_OdbcConnectionString = sConnectionString

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ConnectionString", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property


'---------------------------------------------------------------------------------------
' Function: OpenRecordset (2009-06-15)
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
   
On Error GoTo HandleErr

   If (RecordsetOptions And dbSeeChanges) = 0 Then
      RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If
   Set OpenRecordset = Me.CurrentDbBE.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "OpenRecordset", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

End Function


'---------------------------------------------------------------------------------------
' Function: OpenRecordsetPT (14.02.2009)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OpenRecordsetPT
' </summary>
' <param name="Quelle">SQL-Anweisung</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenForwardOnly)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges+dbSQLPassThrough)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetPT(ByVal sSQL As String, _
                  Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                  Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges Or DAO.RecordsetOptionEnum.dbSQLPassThrough, _
                  Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset

On Error GoTo HandleErr

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
     .SQL = sSQL
     Set rst = .OpenRecordset(RecordsetType, RecordsetOptions, LockEdit)
     .Close
   End With

   Set OpenRecordsetPT = rst
   
ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "OpenRecordsetPT", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
     Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
     Resume Next
   Case Else
     Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: Execute (2009-06-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO.Execute-Anweisung über #CurrentDbBE
' </summary>
' <param name="CommandText">SQL-Anweisung im Jet-SQL-Dialekt</param>
' <param name="RecordsAffected">Anzahl der betroffenen Datensätze</param>
' <param name="Options">Werte von RecordsetOptionEnum analog DAO.Database.Execute</param>
' <remarks>Achtung: Execute-Parameter weichen von DAO-Execute ab, da RecordsAffected ergänzt wurde</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Execute(ByVal CommandText As String, _
          Optional ByRef RecordsAffected As Long, _
          Optional ByVal Options As DAO.RecordsetOptionEnum)

   Dim db As DAO.Database

On Error GoTo HandleErr

   Set db = Me.CurrentDbBE

   If (Options And dbSeeChanges) = 0 Then
      Options = Options + dbSeeChanges
   End If
   db.Execute CommandText, Options
   RecordsAffected = db.RecordsAffected

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "Execute", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
End Sub


'---------------------------------------------------------------------------------------
' Sub: ExecutePT (14.02.2009)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Execute-Anweisung über eine Pass-Through-Abfrage.
' </summary>
' <param name="sSQL">SQL-Anweisung im SQL-Dialekt des DBMS</param>
' <param name="RecordsAffected">Anzahl der betroffenen Datensätze</param>
' <param name="Options">Werte von RecordsetOptionEnum analog DAO.CreateQueryDef.Execute</param>
' <remarks>Achtung: Execute-Parameter weichen von DAO-Execute ab, da RecordsAffected ergänzt wurde</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub ExecutePT(ByVal sSQL As String, _
            Optional ByRef RecordsAffected As Long, _
            Optional ByVal Options As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSQLPassThrough)
         
   Dim db As DAO.Database
   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   Set db = Me.CurrentDb
   
   If (Options And dbSQLPassThrough) = 0 Then
     Options = Options + dbSQLPassThrough
   End If
   
   Set qdf = db.CreateQueryDef("")
   With qdf
     .Connect = Me.ConnectionString
     .ODBCTimeout = 0
     .SQL = sSQL
     .ReturnsRecords = False
     .Execute Options
     .Close
   End With
   
   RecordsAffected = db.RecordsAffected

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExecutePT", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
End Sub


'---------------------------------------------------------------------------------------
' Function: LookupSQL (2009-06-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer kompletten SQL-Anweisung
' </summary>
' <param name="sSQL">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQL(ByVal sSQL As String, Optional ByVal index As Variant = 0&) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   'Alternative:
   'Set rst = Me.OpenRecordsetPT(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         LookupSQL = Null
      Else
         LookupSQL = .Fields(index)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "LookupSQL", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL (2009-06-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer kompletten SQL-Anweisung
' </summary>
' <param name="sSQL">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQLPT(ByVal sSQL As String, Optional ByVal index As Variant = 0&) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordsetPT(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         LookupSQLPT = Null
      Else
         LookupSQLPT = .Fields(index)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "LookupSQLPT", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: Lookup (2009-06-15)
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
Public Function Lookup(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

   Dim strSQL As String

On Error GoTo HandleErr

   strSQL = "SELECT " & Expr & " FROM " & Domain
   If Not (VarType(Criteria) = vbError) Then
      If Len(Criteria) > 0 Then
         strSQL = strSQL & " WHERE " & Criteria
      End If
   End If
   Lookup = Me.LookupSQL(strSQL)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Lookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: DCount (2009-06-15)
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
Public Function DCount(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

   DCount = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DCount", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: DMax (2009-06-15)
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
Public Function DMax(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

   DMax = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DMax", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: DMin (2009-06-15)
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
Public Function DMin(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

    DMin = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DMin", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
    
End Function

'---------------------------------------------------------------------------------------
' Function: DSum (2009-06-15)
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
Public Function DSum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

   DSum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DSum", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn (2009-06-15)
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
Public Function InsertIdentityReturn(ByVal InsertSQL As String) As Variant

   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim lngRecordsAffected As Long
   Dim strIdentitySQL As String
   
   Dim lngHookResumeMode As Long
   Dim varHookResumeMessage As Variant
   
On Error GoTo HandleErr

   Set db = Me.CurrentDbBE
   
   If m_HooksEnabled Then
      RaiseEvent HookInsertIdentityReturn(InsertSQL, db, strIdentitySQL, lngHookResumeMode, varHookResumeMessage)
      If lngHookResumeMode <> 0 Then 'Hook ermittelte bereits das Ergebnis
         InsertIdentityReturn = varHookResumeMessage
         Exit Function
      End If
   End If
     
   db.Execute InsertSQL, dbFailOnError
   lngRecordsAffected = db.RecordsAffected
   
   If lngRecordsAffected = 0 Then
      InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
      Exit Function
   End If
   
   If Len(strIdentitySQL) = 0 Then
      strIdentitySQL = "SELECT @@Identity"
   End If

   Set rst = db.OpenRecordset(strIdentitySQL) 'zuletzt zugewiesener Autowert / ... funktioniert nicht bei MySQL!
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
   Select Case HandleError(Err.Number, "InsertIdentityReturn", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
      Case ACLibErrorResumeMode.aclibErrResume
         Resume
      Case ACLibErrorResumeMode.aclibErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
End Function



'---------------------------------------------------------------------------------------
' Sub: LinkTable (17.02.2009)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Backend-Tabelle im Access-Frontend verknüpfen
' </summary>
' <param name="SourceTableName">Tabelle im Backend</param>
' <param name="LinkedTableName">Tabelle im Frontend</param>
' <param name="bSavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <param name="bRemoveSchemaName">Schema-Kennung entfernen</param>
' <remarks>Sollten eine 2-geteilte Tabellenkennung wie "dbo.Tabelle123" im BE vorliegen und  bRemoveSchemaName auf false gesetzt sein, wird der Punkt durch "_" ersetzt.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub LinkTable(ByVal SourceTableName As String, Optional ByVal LinkedTableName As String = vbNullString, _
            Optional ByVal bSavePWD As Boolean = False, Optional ByVal bRemoveSchemaName As Boolean = True)

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim lngPos As Long
   
On Error GoTo HandleErr

   If Len(Me.ConnectionString) = 0 Then
      Err.Raise ACLibErrorNumbers.ERRNR_NOCONFIG, _
              "LinkTable", "ODBC-Connectionstring fehlt!"
   End If
   
   If Len(LinkedTableName) = 0 Then
      LinkedTableName = SourceTableName
      If bRemoveSchemaName Then
         lngPos = InStr(1, LinkedTableName, ".")
         If lngPos > 0 Then
            LinkedTableName = Mid$(LinkedTableName, lngPos + 1) ' nur Text nach Punkt verwenden
         End If
      End If
      LinkedTableName = Replace(LinkedTableName, ".", "_")
   End If
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   'Prüfen ob Tabelle bereits vorhanden ist.
   For Each tdf In dbs.TableDefs
      If tdf.Name = LinkedTableName Then
         'prüfen ob SourceTableName übereinstimmt
         If tdf.SourceTableName = SourceTableName Then
            RelinkTable LinkedTableName, bSavePWD
            Exit Sub
         Else
            If MsgBox("Der SourceTableName der vorhandene Tabelle '" & LinkedTableName & "' ist '" & tdf.SourceTableName & "' und nicht " & SourceTableName & "'" & vbNewLine & _
                      "Soll die vorhandene Tabelle mit der neuen Datenherkunft verknüpft werden?", vbCritical + vbYesNo + vbDefaultButton2) = vbYes Then
               'Tdf löschen
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
   tdf.Connect = Me.ConnectionString
   dbs.TableDefs.Append tdf
   If bSavePWD Then
      'Passwort in Tabelleneigenschaft "Connect" speichern:
      If (tdf.Attributes And dbAttachSavePWD) = 0 Then
         tdf.Attributes = dbAttachSavePWD
         'Änderungen aktualisieren:
         tdf.RefreshLink
      End If
   End If
   Set tdf = Nothing
   dbs.TableDefs.Refresh
   Set dbs = Nothing
   
ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "LinkTable", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

   
End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkTable (2009-08-30)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verküpfte Tabelle mit möglicher Änderung der Server-Daten neu verknüpfen
' </summary>
' <param name="LinkedTableName">Tabelle im Frontend</param>
' <param name="bSavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <remarks>Die TableDef-Eigenschaften SourceTableName wird nicht geändert</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RelinkTable(ByVal LinkedTableName As String, _
              Optional ByVal bSavePWD As Boolean = False)

   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   
   Dim sConnectionString As String
   
   Dim attr As TableDefAttributeEnum

On Error GoTo HandleErr

   sConnectionString = Me.ConnectionString

   If Len(sConnectionString) = 0 Then
      Err.Raise ACLibErrorNumbers.ERRNR_NOCONFIG, _
              "RelinkTable", "ODBC-Connectionstring fehlt!"
   End If
   
   Set dbs = Me.CurrentDb
   dbs.TableDefs.Refresh
   Set tdf = dbs.TableDefs(LinkedTableName)
   With tdf
      If Left$(.Connect, 4) = "ODBC" Then
         .Connect = sConnectionString
         attr = .Attributes
         If bSavePWD Then
            If (attr And dbAttachSavePWD) = 0 Then
               .Attributes = dbAttachSavePWD
            End If
         Else
            If (attr And dbAttachSavePWD) = dbAttachSavePWD Then
               If (attr And dbAttachedODBC) = dbAttachedODBC Then
                  attr = attr - dbAttachedODBC
               End If
               attr = attr - dbAttachSavePWD
               .Attributes = attr
            End If
         End If
         .RefreshLink
      End If
   End With

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "RelinkTable", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

   
End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkPassThroughQuery (2009-07-27)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Pass-Thorugh-Abfrage neu verknüpfen
' </summary>
' <param name="QueryName">Abfrage im Frontend</param>
' <param name="sConnectionString">Optionaler ODBC-Connectionstring</param>
' <remarks>Mit sConnectionString kann vermieden werden, dass Benutzer u. Passwort gespeichert werden.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RelinkPassThroughQuery(ByVal QueryName As String, _
                         Optional ByVal sConnectionString As String)

   Dim dbs As DAO.Database
   Dim qdf As DAO.QueryDef

On Error GoTo HandleErr

   If Len(sConnectionString) = 0 Then
      sConnectionString = Me.ConnectionString
   End If
   
   If Len(sConnectionString) = 0 Then
      Err.Raise ACLibErrorNumbers.ERRNR_NOCONFIG, _
              "RelinkPassThroughQuery", "ODBC-Connectionstring fehlt!"
   End If
   
   Set dbs = Me.CurrentDb
   dbs.QueryDefs.Refresh
   For Each qdf In dbs.QueryDefs
      With qdf
      If Len(.Connect) > 0 Then
         .Connect = Me.ConnectionString
      End If
      End With
   Next qdf
   
ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "RelinkPassThroughQuery", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

   
End Sub

'---------------------------------------------------------------------------------------
' Sub: RelinkTablesAndQueries (2009-07-27)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Mehrere verknüpfte Tabellen und Pass-Through-Abfragen neu verknüpfen
' </summary>
' <param name="bSavePWD">ODBC-Passwort in Tabellenverknüpfung speichern</param>
' <param name="lEventCounter">Hilfsparamter für Wertübergabe: Anzahl der Verknüpften Elemente für DbObjectRelinked-Ereignisaufruf</param>
' <param name="sConnectionString">Optionaler ODBC-Connectionstring</param>
' <remarks>
'  Je nach Einstellung von lEventCounter wird das Ereignis #DbObjectRelinked ausgelöst (0 = kein Ereignisaufruf, 1 = nach jedem Element Ereignis auslösen, 5 nach jedem 5. Element Ereignis auslösen,...).
'  Mit sConnectionString kann vermieden werden, dass Benutzer u. Passwort gespeichert werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function RelinkTablesAndQueries( _
         Optional ByVal bSavePWD As Boolean = False, _
         Optional ByVal lEventCounter As Long = 0, _
         Optional ByVal sConnectionString As String = vbNullString) As Boolean
   
   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim qdf As DAO.QueryDef
   
   Dim lConnectCounter As Long
   Dim attr As TableDefAttributeEnum
   
On Error GoTo HandleErr

   If Len(sConnectionString) = 0 Then
      sConnectionString = Me.ConnectionString
   End If
   
   If Len(sConnectionString) = 0 Then
      Err.Raise ACLibErrorNumbers.ERRNR_NOCONFIG, _
              "RelinkTablesAndQueries", "ODBC-Connectionstring fehlt!"
   End If

   Set dbs = Me.CurrentDb
   
' Tabellen
   dbs.TableDefs.Refresh
   For Each tdf In dbs.TableDefs
      With tdf
      If Left$(.Connect, 4) = "ODBC" Then
         .Connect = sConnectionString
         .RefreshLink
         If bSavePWD Then
            AddTableDefAttribut tdf, dbAttachSavePWD
         Else
            RemoveTableDefAttribut tdf, dbAttachSavePWD
         End If

         If lEventCounter > 0 Then '0 ... Rückgabe ist deaktiviert
            lConnectCounter = lConnectCounter + 1
            If lConnectCounter Mod lEventCounter = 0 Then
               RaiseEvent DbObjectRelinked(.Name, lConnectCounter)
            End If
         End If

      End If
      End With
   Next tdf

' Abfragen
   dbs.QueryDefs.Refresh
   For Each qdf In dbs.QueryDefs
      With qdf
      If Len(.Connect) > 0 Then
         .Connect = sConnectionString
         If lEventCounter > 0 Then '0 ... Rückgabe ist deaktiviert
            lConnectCounter = lConnectCounter + 1
            If lConnectCounter Mod lEventCounter = 0 Then
               RaiseEvent DbObjectRelinked(.Name, lConnectCounter)
            End If
         End If
      End If
      End With
   Next qdf

   RelinkTablesAndQueries = True

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "RelinkTablesAndQueries", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
        Case ACLibErrorResumeMode.aclibErrResume
            Resume
        Case ACLibErrorResumeMode.aclibErrResumeNext
            Resume Next
        Case Else
            Resume ExitHere
    End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: DeleteOdbcTableDefs (2009-06-28)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle verknüpften Tabelle im Frontend löschen. (Hat keine Auswirkung auf die Backend-Tabellen.)
' </summary>
' <param name="lEventCounter">Hilfsparamter für Wertübergabe: Anzahl der gelöschen Elemente für DbObjectDeleted-Ereignisaufruf</param>
' <returns>Boolean</returns>
' <remarks>
' Je nach Einstellung von lEventCounter wird das Ereignis #DbObjectDeleted ausgelöst
' (0 = kein Ereignisaufruf, 1 = nach jedem Element Ereignis auslösen, 5 nach jedem 5. Element Ereignis auslösen,...).
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DeleteOdbcTableDefs(Optional ByVal lEventCounter As Long = 0) As Boolean
   
   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   Dim i As Long
   Dim tdfCount As Long
   
   Dim lConnectCounter As Long
   Dim sLastDbObjectName As String
   
On Error GoTo HandleErr

   Set dbs = Me.CurrentDb
   
   dbs.TableDefs.Refresh
   tdfCount = dbs.TableDefs.Count - 1
   For i = tdfCount To 0 Step -1
      Set tdf = dbs.TableDefs(i)
      With tdf
      If Left$(.Connect, 4) = "ODBC" Then
         sLastDbObjectName = .Name
         dbs.TableDefs.Delete sLastDbObjectName

         If lEventCounter > 0 Then '0 ... Rückgabe ist deaktiviert
            lConnectCounter = lConnectCounter + 1
            If lConnectCounter Mod lEventCounter = 0 Then
               RaiseEvent DbObjectDeleted(sLastDbObjectName, lConnectCounter)
            End If
         End If

      End If
      End With
   Next

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DeleteOdbcTableDef", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
        Case ACLibErrorResumeMode.aclibErrResume
            Resume
        Case ACLibErrorResumeMode.aclibErrResumeNext
            Resume Next
        Case Else
            Resume ExitHere
    End Select
   
End Function


'Backend-Database-Instanz schließen:
Private Sub closeDaoDbBE()
On Error Resume Next
   If Not (m_DaoDbBE Is Nothing) Then
      m_DaoDbBE.Close
      Set m_DaoDbBE = Nothing
   End If
End Sub

Private Sub RemoveTableDefAttribut(ByRef tdf As DAO.TableDef, attr As TableDefAttributeEnum)

   Dim lngAttributes As Long

On Error GoTo HandleErr

   lngAttributes = tdf.Attributes
   
   If (lngAttributes And attr) = attr Then
   
      If (tdf.Attributes And dbAttachedODBC) = dbAttachedODBC Then
         lngAttributes = lngAttributes - dbAttachedODBC
      ElseIf (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
         lngAttributes = lngAttributes - dbAttachedTable
      End If
      
      'Entfernen
      lngAttributes = lngAttributes - attr
      tdf.Attributes = lngAttributes
      
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "RemoveTableDefAttribut", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub AddTableDefAttribut(ByRef tdf As DAO.TableDef, attr As TableDefAttributeEnum)

   Dim lngAttributes As Long

On Error GoTo HandleErr

   lngAttributes = tdf.Attributes
   
   If (lngAttributes And attr) = 0 Then
   
      If (tdf.Attributes And dbAttachedODBC) = dbAttachedODBC Then
         lngAttributes = lngAttributes Xor dbAttachedODBC
      ElseIf (tdf.Attributes And dbAttachedTable) = dbAttachedTable Then
         lngAttributes = lngAttributes Xor dbAttachedTable
      End If
   
      'Ergänzen
      lngAttributes = lngAttributes Or attr
      tdf.Attributes = lngAttributes
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "AddTableDefAttribut", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub
