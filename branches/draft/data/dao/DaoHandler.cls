VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: DaoHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' DAO-Zugriffsmethoden
' </summary>
' <remarks></remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'DAO-Database
Private m_DaoDb As DAO.Database

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
   Set m_DaoDb = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentDbC-Variante
' </summary>
' <returns>DAO.Database</returns>
' <remarks>
' CurrentDbC-Variante, die zum Unterschied zur Original-Version von Michael Kaplan
' die DAO-Database-Referenz nicht automatisch auf Access.Application.CurrentDb setzt,
' sondern über ein Ereignis anfordert, falls noch keine Referenz vorhanden ist.
' Dadurch kann über die Hauptanwendung eine belibige Database-Referenz übergeben werden.
' (z. B. falls eine Temp-Datenbank als CurrentDb dienen soll)</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
    If (m_DaoDb Is Nothing) Then
        'Error-Event auslösen und hoffen, dass neue Referenz geliefert wird
        RaiseEvent ErrorMissingCurrentDb("DaoHandler.CurrentDb: unavailable CurrentDb-Database", m_DaoDb)
    End If
    Set CurrentDb = m_DaoDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Übergabe der DAO-Database-Referenz
' </summary>
' <param name="NewCurrentDb">Database-Referenz, die als "CurrentDb" verwendet werden soll</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
    Set m_DaoDb = NewCurrentDb
End Property

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'/**
' <summary>
' SQL-Anweisung ausführen
' </summary>
' <param name="CommandText">SQL-Anweisung</param>
' <param name="Options"></param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal Query As String, _
          Optional ByVal Options As DAO.RecordsetOptionEnum) As Long

    If (Options And dbSeeChanges) = 0 Then
        Options = Options + dbSeeChanges
    End If
    Me.CurrentDb.Execute Query, Options
    Execute = Me.CurrentDb.RecordsAffected
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteParamSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sql-Anweisung mit Parametern ausführen
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="ParamValues">Parmeterwerte in passender Reihenfolge</param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteParamSQL(ByVal SqlText As String, _
                           ParamArray ParamValues() As Variant) As Long
 
   Dim qdf As DAO.QueryDef
   Dim i As Long
   Dim ArraySize As Long
   
   ArraySize = UBound(ParamValues)
   ReDim QdfParamDefs(ArraySize, 1)
   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = i
      QdfParamDefs(i, 1) = ParamValues(i)
   Next
   
On Error GoTo HandleErr
   
   Set qdf = Me.CurrentDb.CreateQueryDef("", SqlText)
   ExecuteParamSQL = ExecuteQueryDef(qdf, QdfParamDefs)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gespeicherte Abfrage (optional mit Parametern) ausführen
' </summary>
' <param name="QueryName">Name der gespeicherten Abfrage (QueryDef)</param>
' <param name="QueryParamDefs">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDefByName(ByVal QueryName As String, _
                             Optional ByVal QueryParamDefs As Variant) As Long
   
   Dim qdf As DAO.QueryDef
 
On Error GoTo HandleErr
 
   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   ExecuteQueryDefByName = ExecuteQueryDef(qdf, QueryParamDefs)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "ExecuteParamQdfByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gespeicherte Abfrage (optional mit Parametern) ausführen
' </summary>
' <param name="Qdf">QueryDef-Referenz</param>
' <param name="QueryParamDefs">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExecuteQueryDef(ByVal qdf As DAO.QueryDef, _
                       Optional ByVal QdfParamDefs As Variant) As Long
 
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long
   
   If Not IsMissing(QdfParamDefs) Then
      FillQueryDefParameters qdf, QdfParamDefs
   End If
   
   qdf.Execute dbFailOnError
   
   ExecuteQueryDef = qdf.RecordsAffected

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset öffnen
' </summary>
' <param name="Source">SQL-Anweisung oder Tabllen- bzw. Abfragename</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenDynaset, _
                     Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                     Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset

    If (RecordsetOptions And dbSeeChanges) = 0 Then
        RecordsetOptions = RecordsetOptions + dbSeeChanges
    End If
    Set OpenRecordset = Me.CurrentDb.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetParamSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Parameter-SQL-Anweisung öffnen (verwendet temporäres QueryDef)
' </summary>
' <param name="ParamSqlText">SQL-Anweisung mit Parameter (Parameters-Deklaration)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <param name="ParamValues">Parmeterwerte in passender Reihenfolge</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetParamSQL(ByVal ParamSqlText As String, _
                     ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                     ByVal Options As DAO.RecordsetOptionEnum, _
                     ByVal LockEdit As DAO.LockTypeEnum, _
                     ParamArray ParamValues() As Variant) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim i As Long
   Dim QdfParamDefs() As Variant
   Dim ArraySize As Long
   
   ArraySize = UBound(ParamValues)
   ReDim QdfParamDefs(ArraySize, 1)

   For i = 0 To ArraySize
      QdfParamDefs(i, 0) = i
      QdfParamDefs(i, 1) = ParamValues(i)
   Next
   
On Error GoTo HandleErr
   
   Set qdf = Me.CurrentDb.CreateQueryDef("", ParamSqlText)
   Set OpenRecordsetParamSQL = OpenRecordsetQueryDef(qdf, QdfParamDefs, RecordsetType, Options, LockEdit)
   
On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenRecordsetParamSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDefByName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset von gespeicherter Abfrage (optional mit Parametern) öffnen
' </summary>
' <param name="QueryName">Name der gespeicherten Abfrage (QueryDef)</param>
' <param name="QueryParamDefs">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDefByName(ByVal QueryName As String, _
                             Optional ByVal QueryParamDefs As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal Options As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset

   Dim qdf As DAO.QueryDef
   Dim i As Long
 
On Error GoTo HandleErr
 
   Set qdf = Me.CurrentDb.QueryDefs(QueryName)
   Set OpenRecordsetQueryDefByName = OpenRecordsetQueryDef(qdf, QueryParamDefs, RecordsetType, Options, LockEdit)

On Error GoTo 0

   qdf.Close
   Set qdf = Nothing
   
ExitHere:
   Exit Function
   
HandleErr:
   If Not (qdf Is Nothing) Then
      qdf.Close
      Set qdf = Nothing
   End If
   Err.Raise Err.Number, "OpenQueryDefRecordsetByName:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
 
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetQueryDef
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Recordset von gespeicherter Abfrage (optional mit Parametern) öffnen
' </summary>
' <param name="Qdf">QueryDef-Referenz</param>
' <param name="QueryParamDefs">2-dimensionales Parameter-Array (Array(n,1) ... x(n,0) = Parametername, x(n,1) = Parameterwert)</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetQueryDef(ByVal qdf As DAO.QueryDef, _
                             Optional ByVal QdfParamDefs As Variant, _
                             Optional ByVal RecordsetType As DAO.RecordsetTypeEnum, _
                             Optional ByVal Options As DAO.RecordsetOptionEnum, _
                             Optional ByVal LockEdit As DAO.LockTypeEnum) As DAO.Recordset
 
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long
   
   If Not IsMissing(QdfParamDefs) Then
      FillQueryDefParameters qdf, QdfParamDefs
   End If

   Set OpenRecordsetQueryDef = qdf.OpenRecordset(RecordsetType, Options, LockEdit)

End Function

Private Sub FillQueryDefParameters(ByVal qdf As DAO.QueryDef, ByVal QdfParamDefs As Variant)
   
   Dim ParamDefsArraySize As Long
   Dim ParamIndex As Variant
   Dim i As Long

   ParamDefsArraySize = UBound(QdfParamDefs, 1)
   
   With qdf
      For i = 0 To ParamDefsArraySize
         ParamIndex = QdfParamDefs(i, 0)
         If Len(ParamIndex) > 0 Then
            .Parameters(ParamIndex).Value = QdfParamDefs(i, 1)
         End If
      Next i
   End With
   
End Sub

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer vollständigen SQL-Anweisung
' </summary>
' <param name="SqlText">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <param name="ValueIfNull">Rückgabewert bei fehlendem Datensatz oder NULL-Feldwert (Standard: Null</param>
' <returns>Variant</returns>
' <remarks>Die SQL-Anweisung ist im Jet-SQL-Dialekt zu schreiben.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQL(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

    Dim rst As DAO.Recordset

On Error GoTo HandleErr

    Set rst = Me.OpenRecordset(SqlText, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
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

    strSQL = "SELECT " & Expr & " FROM (" & Domain & ")"
    If Not (VarType(Criteria) = vbError) Then
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
    End If
    Lookup = LookupSQL(strSQL, , ValueIfNull)

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
' Function: DMin
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
' Function: DSum
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
' <remarks>
' Funktioniert erst ab Jet 4.0 (Access 2000) und ist vom DBMS abhängig. Bei MySQL-Tabellen wird z. B. 0 zurückgegeben.
' (Die <see cref=AdodbHandler#InsertIdentityReturn>ADODB-Variante</see> dieser Methode ist bezüglich Einsatz für aktive DBMS besser geeignet.)
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSQL As String) As Variant

    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    Dim lngRecordsAffected As Long

On Error GoTo HandleErr

    Set db = Me.CurrentDb
    db.Execute InsertSQL
    lngRecordsAffected = db.RecordsAffected
    
    If lngRecordsAffected = 0 Then
        InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
        Exit Function
    End If

    Set rst = db.OpenRecordset("SELECT @@Identity") 'zuletzt zugewiesener Autowert ... funktioniert nicht mit MySQL-Tabellen!
    With rst
        If .EOF Then
            InsertIdentityReturn = Null
        Else
            InsertIdentityReturn = .Fields(0)
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

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function
