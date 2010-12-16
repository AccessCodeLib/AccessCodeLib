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
' Klassenmodul: DaoHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO-Zugriffsmethoden
' </summary>
' <remarks></remarks>
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
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'/**
' <summary>
' openRecordset
' </summary>
' <param name="Source">SQL-Anweisung</param>
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
' Function: Execute
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Execute
' </summary>
' <param name="CommandText">SQL-Anweisung</param>
' <param name="RecordsAffected"></param>
' <param name="Options"></param>
' <remarks>Achtung: Execute-Parameter weichen von DAO-Execute ab, da RecordsAffected ergänzt wurde</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Execute(ByVal CommandText As String, _
          Optional ByRef RecordsAffected As Long, _
          Optional ByVal Options As DAO.RecordsetOptionEnum)

    If (Options And dbSeeChanges) = 0 Then
        Options = Options + dbSeeChanges
    End If
    Me.CurrentDb.Execute CommandText, Options
    RecordsAffected = Me.CurrentDb.RecordsAffected
 
End Sub


'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Lookup-Recordset-Ersatzfunktion für die Übergabe einer vollständigen SQL-Anweisung
' </summary>
' <param name="sSQL">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgegeben werden soll</param>
' <returns>Variant</returns>
' <remarks>Die SQL-Anweisung ist im Jet-SQL-Dialekt zu schreiben.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSQL(ByVal sSQL As String, Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

    Dim rst As DAO.Recordset

    Set rst = Me.OpenRecordset(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
    With rst
        If .EOF Then
            LookupSQL = ValueIfNull
        Else
            LookupSQL = Nz(.Fields(Index), ValueIfNull)
        End If
        .Close
    End With
    Set rst = Nothing

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
Public Function Lookup(ByVal expr As String, ByVal Domain As String, _
              Optional ByVal Criteria As Variant, _
              Optional ByVal ValueIfNull As Variant = Null) As Variant

    Dim strSQL As String

    strSQL = "SELECT " & expr & " FROM (" & Domain & ")"
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
Public Function Count(ByVal expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Long
    Count = Nz(Me.Lookup("Count(" & expr & ")", Domain, Criteria), 0)
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
Public Function Max(ByVal expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Max = Me.Lookup("Max(" & expr & ")", Domain, Criteria)
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
Public Function Min(ByVal expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Min = Me.Lookup("Min(" & expr & ")", Domain, Criteria)
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
Public Function Sum(ByVal expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
    Sum = Me.Lookup("Sum(" & expr & ")", Domain, Criteria)
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
    Set db = Nothing

End Function
