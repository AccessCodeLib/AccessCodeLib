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
' Klassenmodul: DaoHandler (2009-06-15)
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
'  <use>base/modErrorHandler.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
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
' Property: CurrentDb (2009-06-15)
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
' Property: CurrentDb (2009-06-15)
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
' Function: OpenRecordset (2009-06-15)
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
Public Function OpenRecordset(ByVal source As String, _
                     Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenDynaset, _
                     Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                     Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset
    
On Error GoTo HandleErr

    If (RecordsetOptions And dbSeeChanges) = 0 Then
        RecordsetOptions = RecordsetOptions + dbSeeChanges
    End If
    Set OpenRecordset = Me.CurrentDb.OpenRecordset(source, RecordsetType, RecordsetOptions, LockEdit)

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
' Function: Execute (2009-06-15)
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
                  
On Error GoTo HandleErr

    If (Options And dbSeeChanges) = 0 Then
        Options = Options + dbSeeChanges
    End If
    Me.CurrentDb.Execute CommandText, Options
    RecordsAffected = Me.CurrentDb.RecordsAffected

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
' Function: LookupSQL (2009-06-15)
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
Public Function LookupSQL(ByVal sSQL As String, Optional ByVal index As Variant = 0&) As Variant

    Dim rst As DAO.Recordset

On Error GoTo HandleErr

    Set rst = Me.OpenRecordset(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
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

    strSQL = "SELECT " & Expr & " FROM (" & Domain & ")"
    If Not (VarType(Criteria) = vbError) Then
        If Len(Criteria) > 0 Then
            strSQL = strSQL & " WHERE " & Criteria
        End If
    End If
    Lookup = LookupSQL(strSQL)

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
' Function: Count (2009-06-15)
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

On Error GoTo HandleErr

    Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Count", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: Max (2009-06-15)
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

On Error GoTo HandleErr

    Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Max", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
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
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

    Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Min", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
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
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant

On Error GoTo HandleErr

    Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Sum", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
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
