Attribute VB_Name = "DaoHandler_Beispiele"
'---------------------------------------------------------------------------------------
' Beispiele: DaoHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Beispiele zu DaoHandler (DAO-Zugriffsmethoden)
' </summary>
' <remarks></remarks>
'\ingroup data_dao_examples
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoHandler_Beispiele.bas</file>
'  <use>data/dao/DaoHandler.cls</use>
'  <use>data/dao/DaoTools.bas</use>
'  <execute>DaoHandler_Beispiele_InitTestTablesAndQueries()</execute>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit


'Informationen
'-------------

' DaoHandler kann ohne Intanzierung verwendet werden
' ... selbstinstanzierender Einsatz ist möglich, da in der DaoHandler-Klasse das Attribut VB_PredeclaredId = True gesetzt ist.


'Allg. Konstanten
Private Const TestTableName As String = "DaoHandlerTestTab"
Private Const TestParamQueryDefName As String = "DaoHandlerTestParamQueryDef"


'Beispiele
'---------

Private Sub Test_DaoHandler_CurrentDb()
'Diese Beispiel nutzt die Standardinstanz (durch VB_PredeclaredId = True erzeugt) von DaoHandler

   Debug.Print "CurrentDb: " & DaoHandler.CurrentDb.Name
   ' Sobald  VB_PredeclaredId = True wirkt und keine spezielle Database-Instanz übergeben worden ist,
   ' wird bei fehlender DB-Instanz Application.CurrentDb verwendet.

End Sub

Private Sub Test_DaoHandler_CurrentDb_auf_temporaere_Db_einstellen()

   Dim TempDb As DAO.Database

   ' Datenbankdatei (Temporäre Datenbank) erstellen
   Set TempDb = CreateTempDb()

   'Referenz übergeben
   Set DaoHandler.CurrentDb = TempDb
   Debug.Print "Temp-DB:"; DaoHandler.CurrentDb.Name
   
   DaoHandler.Dispose ' bei Bedarf alle Referenzen entfernen
   
   ' temporär erstellte Datei wieder löschen
   Dim TempDbPath As String
   TempDbPath = TempDb.Name
   TempDb.Close
   Kill TempDbPath

End Sub

Private Sub Test_extra_DaoHandlerInstanz_verwenden()

   Dim DaoHdl As DaoHandler

   Dim i As Long
   
   'DaoHandler-Instanz erzeugen
   Set DaoHdl = New DaoHandler
   
   ' Datenbankdatei erstellen und and TempDbHandler-Instanz übergeben
   Set DaoHdl.CurrentDb = CreateTempDb()
   
   Debug.Print "Temp-DB:"; DaoHdl.CurrentDb.Name
   
   ' Test-Tabelle erzeugen
   DaoHdl.Execute "create table tabTest (id AUTOINCREMENT primary key, Z int, T varchar(5))"
   
   'DS mit SQL-Anweisung anfügen
   DaoHdl.Execute "insert into tabTest (T) VALUES ('Abc')"
   
   'Anzahl DS abfragen (muss 1 sein)
   Debug.Print "Anzahl DS in tabTest:"; DaoHdl.Count("*", "tabTest")

   Dim rst As DAO.Recordset
   Set rst = DaoHdl.OpenRecordset("select * from tabTest")
   With rst
      Do While Not .EOF
         Debug.Print rst.Fields(0), rst.Fields(1), rst.Fields(2)
         .MoveNext
      Loop
      .Close
   End With
   Set rst = Nothing

   'Aufräumen und DB-Dateien löschen
   Dim TempDbPath As String
   TempDbPath = DaoHdl.CurrentDb.Name
   DaoHdl.CurrentDb.Close
   Kill TempDbPath
   
End Sub

Private Sub Test_mehrere_DaoHandler_verwenden()

   Const AnzahlInstanzen As Long = 3

   Dim DaoHdl(1 To AnzahlInstanzen) As DaoHandler

   Dim i As Long, j As Long
   Dim TempDbPath As String
   
   For i = 1 To AnzahlInstanzen
      
      'DaoHandler-Instanz erzeugen
      Set DaoHdl(i) = New DaoHandler
      
      ' Datenbankdatei erstellen und and TempDbHandler-Instanz übergeben
      Set DaoHdl(i).CurrentDb = CreateTempDb(i)
      
      ' ein paar Test-Tabellen erzeugen (hat Auswirkung auf die unten folgende Anweisung  DaoHdl(i).Count("*", "msysobjects"))
      For j = 1 To i
         DaoHdl(i).Execute "create table Test_" & j & " (id int primary key, T varchar(5))"
      Next
      
   Next

   'Unterschiedliche DaoHandler verwenden
   For i = 1 To AnzahlInstanzen
      Debug.Print "Temp-DB: "; DaoHdl(i).CurrentDb.Name,
      Debug.Print "Anzahl DS in MSysObjects: "; DaoHdl(i).Count("*", "MSysObjects")
   Next
   
   'Aufräumen und DB-Dateien löschen
   For i = 1 To AnzahlInstanzen
      TempDbPath = DaoHdl(i).CurrentDb.Name
      DaoHdl(i).CurrentDb.Close
      DaoHdl(i).Dispose
      Kill TempDbPath
   Next

End Sub

Private Sub Insert_mit_ID_Rueckgabe()
   
   Dim NewId As Long

   NewId = DaoHandler.InsertIdentityReturn("insert into " & TestTableName & " ( T, Z) Values ('abc', 123)")
   Debug.Print "New ID: "; NewId

End Sub


Private Sub Recordset_oeffnen()
   
   Dim rst As DAO.Recordset

   Set rst = DaoHandler.OpenRecordset("select * from " & TestTableName)
   Debug.Print "erstes Feld: "; rst.Fields(0)
   rst.Close

End Sub

Private Sub Recordset_aus_ParamAbfrage_oeffnen()
   
   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset

   'Standard-DAO
   Set qdf = CurrentDb.QueryDefs(TestParamQueryDefName)
   qdf.Parameters(0).Value = 100
   qdf.Parameters(1).Value = 23
   Set rst = qdf.OpenRecordset(dbOpenDynaset)
   Debug.Print rst.Fields("Z")
   rst.Close

   'Verkürzt durch DaoHandler:
   Set qdf = DaoHandler.ParamQueryDefByName(TestParamQueryDefName, 100, 23)
   Set rst = qdf.OpenRecordset(dbOpenDynaset)
   Debug.Print rst.Fields("Z")
   rst.Close

   'Verkürzt durch DaoHandler:
   Set rst = DaoHandler.ParamQueryDefByName(TestParamQueryDefName, 100, 23).OpenRecordset(dbOpenDynaset)
   Debug.Print rst.Fields("Z")
   rst.Close

End Sub

Private Sub Recordset_aus_TempParamAbfrage_oeffnen()
   
   Dim SqlString As String
   Dim qdf As DAO.QueryDef
   Dim rst As DAO.Recordset

   SqlString = "Parameters P1 long, P2 long; Select * from " & TestTableName & " where Z = [P1] + [P2]"

   'Standard-DAO
   Set qdf = CurrentDb.CreateQueryDef("")
   qdf.SQL = SqlString
   qdf.Parameters(0).Value = 100
   qdf.Parameters(1).Value = 23
   Set rst = qdf.OpenRecordset(dbOpenDynaset)
   Debug.Print rst.Fields("Z")
   rst.Close

   'Verkürzt durch DaoHandler:
   Set qdf = DaoHandler.ParamQueryDefSql(SqlString, 100, 23)
   Set rst = qdf.OpenRecordset(dbOpenDynaset)
   Debug.Print rst.Fields("Z")
   rst.Close

   'Verkürzt durch DaoHandler:
   Set rst = DaoHandler.ParamQueryDefSql(SqlString, 100, 23).OpenRecordset()
   Debug.Print rst.Fields("Z")
   rst.Close

End Sub


Private Sub Test_DLookupErsatzfunktionen_verwenden()
'Diese Beispiel nutzt die Standardinstanz von DaoHandler

   Debug.Print "Lookup: ID von Name='MSysObjects': "; DaoHandler.Lookup("id", "msysObjects", "Name='MSysObjects'")
   Debug.Print "LookupSQL: ID von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id from MSysObjects where Name='MSysObjects'")
   
   'Mit Nullwert-Ersatz
   Debug.Print
   Debug.Print "Falls Null soll 'N/A' gezeigt werden:"
   Debug.Print "Lookup: ID von Name='MSysObjects': "; DaoHandler.Lookup("id", "msysObjects", "Name='DiesenEintragGibtEsNicht'", "N/A")
   Debug.Print "LookupSQL: ID von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id from MSysObjects where Name='DiesenEintragGibtEsNicht'", , "N/A")

   'Mit Index-Option (Vorteilhaft, wenn man fertige Abfragen o. ä. verwenden will und im Select-Teil mehrere Felder zur Verfügung stehen
   Debug.Print
   Debug.Print "Bestimmtes Feld aus Select-Teil ansprechen:"
   '1. Feld aus Select-Anweisung
   Debug.Print "LookupSQL: ID von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id, Name from MSysObjects", 0)
   '2. Feld aus Select-Anweisung
   Debug.Print "LookupSQL: Name von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id, Name from MSysObjects", 1)

End Sub

Private Sub Test_ParameterAbfragen_verwenden()
'Diese Beispiel nutzt die Standardinstanz von DaoHandler für temporäre Datenbank

   Dim i As Long
   
   'DaoHandler-Instanz erzeugen
   
   ' Datenbankdatei erstellen und and DbHandler-Instanz übergeben
   Set DaoHandler.CurrentDb = CreateTempDb()
   
   ' Test-Tabelle erzeugen
   DaoHandler.Execute "create table tabTest (id AUTOINCREMENT primary key, Z int, T varchar(5))"
   DaoHandler.Execute "insert into tabTest (T) VALUES ('Abc')"
   
   'QueryDef erstellen und verwenden
   Dim qdf As DAO.QueryDef
   Const QueryDefName As String = "qTest_Insert"
   Const QueryDefSQL As String = "Parameters [NewZ] long, [NewT] varchar(255); insert into tabTest (Z, T) VALUES ([NewZ], [NewT])"
   Set qdf = DaoHandler.CurrentDb.CreateQueryDef(QueryDefName, QueryDefSQL)

   'Einfügen mit Array-Variante
   ' a) 2-Dim Array
   For i = 1 To 5
      DaoHandler.ExecuteQueryDefByName QueryDefName, DaoHandler.GetNamedParamDefArray("NewZ", i, "NewT", "abc")
   Next
   ' b) Werte in 1-Dim-Array (Reihenfolge beachten!)"
   For i = 1 To 5
      DaoHandler.ExecuteQueryDefByName QueryDefName, Array(i, "efg")
   Next

   'Einfügen mit ParamArray-Variante
   For i = 1 To 5
      DaoHandler.ExecuteQueryDefByName QueryDefName, i, "xyz"
   Next
   
   'Anzahl DS abfragen (muss 16 (1+5+5+5) sein)
   Debug.Print "Anzahl DS in tabTest:"; DaoHandler.Count("*", "tabTest")

   Dim rst As DAO.Recordset
   Set rst = DaoHandler.OpenRecordset("select * from tabTest")
   With rst
      Do While Not .EOF
         Debug.Print rst.Fields(0), rst.Fields(1), rst.Fields(2)
         .MoveNext
      Loop
      .Close
   End With
   Set rst = Nothing
   

   'Aufräumen und DB-Dateien löschen
   Dim TempDbPath As String
   TempDbPath = DaoHandler.CurrentDb.Name
   DaoHandler.CurrentDb.Close
   DaoHandler.Dispose 'Damit Verweis auf DB entfernt wird
   Kill TempDbPath
   
End Sub


Private Sub Test_TemporäreParameterAbfrage_verwenden()
'Diese Beispiel nutzt die Standardinstanz von DaoHandler für temporäre Datenbank

   Dim i As Long
   
   'DaoHandler-Instanz erzeugen
   
   ' Datenbankdatei erstellen und and DbHandler-Instanz übergeben
   Set DaoHandler.CurrentDb = CreateTempDb()
   
   ' Test-Tabelle erzeugen
   DaoHandler.Execute "create table tabTest (id AUTOINCREMENT primary key, Z int, T varchar(5))"
   DaoHandler.Execute "insert into tabTest (T) VALUES ('Abc')"
   
   'Temoräer QueryDef erstellen und verwenden
   Dim qdf As DAO.QueryDef
   Const QueryDefSQL As String = "Parameters [NewZ] long, [NewT] varchar(255); insert into tabTest (Z, T) VALUES ([NewZ], [NewT])"
   Set qdf = DaoHandler.ParamQueryDefSql(QueryDefSQL)

   'Einfügen mit Array-Variante
   ' a) 2-Dim Array
   For i = 1 To 5
      DaoHandler.ExecuteQueryDef qdf, DaoHandler.GetNamedParamDefArray("NewZ", i, "NewT", "abc")
   Next
   ' b) Werte in 1-Dim-Array (Reihenfolge beachten!)"
   For i = 1 To 5
      DaoHandler.ExecuteQueryDef qdf, Array(i, "efg")
   Next

   'Einfügen mit ParamArray-Variante
   For i = 1 To 5
      DaoHandler.ExecuteQueryDef qdf, i, "xyz"
   Next
   qdf.Close
   
   'Anzahl DS abfragen (muss 16 (1+5+5+5) sein)
   Debug.Print "Anzahl DS in tabTest:"; DaoHandler.Count("*", "tabTest")

   Dim rst As DAO.Recordset
   Set rst = DaoHandler.OpenRecordset("select * from tabTest")
   With rst
      Do While Not .EOF
         Debug.Print rst.Fields(0), rst.Fields(1), rst.Fields(2)
         .MoveNext
      Loop
      .Close
   End With
   Set rst = Nothing
   

   'Aufräumen und DB-Dateien löschen
   Dim TempDbPath As String
   TempDbPath = DaoHandler.CurrentDb.Name
   DaoHandler.CurrentDb.Close
   DaoHandler.Dispose 'Damit Verweis auf DB entfernt wird
   Kill TempDbPath
   
End Sub




'Hilfsfunktionen für Beispiel-Code
'---------------------------------

Private Function CreateTempDb(Optional ByVal FileNameNr As Long) As DAO.Database
'Temoräre Datenbank erstellen

   Dim TempDbPath As String
   Dim TempDb As DAO.Database
   Dim FileNameSuffix  As String

   FileNameSuffix = CStr(Fix(Timer))
   If FileNameNr <> 0 Then
      FileNameSuffix = FileNameSuffix & "_" & FileNameNr
   End If
   
   TempDbPath = CurrentProject.Path & "\TempDb_" & FileNameSuffix & Mid(CurrentProject.Name, InStrRev(CurrentProject.Name, "."))
   Set CreateTempDb = DBEngine.CreateDatabase(TempDbPath, dbLangGeneral)
   
End Function

Public Function DaoHandler_Beispiele_InitTestTablesAndQueries()
'Abfragen und Tabellen für dieses Beispiel-Modul erzeugen
   If Not DaoTools.TableDefExists(TestTableName) Then
      DaoHandler.Execute "create table " & TestTableName & " (id AUTOINCREMENT Primary Key, T varchar(255), Z int)"
   End If
   If Not DaoTools.QueryDefExists(TestParamQueryDefName) Then
      DaoHandler.CurrentDb.CreateQueryDef TestParamQueryDefName, "PARAMETERS P1 Long, P2 Long; SELECT id, T, Z From " & TestTableName & " WHERE Z=([P1]+[P2])"
   End If
   Application.RefreshDatabaseWindow
End Function
