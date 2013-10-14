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
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Informationen
'-------------

' DaoHandler kann ohne Intanzierung verwendet werden
' ... selbstinstanzierender Einsatz ist m�glich, da in der DaoHandler-Klasse das Attribut VB_PredeclaredId = True gesetzt ist.


'Beispiele
'---------

Private Sub Test_DaoHandler_CurrentDb()
'Diese Beispiel nutzt die Standardinstanz (durch VB_PredeclaredId = True erzeugt) von DaoHandler

   Debug.Print "CurrentDb: " & DaoHandler.CurrentDb.Name
   ' Sobald  VB_PredeclaredId = True wirkt und keine spezielle Database-Instanz �bergeben worden ist,
   ' wird bei fehlender DB-Instanz Application.CurrentDb verwendet.

End Sub

Private Sub Test_DaoHandler_CurrentDb_auf_temporaere_Db_einstellen()

   Dim TempDbPath As String
   Dim TempDb As DAO.Database
   
   TempDbPath = CurrentProject.Path & "\TempDb_" & Fix(Timer) & Mid(CurrentProject.Name, InStrRev(CurrentProject.Name, "."))
   
   ' Datenbankdatei erstellen
   Set TempDb = DBEngine.CreateDatabase(TempDbPath, dbLangGeneral)

   'Referenz �bergeben
   Set DaoHandler.CurrentDb = TempDb
   Debug.Print "Temp-DB:"; DaoHandler.CurrentDb.Name
   
   DaoHandler.Dispose ' bei Bedarf alle Referenzen entfernen
   
   ' tempor�r erstellte Datei wieder l�schen
   TempDb.Close
   Kill TempDbPath

End Sub

Private Sub Test_extra_DaoHandlerInstanz_verwenden()

   Dim DaoHdl As DaoHandler

   Dim TempDbPath As String
   Dim i As Long
   
   'DaoHandler-Instanz erzeugen
   Set DaoHdl = New DaoHandler
   
   ' Datenbankdatei erstellen und and TempDbHandler-Instanz �bergeben
   TempDbPath = CurrentProject.Path & "\TempDb_" & Fix(Timer) & "_" & Mid(CurrentProject.Name, InStrRev(CurrentProject.Name, "."))
   Set DaoHdl.CurrentDb = DBEngine.CreateDatabase(TempDbPath, dbLangGeneral)
   
   Debug.Print "Temp-DB:"; DaoHdl.CurrentDb.Name
   
   ' Test-Tabelle erzeugen
   DaoHdl.Execute "create table tabTest (id AUTOINCREMENT primary key, Z int, T varchar(5))"
   
   'DS mit SQL-Anweisung anf�gen
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

   'Aufr�umen und DB-Dateien l�schen
   DaoHdl.CurrentDb.Close
   Kill TempDbPath
   
End Sub

Private Sub Test_mehrere_DaoHandler_verwenden()

   Const AnzahlInstanzen As Long = 3

   Dim DaoHdl(1 To AnzahlInstanzen) As DaoHandler

   Dim TempDbPath As String
   Dim i As Long, j As Long
   
   Dim TempString As String
   Dim FileExtension As String
   
   FileExtension = Mid(CurrentProject.Name, InStrRev(CurrentProject.Name, "."))
   TempString = CurrentProject.Path & "\TempDb_" & Fix(Timer) & "_"
   
   For i = 1 To AnzahlInstanzen
      
      'DaoHandler-Instanz erzeugen
      Set DaoHdl(i) = New DaoHandler
      
      ' Datenbankdatei erstellen und and TempDbHandler-Instanz �bergeben
      TempDbPath = TempString & i & FileExtension
      Set DaoHdl(i).CurrentDb = DBEngine.CreateDatabase(TempDbPath, dbLangGeneral)
      
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
   
   'Aufr�umen und DB-Dateien l�schen
   For i = 1 To AnzahlInstanzen
      TempDbPath = DaoHdl(i).CurrentDb.Name
      DaoHdl(i).CurrentDb.Close
      Kill TempDbPath
   Next
   
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

   'Mit Index-Option (Vorteilhaft, wenn man fertige Abfragen o. �. verwenden will und im Select-Teil mehrere Felder zur Verf�gung stehen
   Debug.Print
   Debug.Print "Bestimmtes Feld aus Select-Teil ansprechen:"
   '1. Feld aus Select-Anweisung
   Debug.Print "LookupSQL: ID von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id, Name from MSysObjects", 0)
   '2. Feld aus Select-Anweisung
   Debug.Print "LookupSQL: Name von Name='MSysObjects': "; DaoHandler.LookupSQL("Select id, Name from MSysObjects", 1)

End Sub

Private Sub Test_ParameterAbfragen_verwenden()
'Diese Beispiel nutzt die Standardinstanz von DaoHandler

   Dim TempDbPath As String
   Dim i As Long
   
   'DaoHandler-Instanz erzeugen
   
   ' Datenbankdatei erstellen und and DbHandler-Instanz �bergeben
   TempDbPath = CurrentProject.Path & "\TempDb_" & Fix(Timer) & "_" & Mid(CurrentProject.Name, InStrRev(CurrentProject.Name, "."))
   Set DaoHandler.CurrentDb = DBEngine.CreateDatabase(TempDbPath, dbLangGeneral)
   
   ' Test-Tabelle erzeugen
   DaoHandler.Execute "create table tabTest (id AUTOINCREMENT primary key, Z int, T varchar(5))"
   
   'weitere Verwendungsm�glichkeiten
   Debug.Print DaoHandler.CurrentDb.Name
   
   DaoHandler.Execute "insert into tabTest (T) VALUES ('Abc')"
   
   'QueryDef erstellen und verwenden
   Dim qdf As DAO.QueryDef
   Const QueryDefName As String = "qTest_Insert"
   Set qdf = DaoHandler.CurrentDb.CreateQueryDef(QueryDefName, "Parameters [NewZ] long, [NewT] varchar(255); insert into tabTest (Z, T) VALUES ([NewZ], [NewT])")
   'Einf�gen mit Array-Variante
   For i = 1 To 5
      DaoHandler.ExecuteQueryDefByName QueryDefName, DaoHandler.GetNamedParamDefArray("NewZ", i, "NewT", "abc")
   Next
   'Einf�gen mit ParamArray-Variante
   For i = 1 To 5
      DaoHandler.ExecuteQueryDefByName QueryDefName, i, "xyz"
   Next
   
   'Anzahl DS abfragen (muss 11 (1+5+5) sein)
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
   

   'Aufr�umen und DB-Dateien l�schen
   DaoHandler.CurrentDb.Close
   DaoHandler.Dispose 'Damit Verweis auf DB entfernt wird
   Kill TempDbPath
   
End Sub
