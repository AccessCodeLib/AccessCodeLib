Attribute VB_Name = "TEST_OdbcHandlerHook"
'---------------------------------------------------------------------------------------
' Module: TEST_OdbcHandler (Josef Pötzl, 2010-03-27)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Test-Prozeduren für Hook-Konzept von OdbcHandler
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/odbc/TEST_OdbcHandlerHook.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_test/data/odbc/TEST_OdbcHandlerHook_ModifiyIdentitySelectString.cls</use>
'  <use>data/odbc/OdbcHandler.cls</use>
'  <use>data/dao/TempDbHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Sub ChangeIdentitySelectString()

   Dim tempDb As TempDbHandler
   Dim odbcHdl As OdbcHandler
   Dim odbcHook As TEST_OdbcHandlerHook_ModifiyIdentitySelectString
   
   Dim newID As Long
   
   'Temp-Db mit Tabelle anlegen
   Set tempDb = New TempDbHandler
   tempDb.CreateTable "test", "create table test (id counter(1,1), Z int)"

   'Instanzen erzeugen
   Set odbcHdl = New OdbcHandler
   Set odbcHook = New TEST_OdbcHandlerHook_ModifiyIdentitySelectString
   
   '"Hook" aktivieren
   Set odbcHook.OdbcHandlerReference = odbcHdl
   odbcHdl.HooksEnabled = True
   
   'Db-Instanzen übergeben (solange kein ODBC-String benötigt wird, funktioniert das auch mit Jet-Datenbanken)
   Set odbcHdl.CurrentDb = tempDb.CurrentDatabase
   Set odbcHdl.CurrentDbBE = tempDb.CurrentDatabase


   'Test:
   newID = odbcHdl.InsertIdentityReturn("insert into test (Z) values (" & Str(Timer) & ")")
   Debug.Print "ID:", newID


   'Temp-Db löschen
   tempDb.DeleteCurrentDatabase

End Sub
