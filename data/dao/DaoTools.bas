Attribute VB_Name = "DaoTools"
Attribute VB_Description = "Hilfsfunktionen für den Umgang mit DAO"
'---------------------------------------------------------------------------------------
' Module: DaoTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Hilfsfunktionen für den Umgang mit DAO
' </summary>
' <remarks>
' </remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Function: TableDefExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft ob eine Tabelle (TableDef) vorhanden ist
' </summary>
' <param name="sTableDefName">Name der Tabelle</param>
' <param name="dbs">DAO.Database-Referenz (falls keine angegeben wurde, wird CodeDb verwendet)</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TableDefExists(ByVal TableDefName As String, _
                      Optional ByVal dbs As DAO.Database = Nothing) As Boolean
'Man könnte auch die TableDef-Liste durchlaufen.
'Eine weitere Alternative wäre das Auswerten über cnn.OpenSchema(adSchemaTables, ...)
   
   Dim rst As DAO.Recordset

   If dbs Is Nothing Then
      Set dbs = CodeDb
   End If
   
   Set rst = dbs.OpenRecordset("select Name from MSysObjects where Name = '" & Replace(TableDefName, "'", "''") & "' AND Type IN (1, 4, 6)", dbOpenForwardOnly, dbReadOnly)
   TableDefExists = Not rst.EOF
   rst.Close
   
End Function
