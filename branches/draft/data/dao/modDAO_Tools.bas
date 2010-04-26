Attribute VB_Name = "modDAO_Tools"
Attribute VB_Description = "Hilfsfunktionen f�r den Umgang mit DAO"
'---------------------------------------------------------------------------------------
' Module: modDAO_Tools
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsfunktionen f�r den Umgang mit DAO
' </summary>
' <remarks>
' </remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/modDAO_Tools.bas</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Function: TableDefExists (Josef P�tzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Pr�ft ob eine Tabelle (TableDef) vorhanden ist
' </summary>
' <param name="sTableDefName">Name der Tabelle</param>
' <param name="dbs">DAO.Database-Referenz (falls keine angegeben wurde, wird CodeDb verwendet)</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TableDefExists(ByVal sTableDefName As String, Optional ByRef dbs As DAO.Database = Nothing) As Boolean
'Schneller w�re der Zugriff auf MSysObject (select .. from MSysObject where Name = 'Tabellenname' AND Type IN (1, 4, 6)
'Eine weitere Alternative w�re das Auswerten �ber cnn.OpenSchema(adSchemaTables, ...) ... dann werden allerdings keine verkn�pften Tabellen gepr�ft
   
   Dim tdf As DAO.TableDef
   
On Error GoTo HandleErr

   If dbs Is Nothing Then
      Set dbs = CodeDb
   End If
   
   dbs.TableDefs.Refresh
   For Each tdf In dbs.TableDefs
      If tdf.Name = sTableDefName Then
         TableDefExists = True
         Exit For
      End If
   Next

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "TableDefExists", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function
