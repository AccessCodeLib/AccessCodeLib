Attribute VB_Name = "defDbConnection"
'---------------------------------------------------------------------------------------
' Modul: defDbConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sammlung von globalen Typen, Enums usw. für die DbConnection-Klassen
' </summary>
' <remarks></remarks>
'\ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/defDbConnection.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
'
' Type + Enums
'

'---------------------------------------------------------------------------------------
' Type: DbmsConnectionStrings
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Type für die gleichzeitige Übergabe der Connectionsstrings
' </summary>
' <list type="table">
'   <item><term>OledbConnectionString</term><description>OLEDB-Connectionstring für ADODB-Verbindung</description></item>
'   <item><term>OdbcConnectionString</term><description>ODBC-Connectionstring für DAO-Verbindung</description></item>
'   <item><term>DatabaseFile</term><description>Datenbankdateiname inkl. Pfad falls ein File-Backend eingesetzt wird</description></item>
' </list>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Type DbmsConnectionStrings
   OledbConnectionString As String
   OdbcConnectionString As String
   DatabaseFile As String
End Type


'---------------------------------------------------------------------------------------
' Enum: DbmsConnectionModes
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DbmsConnectionModes
' </summary>
' <list type="table">
'   <item><term>DMBS_DSNless (1)</term><description>ohne DSN</description></item>
'   <item><term>DMBS_DSN (2)</term><description>mit DSN</description></item>
'   <item><term>aDBMS_USERDEF (128)</term><description>benutzerdefinierte Connectionstrings</description></item>
' </list>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Enum DbmsConnectionModes
   DMBS_DSNless = 1   'ohne DSN
   DMBS_DSN = 2       'mit DSN
   DBMS_USERDEF = 128 'benutzerdefinierte Connectionstrings
End Enum
