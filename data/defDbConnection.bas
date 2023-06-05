Attribute VB_Name = "defDbConnection"
'---------------------------------------------------------------------------------------
' Package: data.defDbConnection
'---------------------------------------------------------------------------------------
'
' Set of global types, enums etc. for the DbConnection classes
'
'---------------------------------------------------------------------------------------

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
' Group: Types
'

'---------------------------------------------------------------------------------------
' Type: DbmsConnectionStrings
'---------------------------------------------------------------------------------------
'
' Type for the simultaneous transfer of the connection strings
'
'     OledbConnectionString   - OLEDB connectionstring for ADODB connection
'     OdbcConnectionString    - ODBC connectionstring for DAO connection
'     DatabaseFile            - Database file name incl. path if a file backend is used
'
'---------------------------------------------------------------------------------------
Public Type DbmsConnectionStrings
   OledbConnectionString As String
   OdbcConnectionString As String
   DatabaseFile As String
End Type


'---------------------------------------------------------------------------------------
' Group: Enums
'

'---------------------------------------------------------------------------------------
' Enum: DbmsConnectionModes
'---------------------------------------------------------------------------------------
'
'     DMBS_DSNless   - (1)    without DSN
'     DMBS_DSN       - (2)    with DSN
'     DBMS_USERDEF   - (128)  User-defined connection strings
'
'---------------------------------------------------------------------------------------
Public Enum DbmsConnectionModes
   DMBS_DSNless = 1   'ohne DSN
   DMBS_DSN = 2       'mit DSN
   DBMS_USERDEF = 128 'benutzerdefinierte Connectionstrings
End Enum
