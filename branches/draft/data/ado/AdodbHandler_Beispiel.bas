Attribute VB_Name = "AdodbHandlerBeispiel"
Option Compare Database
Option Explicit


'
' Beispiel-Code f�r den Einsatz von AdodbHandler
'

'<codelib>
'  <file>/data/ado/AdodbHandler_Beispiel.bas</file>
'  <use>data/ado/AdodbHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------

Private Sub AdodbHandlerInitialisieren()

   Dim adoHdl As AdodbHandler
   Set adoHdl = New AdodbHandler
   
   
   With adoHdl
      'OLEDB-Connectionstring �bergeben
      .ConnectionString = CurrentProject.Connection
   End With

   Dim rst As ADODB.Recordset

   'ADODB-Recordset �ffnen
   Set rst = adoHdl.OpenRecordset("select * from MSysObjects", adOpenKeyset, adLockReadOnly, adUseClient, True)
   
   With rst
      Debug.Print rst.RecordCount
      .Close
   End With


End Sub
