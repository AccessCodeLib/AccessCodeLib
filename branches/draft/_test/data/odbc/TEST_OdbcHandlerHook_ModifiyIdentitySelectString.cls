VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TEST_OdbcHandlerHook_ModifiyIdentitySelectString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module: TEST_OdbcHandler (Josef Pötzl, 2010-03-27)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hook-Test für OdbcHandler: Verändert SQL-Anweisung zum Abfragen des neuen ID-Wertes
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/odbc/TEST_OdbcHandlerHook_ModifiyIdentitySelectString.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/odbc/OdbcHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1

Public Property Set OdbcHandlerReference(ref As OdbcHandler)
   Set m_OdbcHandler = ref
End Property

Private Sub m_OdbcHandler_HookInsertIdentityReturn(InsertSQL As String, ActiveDb As DAO.Database, IdentitySelectString As String, HookResumeMode As Long, HookResumeMessage As Variant)
   
   Debug.Print "<hook>"
   Debug.Print "HookInsertIdentityReturn: "
   Debug.Print , "Insert-Anweisung wird normal ausgeführt"
   Debug.Print , "SQL-Anweisung für Rückgabewert wird von Hook überschrieben"
   Debug.Print "InsertSQL", InsertSQL
   Debug.Print "ActiveDb", ActiveDb.Name
   
   IdentitySelectString = "select Top 1 12345 as ID from MSysObjects T order by T.id"
   Debug.Print "IdentitySelectString", IdentitySelectString
   
   Debug.Print "</hook>"
   
End Sub
