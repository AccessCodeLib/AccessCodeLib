VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdbcHandlerHookTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'
' Test class for hook concept of OdbcHandler
'
' Author:
'     Josef Poetzl
'
' Call:
'     TestSuite.AddByClassName("OdbcHandlerHookTests").Run
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/odbc/OdbcHandlerHookTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/odbc/OdbcHandler.cls</use>
'  <use>data/dao/TempDbHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1
Private m_TempDbHandler As TempDbHandler

Public Sub Setup()
   InitTempDb
   InitOdbcHandler
End Sub

Public Sub Teardown()
   Set m_OdbcHandler = Nothing
   Set m_TempDbHandler = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub InsertIdentityReturn_HookChangeSelectString_Returns12345()

   Const Expected As Long = 12345
   Dim Actual As Long

   Actual = m_OdbcHandler.InsertIdentityReturn("insert into test (N) values (" & Str(Timer) & ")")
   
   Assert.AreEqual Expected, Actual

End Sub

Private Sub m_OdbcHandler_HookInsertIdentityReturn(InsertSQL As String, ActiveDb As DAO.Database, IdentitySelectString As String, ResponseMode As Long, ResponseMessage As Variant)
   
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


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub InitTempDb()
   Set m_TempDbHandler = New TempDbHandler
   m_TempDbHandler.CreateTable "test", "create table test (id counter(1,1), N int)"
End Sub

Private Sub InitOdbcHandler()
   Set m_OdbcHandler = New OdbcHandler
   m_OdbcHandler.HooksEnabled = True
   
   'Db-Instanzen übergeben (solange kein ODBC-String benötigt wird, funktioniert das auch mit Jet-Datenbanken)
   Set m_OdbcHandler.CurrentDb = m_TempDbHandler.CurrentDatabase
   Set m_OdbcHandler.CurrentDbBE = m_TempDbHandler.CurrentDatabase
End Sub
