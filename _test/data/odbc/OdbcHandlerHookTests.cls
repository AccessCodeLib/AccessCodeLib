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
'/**
' <summary>
' Test-Klasse für Hook-Konzept von OdbcHandler
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/odbc/OdbcHandlerHookTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/odbc/OdbcHandler.cls</use>
'  <use>data/dao/TempDbHandler.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1
Private m_TempDbHandler As TempDbHandler


'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

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

   Const expected As Long = 12345
   Dim actual As Long

   actual = m_OdbcHandler.InsertIdentityReturn("insert into test (Z) values (" & Str(Timer) & ")")
   
   Assert.AreEqual expected, actual

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
   m_TempDbHandler.CreateTable "test", "create table test (id counter(1,1), Z int)"
End Sub

Private Sub InitOdbcHandler()
   Set m_OdbcHandler = New OdbcHandler
   m_OdbcHandler.HooksEnabled = True
   
   'Db-Instanzen übergeben (solange kein ODBC-String benötigt wird, funktioniert das auch mit Jet-Datenbanken)
   Set m_OdbcHandler.CurrentDb = m_TempDbHandler.CurrentDatabase
   Set m_OdbcHandler.CurrentDbBE = m_TempDbHandler.CurrentDatabase
End Sub
