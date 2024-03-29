VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OdbcHandlerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Test-Klasse f�r OdbcHandler
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/odbc/OdbcHandlerTests.cls</file>
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

Private Const TestTableName As String = "Test"
Private Const TestTableDDL As String = "CREATE TABLE " & TestTableName & " (id counter(1,1), T varchar(255), Z int)"

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
   DisposeTempDb
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub CurrentDbBE_Jet_EqualToTempDb()
   Dim actual As DAO.Database
   Set actual = m_OdbcHandler.CurrentDbBE
   Assert.That actual.Name, Iz.EqualTo(m_TempDbHandler.CurrentDatabase.Name)
End Sub

Public Sub ConnectionString_Jet_ReturnsNullString()
   
   Dim expected As String
   Dim actual As String
   
   expected = OdbcConnectionString
   actual = m_OdbcHandler.ConnectionString
   Assert.AreEqual expected, actual
   
End Sub

Public Sub RelinkTable_Jet_ReturnsTrue()

   Const expected As Boolean = True
   Dim actual As Boolean

   actual = m_OdbcHandler.RelinkTable(TestTableName)
   
   Assert.AreEqual expected, actual
   
End Sub


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub InitTempDb()
   Set m_TempDbHandler = New TempDbHandler
   m_TempDbHandler.CreateTable TestTableName, TestTableDDL
   LinkTestTable
End Sub

Private Sub InitOdbcHandler()
   Set m_OdbcHandler = New OdbcHandler
   m_OdbcHandler.HooksEnabled = True
   
   'Db-Instanzen �bergeben (solange kein "richtiger" ODBC-String ben�tigt wird, funktioniert das auch mit Jet-Datenbanken)
   Set m_OdbcHandler.CurrentDbBE = m_TempDbHandler.CurrentDatabase
   m_OdbcHandler.ConnectionString = OdbcConnectionString
   Set m_OdbcHandler.CurrentDb = CurrentDb
   
End Sub

Private Sub LinkTestTable()
   
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set db = CurrentDb
   Set tdf = db.CreateTableDef(TestTableName)
   tdf.SourceTableName = TestTableName
   tdf.Connect = OdbcConnectionString
   db.TableDefs.Append tdf
   db.TableDefs.Refresh

End Sub

Private Property Get OdbcConnectionString() As String
   OdbcConnectionString = ";Database=" & m_TempDbHandler.CurrentDatabase.Name
End Property

Private Sub DisposeTempDb()
   CurrentDb.TableDefs.Delete TestTableName
   Set m_TempDbHandler = Nothing
End Sub
