VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionHandlerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für DaoTools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/DbConnectionHandlerTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionHandler.cls</use>
'  <use>data/dao/TempDbHandler.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit


Private WithEvents m_DbConnectionHandler As DbConnectionHandler
Attribute m_DbConnectionHandler.VB_VarHelpID = -1
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
   InitDbConnectionHandler
End Sub

Public Sub Teardown()
   m_DbConnectionHandler.Dispose
   Set m_DbConnectionHandler = Nothing
   Set m_TempDbHandler = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub ADODB_Jet_CheckDataSource()

   Dim expected As String
   Dim actual As String
   
   expected = m_TempDbHandler.CurrentDatabase.Name
   actual = m_DbConnectionHandler.ADODB.CurrentConnection.Properties("Data Source")
   
   Assert.AreEqual expected, actual
   
End Sub

Public Sub CurrentConnection_Jet_CheckDataSource()

   Dim expected As String
   Dim actual As String
   
   expected = m_TempDbHandler.CurrentDatabase.Name
   actual = m_DbConnectionHandler.CurrentConnection.Properties("Data Source")
   
   Assert.AreEqual expected, actual
   
End Sub

Public Sub OledbConnectionString_Jet_CheckDataSource()

   Dim expected As String
   Dim actual As String
   
   m_DbConnectionHandler.ADODB.CurrentConnection.Close
   Set m_DbConnectionHandler.ADODB.CurrentConnection = Nothing
   Debug.Assert Len(m_DbConnectionHandler.OledbConnectionString) = 0
   Debug.Assert m_DbConnectionHandler.ADODB.CurrentConnection Is Nothing
   
   m_DbConnectionHandler.OledbConnectionString = AdodbConnectionString
   
   expected = m_TempDbHandler.CurrentDatabase.Name
   actual = m_DbConnectionHandler.CurrentConnection.Properties("Data Source")
   
   Assert.AreEqual expected, actual
   
End Sub

Public Sub DAO_Jet_EqualToCurrentDb()
   Dim actual As DAO.Database
   Set actual = m_DbConnectionHandler.DAO.CurrentDb
   Assert.That actual.Name, Iz.EqualTo(CurrentDb.Name)
End Sub

Public Sub CurrentDb_Jet_EqualToCurrentDb()
   Dim actual As DAO.Database
   Set actual = m_DbConnectionHandler.CurrentDb
   Assert.That actual.Name, Iz.EqualTo(CurrentDb.Name)
End Sub

Public Sub ODBC_Jet_CheckDatabase()
   Dim actual As DAO.Database
   Set actual = m_DbConnectionHandler.ODBC.CurrentDbBE
   Assert.That actual.Name, Iz.EqualTo(m_TempDbHandler.CurrentDatabase.Name)
End Sub

Public Sub OdbcConnectionString_Jet_ReturnsNullString()

   Const expected As String = vbNullString
   Dim actual As String
   
   actual = m_DbConnectionHandler.OdbcConnectionString
   
   Assert.AreEqual expected, actual
   
End Sub

Public Sub CurrentDbBE_Jet_EqualToTempDb()
   Dim actual As DAO.Database
   Set actual = m_DbConnectionHandler.CurrentDbBE
   Assert.That actual.Name, Iz.EqualTo(m_TempDbHandler.CurrentDatabase.Name)
End Sub


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub InitTempDb()
   Set m_TempDbHandler = New TempDbHandler
   m_TempDbHandler.CreateTable "test", "create table test (id counter(1,1), Z int)"
   Dim DatabaseFile As String
   DatabaseFile = m_TempDbHandler.CurrentDatabase.Name
   m_TempDbHandler.CloseCurrentDatabase
   m_TempDbHandler.OpenDatabase DatabaseFile, False, False
End Sub

Private Sub InitDbConnectionHandler()
   Set m_DbConnectionHandler = New DbConnectionHandler
   With m_DbConnectionHandler
      Set .CurrentDb = CurrentDb
      Set .CurrentDbBE = m_TempDbHandler.CurrentDatabase
      Set .CurrentConnection = GetAdodbConnection
   End With
End Sub

Private Function GetAdodbConnection() As ADODB.Connection
   Dim cnn As ADODB.Connection
   Set cnn = New ADODB.Connection
   cnn.Open AdodbConnectionString
   Set GetAdodbConnection = cnn
End Function

Private Property Get AdodbConnectionString() As String
   AdodbConnectionString = "Provider=Microsoft.JET.OLEDB.4.0;User ID=Admin;Data Source=" & m_TempDbHandler.CurrentDatabase.Name & ";Mode=Share Deny None;"
End Property
