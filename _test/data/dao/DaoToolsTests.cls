VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _test.data.dao.DaoToolsTests
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'
' AccUnit test class for DaoTools
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/dao/DaoToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

'Implements SimplyVBUnit.ITestFixture
'Implements AccUnit_Integration.ITestManagerBridge
'Private TestManager As AccUnit_Integration.TestManager
'Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
'Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
'Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub TableDefExists_UseSystemTableInCodeDb_ReturnsTrue()
   Const Expected As Boolean = True
   Dim Actual As Boolean
   Actual = DaoTools.TableDefExists("MSysObjects")
   Assert.AreEqual Expected, Actual
End Sub

Public Sub TableDefExists_UseMissingNameInCodeDb_ReturnsFalse()
   
   Const Expected As Boolean = False
   Dim NameToTest As String
   Dim Actual As Boolean
   
   NameToTest = GetTableMissingNameInCodeDb
   Actual = DaoTools.TableDefExists(NameToTest)
   Assert.AreEqual Expected, Actual
   
End Sub



'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Function GetTableMissingNameInCodeDb() As String

   Dim TableName As String
   Dim tdf As DAO.TableDef
   Dim NameExists As Boolean
   
   TableName = "T" & CStr(Fix(Timer))
   Do
      NameExists = False
      TableName = TableName & "x"
      For Each tdf In CodeDb.TableDefs
         If tdf.Name = TableName Then
            NameExists = True
            Exit For
         End If
      Next
   Loop While NameExists
   
   GetTableMissingNameInCodeDb = TableName
   
End Function
