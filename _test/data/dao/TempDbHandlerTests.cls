VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempDbHandlerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: TempDbHandlerTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für TempDbHandler
' </summary>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/dao/TempDbHandlerTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/TempDbHandler.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_TempDbHandler As TempDbHandler
Private Const TestTableName As String = "Test"
Private Const TestTableDDL As String = "CREATE TABLE " & TestTableName & " (id counter(1,1), T varchar(255), Z int)"
Private Const InsertSQL As String = "Insert INTO " & TestTableName & " (T, Z) VALUES ('{0}', {1})"
Private Const DeleteRecordsSQL As String = "DELETE FROM " & TestTableName

'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set m_TempDbHandler = New TempDbHandler
End Sub

Public Sub TearDown()
   m_TempDbHandler.Dispose
   Set m_TempDbHandler = Nothing
End Sub


'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub CheckTable_TableMissing_CreateTable()
   m_TempDbHandler.CheckTable TestTableName, TestTableDDL
   Assert.IsTrue TestTableExists
End Sub

Public Sub CreateTable_TableMissing_CreateTable()
   m_TempDbHandler.CreateTable TestTableName, TestTableDDL
   Assert.IsTrue TestTableExists
End Sub

Public Sub CreateTable_TableExists_RaiseError()

   Const expected As Long = 0
   Dim actual As Long
   
   CreateTestTableInTempDB
   
   Assert.Throws vbObjectError + 9001
   m_TempDbHandler.CreateTable TestTableName, TestTableDDL
   
End Sub

Public Sub DropTable_TableExits_TableDeleted()

   CreateTestTableInTempDB
   
   m_TempDbHandler.DropTable TestTableName
   
   Assert.IsFalse TestTableExists
   
End Sub

Public Sub DropLinkedTableDefs_TableMissing_NoError()
   
On Error GoTo HandleErr

   m_TempDbHandler.DropTable TestTableName

ExitHere:
   Exit Sub

HandleErr:
   Select Case Err.Number
   Case 3376 'Table does not exist.
      Assert.Fail "Error 3376 raised"
   Case Else
      Assert.Pass
   End Select
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordset_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal lId As Long, ByVal expected As Variant)
   
   Const SelectDDLTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
   Dim actual As Variant
   Dim rst As DAO.Recordset
   
   CreateTestTableInTempDB
   InsertRecordsInTestTable
   
   Set rst = m_TempDbHandler.OpenRecordset(FormatString(SelectDDLTemplate, lId))
   actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
'AccUnit:Row(0, 2, 2)
'AccUnit:Row(1, 2, "b")
'AccUnit:Row(2, 2, 2)
Public Sub DLookupSQL_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldIndex As Variant, ByVal lId As Long, ByVal expected As Variant)

   Const SelectDDLTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
   Dim actual As Variant
   
   CreateTestTableInTempDB
   InsertRecordsInTestTable
   
   actual = m_TempDbHandler.LookupSQL(FormatString(SelectDDLTemplate, lId), FieldIndex)

   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

Public Sub Execute_DeleteRecordsInTestTable_EmptyTestTable()
   
   Const expected As Long = 0
   Dim actual As Long
   
   CreateTestTableInTempDB
   InsertRecordsInTestTable
   
   m_TempDbHandler.Execute DeleteRecordsSQL, dbFailOnError
   actual = RecordCountInTestTable
   
   ' Assert
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub CreateTestTableInTempDB()
   m_TempDbHandler.CurrentDatabase.Execute TestTableDDL
End Sub

Private Function InsertRecordsInTestTable() As Long
   Dim db As DAO.Database
   Dim i As Long
   
   Set db = m_TempDbHandler.CurrentDatabase
   
   For i = 1 To 5
      db.Execute FormatString(InsertSQL, Chr(i + 96), i)
   Next
   
   If RecordCountInTestTable = 0 Then
      Err.Raise vbObjectError, "InsertRecordsInTestTable", "missing records"
   End If
   
   InsertRecordsInTestTable = 5
   
End Function

Private Property Get RecordCountInTestTable() As Long
   Dim db As DAO.Database
   Dim rst As DAO.Recordset
   Dim lngRecordCount As Long
   
   Set db = m_TempDbHandler.CurrentDatabase
   Set rst = db.OpenRecordset("select count(*) from " & TestTableName)
   lngRecordCount = rst.Fields(0)
   rst.Close
   
   RecordCountInTestTable = lngRecordCount
   
End Property

Private Property Get TestTableExists() As Boolean

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set db = m_TempDbHandler.CurrentDatabase
   db.TableDefs.Refresh
   For Each tdf In db.TableDefs
      If tdf.Name = TestTableName Then
         TestTableExists = True
         Exit Property
      End If
   Next
   
End Property

Private Function FormatString(ByVal sFormat As String, ParamArray Args() As Variant) As String

   Dim strText As String
   Dim i As Long
   strText = sFormat
   For i = 0 To UBound(Args)
      strText = Replace(strText, "{" & i & "}", Args(i))
   Next
   FormatString = strText

End Function
