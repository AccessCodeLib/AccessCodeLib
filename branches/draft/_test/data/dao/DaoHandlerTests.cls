VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoHandlerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse f�r DaoHandler
' </summary>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/dao/DaoHandlerTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoHandler.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_DaoHandler As DaoHandler

Private Const TestTableName As String = "DaoHandlerTestTable"
Private Const TestQueryDefNameSelect As String = "DaoHandlerTestQdfSelect"
Private Const TestQueryDefNameInsert As String = "DaoHandlerTestQdfInsert"

Private Const TestTableCreateDDL As String = "CREATE TABLE " & TestTableName & " (id counter(1,1), T varchar(255), Z int)"
Private Const TestTableDropDDL As String = "Drop TABLE " & TestTableName
Private Const SelectSqlTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
Private Const SelectParamSql As String = "select ID, T, Z from " & TestTableName & " where ID=[FilterId]"
Private Const InsertSqlTemplate As String = "Insert INTO " & TestTableName & " (T, Z) VALUES ('{0}', {1})"
Private Const InsertParamSql As String = "Parameters [P1] text(255), [P2] Long; Insert INTO " & TestTableName & " (T, Z) VALUES ([P1], [P2])"
Private Const DeleteRecordsSQL As String = "DELETE FROM " & TestTableName

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
   Set m_DaoHandler = New DaoHandler
   Set m_DaoHandler.CurrentDb = CurrentDb
   CreateTestTable
   InsertRecordsInTestTable
   CreateTestQueryDef
End Sub

Public Sub Teardown()
   m_DaoHandler.Dispose
   Set m_DaoHandler = Nothing
   DropTestQueryDef
   DropTestTable
End Sub


'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub Execute_DeleteStatement_EmptyTable()

   Const expected As Long = 0
   Dim actual As Long
   
   m_DaoHandler.Execute DeleteRecordsSQL, dbFailOnError
   actual = RecordCountInTestTable
   
   ' Assert
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

Public Sub Execute_InsertRecord_NewRecord()

   Const expected As Long = 0
   Dim actual As Long
   Dim SqlText As String
   Dim RecordCountBeforeInsert As Long
   
   RecordCountBeforeInsert = RecordCountInTestTable
   SqlText = FormatString(InsertSqlTemplate, "abc", 123)
   m_DaoHandler.Execute SqlText, dbFailOnError
   
   actual = RecordCountInTestTable
   Assert.That actual, Iz.GreaterThan(RecordCountBeforeInsert)
   
End Sub

Public Sub Execute_InsertRecord_RecordsAffectedIs1()

   Const expected As Long = 1
   Dim actual As Long
   Dim SqlText As String
   
   SqlText = FormatString(InsertSqlTemplate, "abc", 123)
   actual = m_DaoHandler.Execute(SqlText, dbFailOnError)
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

Public Sub ExecuteParamSql_InsertRecord_RecordsAffectedIs1()

   Const expected As Long = 1
   Dim actual As Long
   Dim SqlText As String
   
   actual = m_DaoHandler.ExecuteParamSQL(InsertParamSql, "abc", 123)
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordset_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal expected As Variant)
   
   
   Dim actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordset(FormatString(SelectSqlTemplate, Id))
   actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordsetParamSQL_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal expected As Variant)
   
   Const SelectSql As String = "Parameters [FilterId] long; select ID, T, Z from " & TestTableName & " where ID=[FilterId]"
   Dim actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordsetParamSQL(SelectSql, dbOpenForwardOnly, dbFailOnError, dbOptimistic, Id)
   actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordsetQueryDefByName_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal expected As Variant)
   
   Dim actual As Variant
   Dim rst As DAO.Recordset
   
   Dim QdfParamDefs(0, 1) As Variant
   
   QdfParamDefs(0, 0) = "FilterId"
   QdfParamDefs(0, 1) = Id

   Set rst = m_DaoHandler.OpenRecordsetQueryDefByName(TestQueryDefNameSelect, QdfParamDefs, dbOpenForwardOnly, dbFailOnError, dbOptimistic)
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
Public Sub DLookupSQL_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldIndex As Variant, _
                                                                    ByVal Id As Long, _
                                                                    ByVal expected As Variant)

   Const SelectDDLTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
   Dim actual As Variant
   
   actual = m_DaoHandler.LookupSQL(FormatString(SelectDDLTemplate, Id), FieldIndex)
   Assert.That actual, Iz.EqualTo(expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub Lookup_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldName As Variant, _
                                                                ByVal Id As Long, _
                                                                ByVal expected As Variant)
   Dim actual As Variant
   actual = m_DaoHandler.Lookup(FieldName, TestTableName, FormatString("ID={0}", Id))
   Assert.That actual, Iz.EqualTo(expected)
End Sub

'AccUnit:Row("=", 2, 1)
'AccUnit:Row("=", 0, 0)
'AccUnit:Row("<=", 2 , 2)
Public Sub Count_FilterRecordWithID_ReturnCount(ByVal Expr As String, ByVal Id As Long, ByVal expected As Long)
   Dim actual As Variant
   actual = m_DaoHandler.Count("*", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   Assert.That actual, Iz.EqualTo(expected)
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 2)
'AccUnit:Row(">=", 2 , 5)
Public Sub Max_FilterRecordWithID_ReturnMaxZ(ByVal Expr As String, ByVal Id As Long, ByVal expected As Variant)

   Dim actual As Variant
   actual = m_DaoHandler.Max("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(expected) Then
      Assert.IsNull actual
   Else
      Assert.AreEqual expected, actual
   End If
   
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 1)
'AccUnit:Row(">=", 2 , 2)
Public Sub Min_FilterRecordWithID_ReturnMinZ(ByVal Expr As String, ByVal Id As Long, ByVal expected As Variant)

   Dim actual As Variant
   actual = m_DaoHandler.Min("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(expected) Then
      Assert.IsNull actual
   Else
      Assert.AreEqual expected, actual
   End If
   
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 3)
'AccUnit:Row(">=", 2 , 14)
Public Sub Sum_FilterRecordWithID_ReturnSumZ(ByVal Expr As String, ByVal Id As Long, ByVal expected As Variant)

   Dim actual As Variant
   actual = m_DaoHandler.Sum("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(expected) Then
      Assert.IsNull actual
   Else
      Assert.AreEqual expected, actual
   End If
   
End Sub

Public Sub InsertIdentityReturn_InsertNewRecord_ReturnIdentityFromNewRecord()

   Const NewZ As Long = 999
   Dim actual As Long
   Dim expected As Long
   
   actual = m_DaoHandler.InsertIdentityReturn(FormatString(InsertSqlTemplate, "ABC", NewZ))
   expected = Nz(DLookup("ID", TestTableName, "Z=" & NewZ), 0)
   Assert.AreEqual expected, actual
   
End Sub


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub CreateTestTable()
On Error GoTo HandleErr

   CurrentDb.Execute TestTableCreateDDL

ExitHere:
   Exit Sub

HandleErr:
   Select Case Err.Number
      Case 3010
         DropTestTable
         Resume
      Case Else
         Err.Raise Err.Number, "CreateTestTable:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   End Select
End Sub

Private Sub DropTestTable()
   CurrentDb.Execute TestTableDropDDL
End Sub

Private Function InsertRecordsInTestTable() As Long
   Dim db As DAO.Database
   Dim i As Long
   
   Set db = CurrentDb
   
   For i = 1 To 5
      db.Execute FormatString(InsertSqlTemplate, Chr(i + 96), i)
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
   
   Set db = CurrentDb
   Set rst = db.OpenRecordset("select count(*) from " & TestTableName)
   lngRecordCount = rst.Fields(0)
   rst.Close
   
   RecordCountInTestTable = lngRecordCount
   
End Property

Private Property Get TestTableExists() As Boolean

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set db = CurrentDb
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

Private Sub CreateTestQueryDef()
   
   Dim db As DAO.Database
   
   Set db = CodeDb
   
   db.CreateQueryDef TestQueryDefNameSelect, SelectParamSql
   db.CreateQueryDef TestQueryDefNameInsert, InsertParamSql
   
End Sub

Private Sub DropTestQueryDef()
   
   Dim db As DAO.Database
   
   Set db = CodeDb
   
   db.QueryDefs.Delete TestQueryDefNameSelect
   db.QueryDefs.Delete TestQueryDefNameInsert
   
End Sub