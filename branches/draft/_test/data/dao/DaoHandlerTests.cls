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
' AccUnit-Testklasse für DaoHandler
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
Private Const TestQueryDefNameSelect2 As String = "TestQueryDefNameSelect2"
Private Const TestQueryDefNameInsert As String = "DaoHandlerTestQdfInsert"

Private Const TestTableCreateDDL As String = "CREATE TABLE " & TestTableName & " (id counter(1,1), T varchar(255), Z int)"
Private Const TestTableDropDDL As String = "Drop TABLE " & TestTableName
Private Const SelectSqlTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
Private Const SelectParamSql As String = "select ID, T, Z from " & TestTableName & " where ID=[FilterId]"
Private Const SelectParamSql2 As String = "Parameters [P1] Long, [P2] Long; select ID, T, Z from " & TestTableName & " where Z=[P1]+[P2]"
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
   
   'DropTestQueryDef
   'DropTestTable
   
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

   Const Expected As Long = 0
   Dim Actual As Long
   
   m_DaoHandler.Execute DeleteRecordsSQL, dbFailOnError
   Actual = RecordCountInTestTable
   
   ' Assert
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub Execute_InsertRecord_NewRecord()

   Const Expected As Long = 0
   Dim Actual As Long
   Dim SqlText As String
   Dim RecordCountBeforeInsert As Long
   
   RecordCountBeforeInsert = RecordCountInTestTable
   SqlText = FormatString(InsertSqlTemplate, "abc", 123)
   m_DaoHandler.Execute SqlText, dbFailOnError
   
   Actual = RecordCountInTestTable
   Assert.That Actual, Iz.GreaterThan(RecordCountBeforeInsert)
   
End Sub

Public Sub Execute_InsertRecord_RecordsAffectedIs1()

   Const Expected As Long = 1
   Dim Actual As Long
   Dim SqlText As String
   
   SqlText = FormatString(InsertSqlTemplate, "abc", 123)
   Actual = m_DaoHandler.Execute(SqlText, dbFailOnError)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ExecuteQueryDef_0Param_CheckInsertedCount()
   
   Dim qdf As DAO.QueryDef
   Dim Actual As Long
   Const Expected As Long = 1
   
   Set qdf = m_DaoHandler.CurrentDb.CreateQueryDef("", FormatString(InsertSqlTemplate, "abc", 123))
   
   Actual = m_DaoHandler.ExecuteQueryDef(qdf)
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub ExecuteQueryDef_2Param_CheckInsertedCount()
   
   Dim qdf As DAO.QueryDef
   Dim Actual As Long
   Const Expected As Long = 1
   
   Set qdf = m_DaoHandler.CurrentDb.CreateQueryDef("", InsertParamSql)
   
   Actual = m_DaoHandler.ExecuteQueryDef(qdf, "abc", 123)
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ExecuteQueryDefByName_0Param_CheckInsertedCount()
   
   Dim qdf As DAO.QueryDef
   Dim Actual As Long
   Const Expected As Long = 1
   
   Set qdf = m_DaoHandler.CurrentDb.QueryDefs(TestQueryDefNameInsert)
   qdf.SQL = FormatString(InsertSqlTemplate, "abc", 123)
   qdf.Close
   
   Actual = m_DaoHandler.ExecuteQueryDefByName(TestQueryDefNameInsert)
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ExecuteQueryDefByName_2Param_CheckInsertedCount()
   
   Dim qdf As DAO.QueryDef
   Dim Actual As Long
   Const Expected As Long = 1
   
   Actual = m_DaoHandler.ExecuteQueryDefByName(TestQueryDefNameInsert, "abc", 123)
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ExecuteParamSql_InsertRecord_RecordsAffectedIs1()

   Const Expected As Long = 1
   Dim Actual As Long
   Dim SqlText As String
   
   Actual = m_DaoHandler.ExecuteParamSQL(InsertParamSql, "abc", 123)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

Public Sub ExecuteQueryDefByName_UseParamArray_CheckAffected()

   Const Expected As Long = 1
   Dim Actual As Long
   
   Actual = m_DaoHandler.ExecuteQueryDefByName(TestQueryDefNameInsert, "abc", 123)
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ExecuteQueryDefByName_UseGetParamDefArray_CheckAffected()

   Const Expected As Long = 1
   Dim Actual As Long
   
   Actual = m_DaoHandler.ExecuteQueryDefByName(TestQueryDefNameInsert, m_DaoHandler.GetParamDefArray("abc", 123))
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ExecuteQueryDefByName_UseGetNamedParamDefArray_CheckAffected()

   Const Expected As Long = 1
   Dim Actual As Long
   
   Actual = m_DaoHandler.ExecuteQueryDefByName(TestQueryDefNameInsert, m_DaoHandler.GetNamedParamDefArray("P1", "abc", "P2", 123))
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordset_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   
   Dim Actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordset(FormatString(SelectSqlTemplate, Id))
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordsetParamSQL_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   Const SelectSql As String = "Parameters [FilterId] long; select ID, T, Z from " & TestTableName & " where ID=[FilterId]"
   Dim Actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordsetParamSQL(SelectSql, dbOpenForwardOnly, dbFailOnError, dbOptimistic, Id)
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordsetParamSQL2_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   Const SelectSql As String = "Parameters [FilterId] long; select ID, T, Z from " & TestTableName & " where ID=[FilterId]"
   Dim Actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordsetParamSQL2(SelectSql, Id)
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub OpenRecordsetQueryDefByName_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   Dim Actual As Variant
   Dim rst As DAO.Recordset
   
   Dim QdfParamDefs(0, 1) As Variant
   
   QdfParamDefs(0, 0) = "FilterId"
   QdfParamDefs(0, 1) = Id

   Set rst = m_DaoHandler.OpenRecordsetQueryDefByName(TestQueryDefNameSelect, QdfParamDefs, dbOpenForwardOnly, dbFailOnError, dbOptimistic)
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
'AccUnit:Row(0, 2, 2)
'AccUnit:Row(1, 2, "b")
'AccUnit:Row(2, 2, 2)
Public Sub DLookupSQL_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldIndex As Variant, _
                                                                    ByVal Id As Long, _
                                                                    ByVal Expected As Variant)

   Const SelectDDLTemplate As String = "select ID, T, Z from " & TestTableName & " where ID={0}"
   Dim Actual As Variant
   
   Actual = m_DaoHandler.LookupSQL(FormatString(SelectDDLTemplate, Id), FieldIndex)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("Z", 2, 2)
Public Sub Lookup_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldName As Variant, _
                                                                ByVal Id As Long, _
                                                                ByVal Expected As Variant)
   Dim Actual As Variant
   Actual = m_DaoHandler.Lookup(FieldName, TestTableName, FormatString("ID={0}", Id))
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("=", 2, 1)
'AccUnit:Row("=", 0, 0)
'AccUnit:Row("<=", 2 , 2)
Public Sub Count_FilterRecordWithID_ReturnCount(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Long)
   Dim Actual As Variant
   Actual = m_DaoHandler.Count("*", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 2)
'AccUnit:Row(">=", 2 , 5)
Public Sub Max_FilterRecordWithID_ReturnMaxZ(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Max("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(Expected) Then
      Assert.IsNull Actual
   Else
      Assert.AreEqual Expected, Actual
   End If
   
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 1)
'AccUnit:Row(">=", 2 , 2)
Public Sub Min_FilterRecordWithID_ReturnMinZ(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Min("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(Expected) Then
      Assert.IsNull Actual
   Else
      Assert.AreEqual Expected, Actual
   End If
   
End Sub

'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 3)
'AccUnit:Row(">=", 2 , 14)
Public Sub Sum_FilterRecordWithID_ReturnSumZ(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Sum("Z", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(Expected) Then
      Assert.IsNull Actual
   Else
      Assert.AreEqual Expected, Actual
   End If
   
End Sub

Public Sub InsertIdentityReturn_InsertNewRecord_ReturnIdentityFromNewRecord()

   Const NewZ As Long = 999
   Dim Actual As Long
   Dim Expected As Long
   
   Actual = m_DaoHandler.InsertIdentityReturn(FormatString(InsertSqlTemplate, "ABC", NewZ))
   Expected = Nz(DLookup("ID", TestTableName, "Z=" & NewZ), 0)
   Assert.AreEqual Expected, Actual
   
End Sub

Public Sub CreateParamDefArray_Use4ParamArrays_CheckArray()

   Dim Actual() As Variant
   Dim Expected(3, 1) As Variant
   Dim i As Long
   
   For i = 0 To 3
      Expected(i, 0) = i
      Expected(i, 1) = i * 2 + 1
   Next
   
   Actual = m_DaoHandler.GetParamDefArray(1, 3, 5, 7)
   
   Assert.AreEqual Expected, Actual

End Sub

'AccUnit:Row(4)
'AccUnit:Row(1)
'AccUnit:Row(0)
Public Sub GetParamDefArray_UseArray_CheckArray(ByVal ParamCount As Long)
   
   Dim Actual() As Variant
   Dim Expected() As Variant
   Dim Params() As Variant
   Dim i As Long
   Dim ArraySize As Long
   
   If ParamCount = 0 Then
      Actual = m_DaoHandler.GetParamDefArray()
   Else
   
      ArraySize = ParamCount - 1
      ReDim Expected(ArraySize, 1)
      ReDim Params(ArraySize)
   
      For i = 0 To ArraySize
         Params(i) = i * 2 + 1
         Expected(i, 0) = i
         Expected(i, 1) = Params(i)
      Next
   
      Actual = m_DaoHandler.GetParamDefArray(Params)
   
   End If
   
   Assert.AreEqual Expected, Actual

End Sub

Public Sub GetNamedParamDefArray_With2ParamPairs_CheckArray()

   Dim Actual() As Variant
   Dim Expected(1, 1) As Variant
   Dim i As Long
   
   For i = 0 To 1
      Expected(i, 0) = Chr(Asc("a") + i)
      Expected(i, 1) = i * 2 + 1
   Next
   
   Actual = m_DaoHandler.GetNamedParamDefArray("a", 1, "b", 3)
   
   Assert.AreEqual Expected, Actual

End Sub

Public Sub GetNamedParamDefArray_WithWrongParamPair_ExpectedRaiseError()

   Dim Actual() As Variant

   Assert.Throws vbObjectError, "DaoHandler.GetNamedParamDefArray", "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
   Actual = m_DaoHandler.GetNamedParamDefArray("a")
   
End Sub

Public Sub GetNamedParamDefArray_WithWrongParamPair2_ExpectedRaiseError()

   Dim Actual() As Variant

   Assert.Throws vbObjectError, "DaoHandler.GetNamedParamDefArray", "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
   Actual = m_DaoHandler.GetNamedParamDefArray("a", 1, "b")
   
End Sub

'AccUnit:Row(0, 0, 0, 0)
'AccUnit:Row(0, 4, 0, 0) 'dbReadOnly = 4
'AccUnit:Row(0, 0, 4, 3001) 'dbReadOnly = 4
'AccUnit:Row(0, 4, 1, 3001) 'dbReadOnly = 4, dbOptimisticValue = 1
'AccUnit:Row(0, 4, 2, 3001) 'dbReadOnly = 4, dbPessimistic = 2
'AccUnit:Row(0, 4, 3, 3001) 'dbReadOnly = 4, dbOptimistic = 3
'AccUnit:Row(0, 4, 4, 3001) 'dbReadOnly = 4, dbReadOnly = 4
'AccUnit:Row(0, 4, 5, 3001) 'dbReadOnly = 4, dbOptimisticBatch = 5
Public Sub OpenRecordset_EnumValues_CheckErrNumber(ByVal RecordsetTyp As Long, ByVal RecordsetOptions As Long, LockEdit As Long, _
                                                             ByVal ExpectedErrorNumber As Long)

   Dim rst As DAO.Recordset
   
   If ExpectedErrorNumber <> 0 Then
      Assert.Throws ExpectedErrorNumber
   End If
   
   Set rst = m_DaoHandler.OpenRecordset("select * from " & TestTableName, RecordsetTyp, RecordsetOptions, LockEdit)
   rst.Close

End Sub

'AccUnit:Row(0, 0, 0)
'AccUnit:Row(0, 4, 0) 'dbReadOnly = 4
'AccUnit:Row(0, 0, 4) 'dbReadOnly = 4
'AccUnit:Row(0, 4, 1) 'dbReadOnly = 4, dbOptimisticValue = 1
'AccUnit:Row(0, 4, 2) 'dbReadOnly = 4, dbPessimistic = 2
'AccUnit:Row(0, 4, 3) 'dbReadOnly = 4, dbOptimistic = 3
'AccUnit:Row(0, 4, 4) 'dbReadOnly = 4, dbOptimistic = 4
'AccUnit:Row(0, 4, 5) 'dbReadOnly = 4, dbOptimisticBatch = 5
Public Sub OpenRecordset_EnumValues_ErrorLikeDaoOpenRecordset(ByVal RecordsetTyp As Long, ByVal RecordsetOptions As Long, ByVal LockEdit As Long)

   Dim rst As DAO.Recordset
   Dim Actual As Boolean
   Const Expected As Boolean = False
   
   Dim ExpectedErrorNumber As Long
   
   On Error Resume Next
   Set rst = CurrentDb.OpenRecordset("select * from " & TestTableName, Replace0WithIsMissing(RecordsetTyp), Replace0WithIsMissing(RecordsetOptions), Replace0WithIsMissing(LockEdit))
   
   If Err.Number <> 0 Then
      ExpectedErrorNumber = Err.Number
   Else
      rst.Close
   End If
   
   On Error GoTo 0
   
   If ExpectedErrorNumber <> 0 Then
      Assert.Throws ExpectedErrorNumber
   End If
   
   Set rst = m_DaoHandler.OpenRecordset("select * from " & TestTableName, RecordsetTyp, RecordsetOptions, LockEdit)
   rst.Close

End Sub

'AccUnit:Row(1, 2 , 3)
'AccUnit:Row(2, 3 , 5)
'AccUnit:Row(100, 3 , "EOF")
Public Sub ParamQdfSql_2ParamOpenRecordset_CheckID(ByVal Param1 As Long, ByVal Param2 As String, ByVal Expected As Variant)
' "Parameters P1 Long, P2 Long; select ID, T, Z from " & TestTableName & " where Z=[P1]+[P2]"
   
   Dim SqlText As String
   Dim rst As DAO.Recordset

   SqlText = SelectParamSql2

   Set rst = m_DaoHandler.ParamQueryDefSql(SqlText, Param1, Param2).OpenRecordset()

   Dim Actual As Variant

   If rst.EOF Then
      Actual = "EOF"
   Else
      Actual = rst.Fields("id")
   End If

   Assert.AreEqual Expected, Actual

End Sub

'AccUnit:Row(1, 2 , 3)
'AccUnit:Row(2, 3 , 5)
'AccUnit:Row(100, 3 , "EOF")
Public Sub ParamQdfByName_2ParamOpenRecordset_CheckID(ByVal Param1 As Long, ByVal Param2 As String, ByVal Expected As Variant)
' "Parameters P1 Long, P2 Long; select ID, T, Z from " & TestTableName & " where Z=[P1]+[P2]"
   
   Dim rst As DAO.Recordset
   Set rst = m_DaoHandler.ParamQueryDefByName(TestQueryDefNameSelect2, Param1, Param2).OpenRecordset()

   Dim Actual As Variant

   If rst.EOF Then
      Actual = "EOF"
   Else
      Actual = rst.Fields("id")
   End If

   Assert.AreEqual Expected, Actual

End Sub

'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Function Replace0WithIsMissing(ByVal CheckValue As Long) As Variant
   If CheckValue = 0 Then
      Replace0WithIsMissing = GetIsMissing()
   Else
      Replace0WithIsMissing = CheckValue
   End If
End Function

Private Function GetIsMissing(Optional DoNotFillThisParam As Variant) As Variant
   GetIsMissing = DoNotFillThisParam
End Function

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
   db.CreateQueryDef TestQueryDefNameSelect2, SelectParamSql2
   db.CreateQueryDef TestQueryDefNameInsert, InsertParamSql
   
End Sub

Private Sub DropTestQueryDef()
   
   Dim db As DAO.Database
   
   Set db = CodeDb
   
   db.QueryDefs.Delete TestQueryDefNameSelect
   db.QueryDefs.Delete TestQueryDefNameSelect2
   db.QueryDefs.Delete TestQueryDefNameInsert
   
End Sub
