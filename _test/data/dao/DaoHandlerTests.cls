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
' Class: _test.data.dao.DaoHandlerTests
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'
' Test class for DaoHandler
'
' Author:
'     Josef Poetzl
'
' Call:
'     TestSuite.AddByClassName("DaoHandlerTests").Run
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/dao/DaoHandlerTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoHandler.cls</use>
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

Private Const TestTableCreateDDL As String = "CREATE TABLE " & TestTableName & " (id counter(1,1), T varchar(255), N int)"
Private Const TestTableDropDDL As String = "Drop TABLE " & TestTableName
Private Const SelectSqlTemplate As String = "select ID, T, N from " & TestTableName & " where ID={0}"
Private Const SelectParamSql As String = "select ID, T, N from " & TestTableName & " where ID=[FilterId]"
Private Const SelectParamSql2 As String = "Parameters [P1] Long, [P2] Long; select ID, T, N from " & TestTableName & " where N=[P1]+[P2]"
Private Const InsertSqlTemplate As String = "Insert INTO " & TestTableName & " (T, N) VALUES ('{0}', {1})"
Private Const InsertParamSql As String = "Parameters [P1] text(255), [P2] Long; Insert INTO " & TestTableName & " (T, N) VALUES ([P1], [P2])"
Private Const DeleteRecordsSQL As String = "DELETE FROM " & TestTableName

Public Sub Setup()
   Set m_DaoHandler = New DaoHandler
   Set m_DaoHandler.CurrentDb = CurrentDb
   
   CreateTestTable
   InsertRecordsInTestTable
   CreateTestQueryDefs
End Sub

Public Sub Teardown()
   m_DaoHandler.Dispose
   Set m_DaoHandler = Nothing
   'DropTestQueryDefs
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
   
   Actual = m_DaoHandler.ExecuteParamSql(InsertParamSql, "abc", 123)
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
'AccUnit:Row("N", 2, 2)
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
'AccUnit:Row("N", 2, 2)
Public Sub OpenRecordsetParamSQL_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   Const SelectSql As String = "Parameters [FilterId] long; select ID, T, N from " & TestTableName & " where ID=[FilterId]"
   Dim Actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordsetParamSql(SelectSql, dbOpenForwardOnly, dbFailOnError, dbOptimistic, Id)
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("N", 2, 2)
Public Sub OpenRecordsetParamSQL2_FilterRecordWithID_ReturnFieldValues(ByVal FieldName As String, ByVal Id As Long, ByVal Expected As Variant)
   
   Const SelectSql As String = "Parameters [FilterId] long; select ID, T, N from " & TestTableName & " where ID=[FilterId]"
   Dim Actual As Variant
   Dim rst As DAO.Recordset

   Set rst = m_DaoHandler.OpenRecordsetParamSql2(SelectSql, Id)
   Actual = rst.Fields(FieldName).Value
   rst.Close
   
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("N", 2, 2)
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
'AccUnit:Row("N", 2, 2)
'AccUnit:Row(0, 2, 2)
'AccUnit:Row(1, 2, "b")
'AccUnit:Row(2, 2, 2)
Public Sub DLookupSQL_FilterRecordWithID_ReturnValuesFromFieldIndex(ByVal FieldIndex As Variant, _
                                                                    ByVal Id As Long, _
                                                                    ByVal Expected As Variant)

   Const SelectDDLTemplate As String = "select ID, T, N from " & TestTableName & " where ID={0}"
   Dim Actual As Variant
   
   Actual = m_DaoHandler.LookupSql(FormatString(SelectDDLTemplate, Id), FieldIndex)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub

'AccUnit:Row("ID", 2, 2)
'AccUnit:Row("T", 2, "b")
'AccUnit:Row("N", 2, 2)
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

'AccUnit:Row("=", 2, True, 1)
'AccUnit:Row("=", 0, True, 0)
'AccUnit:Row("<=", 6, False, 6)
'AccUnit:Row("<=", 6, True, 5)
'AccUnit:Row("<=", 999, False, 12)
'AccUnit:Row("<=", 999, True, 7)
Public Sub DistinctCount_FilterRecordWithID_ReturnDistinctCount(ByVal Expr As String, ByVal Id As Long, _
                              ByVal UseDistinct As Boolean, ByVal Expected As Long)
   Dim Actual As Variant

   InsertRecordsInTestTable 7 ' => 1..5 inserted in Setup => 1..5 are therefore double

   Actual = m_DaoHandler.Count("N", TestTableName, FormatString("ID{0}{1}", Expr, Id), UseDistinct)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub


'AccUnit:Row("=", 2, 2)
'AccUnit:Row("=", 0, Null)
'AccUnit:Row("<=", 2 , 2)
'AccUnit:Row(">=", 2 , 5)
Public Sub Max_FilterRecordWithID_ReturnMaxN(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Max("N", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
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
Public Sub Min_FilterRecordWithID_ReturnMinN(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Min("N", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
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
Public Sub Sum_FilterRecordWithID_ReturnSumN(ByVal Expr As String, ByVal Id As Long, ByVal Expected As Variant)

   Dim Actual As Variant
   Actual = m_DaoHandler.Sum("N", TestTableName, FormatString("ID{0}{1}", Expr, Id))
   
   If IsNull(Expected) Then
      Assert.IsNull Actual
   Else
      Assert.AreEqual Expected, Actual
   End If
   
End Sub

Public Sub InsertIdentityReturn_InsertNewRecord_ReturnIdentityFromNewRecord()

   Const NewN As Long = 999
   Dim Actual As Long
   Dim Expected As Long
   
   Actual = m_DaoHandler.InsertIdentityReturn(FormatString(InsertSqlTemplate, "ABC", NewN))
   Expected = Nz(DLookup("ID", TestTableName, "N=" & NewN), 0)
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
   
   Assert.That Actual, Iz.EqualTo(Expected)

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

   Assert.Throws vbObjectError, "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
   Actual = m_DaoHandler.GetNamedParamDefArray("a")
   
End Sub

Public Sub GetNamedParamDefArray_WithWrongParamPair2_ExpectedRaiseError()

   Dim Actual() As Variant

   Assert.Throws vbObjectError, "ParamArray muss immer Paare für Parameter-Index und Parameter-Wert enthalten."
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
' "Parameters P1 Long, P2 Long; select ID, T, N from " & TestTableName & " where N=[P1]+[P2]"
   
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
' "Parameters P1 Long, P2 Long; select ID, T, N from " & TestTableName & " where N=[P1]+[P2]"
   
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

Private Function InsertRecordsInTestTable(Optional ByVal RecordNumToInsert As Long = 5) As Long
   Dim db As DAO.Database
   Dim i As Long
   
   Set db = CurrentDb
   
   For i = 1 To RecordNumToInsert
      db.Execute FormatString(InsertSqlTemplate, Chr(i + 96), i)
   Next
   
   If RecordCountInTestTable = 0 Then
      Err.Raise vbObjectError, "InsertRecordsInTestTable", "missing records"
   End If
   
   InsertRecordsInTestTable = RecordNumToInsert
   
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

Private Sub CreateTestQueryDefs()
   
   Dim db As DAO.Database
   
   Set db = CodeDb
   
   CreateTestQueryDef db, TestQueryDefNameSelect, SelectParamSql
   CreateTestQueryDef db, TestQueryDefNameSelect2, SelectParamSql2
   CreateTestQueryDef db, TestQueryDefNameInsert, InsertParamSql
   
End Sub

Private Sub CreateTestQueryDef(ByVal db As DAO.Database, ByVal QdfName As String, ByVal QdfSql As String)

On Error GoTo HandleErr

   db.CreateQueryDef QdfName, QdfSql
   
ExitHere:
   Exit Sub

HandleErr:
   Select Case Err.Number
      Case 3012
         db.QueryDefs(QdfName).SQL = QdfSql
         Resume ExitHere
      Case Else
         Err.Raise Err.Number, "CreateTestQueryDef:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
   End Select
   
End Sub

Private Sub DropTestQueryDefs()
   
   Dim db As DAO.Database
   
   Set db = CodeDb
   
   db.QueryDefs.Delete TestQueryDefNameSelect
   db.QueryDefs.Delete TestQueryDefNameSelect2
   db.QueryDefs.Delete TestQueryDefNameInsert
   
End Sub
