VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlCollectionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' FilterControlCollectionTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für StringCollection
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/form/filter/FilterControlCollectionTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControlCollection.cls</use>
'  <use>_test/form/filter/FilterControlTestForm.frm</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private TestForm As Form_FilterControlTestForm
Private FilterCtrCol As FilterControlCollection
Private TestSqlTools As SqlTools

'--------------------------------------------------------------------
' Test Preparation / Cleanup
'--------------------------------------------------------------------
Public Sub Setup()
   Set TestForm = New Form_FilterControlTestForm
   TestForm.Visible = True
   Set FilterCtrCol = New FilterControlCollection
   Set TestSqlTools = SqlTools.FromDialect(SqlDialect.SQL_DAO, "\#yyyy-mm-dd\#")
End Sub

Public Sub TearDown()
   Set FilterCtrCol = Nothing
   Set TestForm = Nothing
   Set TestSqlTools = Nothing
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("T", 3, 2, "fctlTextFld", "abc", "", Null, Null, "(T = 'abc')")
'AccUnit:Row("N", 2, 2, "fctlTextFld", "1", "", Null, Null, "(N = 1)")
'AccUnit:Row("D", 4, 512, "fctlDateFldMin", System.Convert.ToDateTime("2003-02-01"), "fctlDateFldMax", Null, Null, "(D >= #2003-02-01#)")
'AccUnit:Row("D", 4, 512, "fctlDateFldMin", System.Convert.ToDateTime("2003-02-01"), "fctlDateFldMax", System.Convert.ToDateTime("2003-03-12"), Null, "(D Between #2003-02-01# And #2003-03-12#)")
Public Sub Add_SetFilterControlValue_CheckSql(ByVal DataFieldName As String, ByVal DataType As SqlFieldDataType, ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal ControlName As String, ByVal DataValue As Variant, _
               ByVal Control2Name As String, ByVal DataValue2 As Variant, ByVal IgnoreValue As Variant, _
               ByVal Expected As String)
 
   Dim Actual As String
   Dim ControlRef As Control
   Dim Control2Ref As Control
   
   Set ControlRef = TestForm.Controls(ControlName)
   ControlRef.Value = DataValue
   If Len(Control2Name) > 0 Then
      Set Control2Ref = TestForm.Controls(Control2Name)
      Control2Ref.Value = DataValue2
   End If
 
   FilterCtrCol.Add DataFieldName, DataType, RelationalOperator, ControlRef, Control2Ref, IgnoreValue
   
   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub AddGroup_OrGroup_CheckSql()
   
   Dim Actual As String
   Const Expected As String = "(T = 'abc') And ((N = 123) Or (B = True))"
   
   TestForm.fctlTextFld.Value = "abc"
   TestForm.fctlNumIntFld.Value = 123
   TestForm.fctlBooleanFld.Value = True
   
   With FilterCtrCol
      .Add "T", SQL_Text, SQL_Equal, TestForm.fctlTextFld
      With .AddGroup(SQL_Or)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With
   
   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub AddSubSelectCriteria_InWith2Param_CheckSql()

   Dim Actual As String
   Const Expected As String = "(T In (Select Tx From Tab123 Where (N = 123) And (B = True)))"
   
   TestForm.fctlNumIntFld.Value = 123
   TestForm.fctlBooleanFld.Value = True
   
   With FilterCtrCol
      With .AddSubSelectCriteria("T", SQL_In, "Select Tx From Tab123", True)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With
   
   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub AddSubSelectCriteria_InWith2ParamAndWhereStatement_CheckSql()

   Dim Actual As String
   Const Expected As String = "(T In (Select Tx From Tab123 Where ((N = 123) And (B = True)) And X=123))"
   
   TestForm.fctlNumIntFld.Value = 123
   TestForm.fctlBooleanFld.Value = True
   
   With FilterCtrCol
      With .AddSubSelectCriteria("T", SQL_In, "Select Tx From Tab123 [WhereStatement] And X=123", True, True)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With
   
   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub AddSubSelectCriteria_InWithNullParam_CheckSql()

   Dim Actual As String
   Const Expected As String = ""
   
   TestForm.fctlNumIntFld.Value = Null
   TestForm.fctlBooleanFld.Value = Null
   
   With FilterCtrCol
      With .AddSubSelectCriteria("T", SQL_In, "Select Tx From Tab123", True)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With
   
   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)
   
End Sub


Public Sub AddExistsCriteria_With2Param_CheckSql()

   Dim Actual As String
   Const Expected As String = "(Exists (Select 1 From Tab123 Where (N = 123) And (B = True)))"

   TestForm.fctlNumIntFld.Value = 123
   TestForm.fctlBooleanFld.Value = True

   With FilterCtrCol
      With .AddExistsCriteria("Select 1 From Tab123", True)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With

   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub


Public Sub AddExistsCriteria_With2ParamAndWhereStatement_CheckSql()

   Dim Actual As String
   Const Expected As String = "(Exists (Select 1 From Tab123 AS T Where ((N = 123) And (B = True)) And T.fiXyz = S.idXyz))"

   TestForm.fctlNumIntFld.Value = 123
   TestForm.fctlBooleanFld.Value = True

   With FilterCtrCol
      With .AddExistsCriteria("Select 1 From Tab123 AS T [WhereStatement] And T.fiXyz = S.idXyz", True, SQL_And, False, True)
         .Add "N", SQL_Numeric, SQL_Equal, TestForm.fctlNumIntFld
         .Add "B", SQL_Boolean, SQL_Equal, TestForm.fctlBooleanFld
      End With
   End With

   Actual = FilterCtrCol.CriteriaString(TestSqlTools)
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
