VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlToolsBuildCriteriaTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' SqlToolsBuildCriteriaTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für SqlTools.BuildCriteria
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/SqlToolsBuildCriteriaTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</replace>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

Private m_SqlTools As SqlTools

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   Set m_SqlTools = New SqlTools
End Sub

Public Sub Teardown()
   Set m_SqlTools = Nothing
End Sub

Private Property Get DaoTools() As SqlTools
   Set DaoTools = m_SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
End Property

Private Property Get TSqlTools() As SqlTools
   Set TSqlTools = m_SqlTools.Clone("'yyyymmdd'", "1", "%")
End Property

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row(2, Null, Null, Null, "", "").Name("Emtpy")
'AccUnit:Row(2, "a", Null, Null, "F = 'a'", "F = 'a'").Name("Equal")
'AccUnit:Row(2+8, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, "a", Null, Null, "F <= 'a'", "F <= 'a'").Name("Equal+LessThan")
'AccUnit:Row(256, "a", Null, Null, "F Like 'a'", "F Like 'a'").Name("Like")
'AccUnit:Row(256, "a*", Null, Null, "F Like 'a*'", "F Like 'a%'").Name("Like+*")
'AccUnit:Row(256+2048, "a", Null, Null, "F Like 'a*'", "F Like 'a%'").Name("Like+WildCardSuffix")
'AccUnit:Row(256+4096, "a", Null, Null, "F Like '*a'", "F Like '%a'").Name("Like+WildCardPrefix")
'AccUnit:Row(256+2048+4096, "a", Null, Null, "F Like '*a*'", "F Like '%a%'").Name("Like+WildCardSuffix+Prefix")
'AccUnit:Row(512, "a", "d", Null, "F Between 'a' And 'd'", "F Between 'a' And 'd'").Name("Between")
'AccUnit:Row(512, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Between+V2=Null")
'AccUnit:Row(512, Null, "d", Null, "F <= 'd'", "F <= 'd'").Name("Between+V1=Null")
'AccUnit:Row(1 + 256, "d", Null, Null, "F Not Like 'd'", "F Not Like 'd'").Name("not like")
'AccUnit:Row(1 + 512, "a", "d", Null, "F < 'a' And F > 'd'", "F < 'a' And F > 'd'").Name("Not Between")
'AccUnit:Row(1+2, "a", Null, Null, "F <> 'a'", "F <> 'a'").Name("Not Equal")
'AccUnit:Row(1+2+4, "a", Null, Null, "F > 'a'", "F > 'a'").Name("Not <=")
'AccUnit:Row(1+4, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Not <")
'AccUnit:Row(1+2+8, "a", Null, Null, "F < 'a'", "F < 'a'").Name("Not >=")
Public Sub BuildCriteria_TextFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDao As String, ByVal ExpectedTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedDao)

   Actual = TSqlTools.BuildCriteria("F", SQL_Text, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedTSql)

End Sub

'AccUnit:Row(2, Null, Null, Null, "").Name("Emtpy")
'AccUnit:Row(2, 123, Null, Null, "F = 123").Name("Equal")
'AccUnit:Row(2, 1.23, Null, Null, "F = 1.23").Name("Equal")
'AccUnit:Row(2+8, 123, Null, Null, "F >= 123").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, 123, Null, Null, "F <= 123").Name("Equal+LessThan")
'AccUnit:Row(512, 123, 456, Null, "F Between 123 And 456").Name("Between")
'AccUnit:Row(512, 123, Null, Null, "F >= 123").Name("Between+V2=Null")
'AccUnit:Row(512, Null, 456, Null, "F <= 456").Name("Between+V1=Null")
Public Sub BuildCriteria_NumericFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDaoOrTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Numeric, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedDaoOrTSql)

   Actual = TSqlTools.BuildCriteria("F", SQL_Numeric, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedDaoOrTSql)

End Sub

'AccUnit:Row(2, Null, Null, Null, "", "").Name("Emtpy")
'AccUnit:Row(2, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F = #2014-05-16#", "F = '20140516'").Name("Equal")
'AccUnit:Row(2+8, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F >= #2014-05-16#", "F >= '20140516'").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F <= #2014-05-16#", "F <= '20140516'").Name("Equal+LessThan")
'AccUnit:Row(2+4+2048, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F < #2014-05-17#", "F < '20140517'").Name("Equal+LessThan+WildCardSuffix")
'AccUnit:Row(512, System.Convert.ToDateTime("2014-05-16"), System.Convert.ToDateTime("2014-07-18"), Null, "F Between #2014-05-16# And #2014-07-18#", "F Between '20140516' And '20140718'").Name("Between+LessThan+WildCardSuffix")
'AccUnit:Row(512+2048, System.Convert.ToDateTime("2014-05-16"), System.Convert.ToDateTime("2014-07-18"), Null, "F >= #2014-05-16# And F < #2014-07-19#", "F >= '20140516' And F < '20140719'").Name("Between+LessThan+WildCardSuffix")
Public Sub BuildCriteria_DateFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDao As String, ByVal ExpectedTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedDao)

   Actual = TSqlTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That Actual, Iz.EqualTo(ExpectedTSql)

End Sub

'AccUnit:Row(2, Null, Null, Null, "").Name("Emtpy")
'AccUnit:Row(2, "a", "b", "c", "F = 'a' Or F = 'b' Or F = 'c'").Name("Equal")
'AccUnit:Row(2+8, "a", "b", "c", "F >= 'a' Or F >= 'b' Or F >= 'c'").Name("Equal+GreaterThan")
'AccUnit:Row(256, "a*", "b*", "c*", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'").Name("Like+*")
'AccUnit:Row(256+2048, "a", "b", "c", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'").Name("Like+WildCardSuffix")
'AccUnit:Row(256+2048+4096, "a", "b", "c", "F Like '*a*' Or F Like '*b*' Or F Like '*c*'").Name("Like+WildCardSuffix+Prefix")
'AccUnit:Row(1024, "a", "b", "c", "F In ('a','b','c')").Name("In")
'AccUnit:Row(1+1024, "a", "b", "c", "F Not In ('a','b','c')").Name("Not In")
Public Sub BuildCriteria_TextValueArrayFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue1 As Variant, ByVal FilterValue2 As Variant, ByVal FilterValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, RelationalOperator, Array(FilterValue1, FilterValue2, FilterValue3))
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(Null, "a", "b", Null, "")
'AccUnit:Row("a", "a", "b", Null, "")
'AccUnit:Row("x", "a", "b", Null, "F = 'x'")
Public Sub BuildCriteria_IgnoreValueArray( _
                           ByVal FilterValue1 As Variant, _
                           ByVal IgnoreValue1 As Variant, ByVal IgnoreValue2 As Variant, ByVal IgnoreValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_Equal, FilterValue1, , Array(IgnoreValue1, IgnoreValue2, IgnoreValue3))
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(Null, "a", "b", "a", "b", Null, "")
'AccUnit:Row(Null, "a", "b", "a", "b", "c", "F Is Null")
'AccUnit:Row(Null, "a", "b", "x", "y", "z", "F Is Null Or F In ('a','b')")
Public Sub BuildCriteria_InAndIgnoreValueArray( _
                           ByVal FilterValue1 As Variant, ByVal FilterValue2 As Variant, ByVal FilterValue3 As Variant, _
                           ByVal IgnoreValue1 As Variant, ByVal IgnoreValue2 As Variant, ByVal IgnoreValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_In, Array(FilterValue1, FilterValue2, FilterValue3), , Array(IgnoreValue1, IgnoreValue2, IgnoreValue3))
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
