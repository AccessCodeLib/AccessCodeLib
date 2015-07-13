VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterStringBuilderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' FilterStringBuilderTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse f�r FilterStringBuilder
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/FilterStringBuilderTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/FilterStringBuilder.cls</replace>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub Add_StringNumericDateBooleanFilter_CheckDoaSyntax()
   
   Const Expected As String = "(S Like 'a*') And (N >= 5.1) And (D < #2015-02-01#) And (B = True)"
   
   Dim Actual As Variant
   
   With New FilterStringBuilder
      
      Set .SqlTool = SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
      
      .Add "S", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
      .Add "N", SQL_Numeric, SQL_GreaterThan + SQL_Equal, 5.1
      .Add "D", SQL_Date, SQL_LessThan, #2/1/2015#
      .Add "B", SQL_Boolean, SQL_Equal, True
      
      Actual = .ToString(SQL_And)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddCriteria_AddConString_CheckSqlText()
   
   Const Expected As String = "(a > b)"
   
   Dim Actual As Variant
   
   With New FilterStringBuilder
   
      .AddCriteria ("a > b")
      
      Actual = .ToString(SQL_And)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ConfigSqlFormat_StateUnderTest_ExpectedBehaviour()
   
   Const TestDateFormat As String = "\#YYYY-mm-DD\#"
   Const TestBooleanString As String = "T"
   Const TestWildCardString As String = "?"
   
   With New FilterStringBuilder
   
      .ConfigSqlFormat TestDateFormat, TestBooleanString, TestWildCardString
      
      Assert.That .SqlTool.SqlDateFormat, Iz.EqualTo(TestDateFormat)
      Assert.That .SqlTool.SqlBooleanTrueString, Iz.EqualTo(TestBooleanString)
      Assert.That .SqlTool.SqlWildCardString, Iz.EqualTo(TestWildCardString)

   End With
   
End Sub