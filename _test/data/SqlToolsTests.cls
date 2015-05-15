VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für SqlTools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/SqlToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

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
   SqlTools.SqlDateFormat = vbNullString
End Sub

Public Sub Teardown()
   SqlTools.SqlDateFormat = vbNullString
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------


'
' TextToSqlText
'

'AccUnit:Row("abcdefg", "'", False, "'abcdefg'")
'AccUnit:Row("abcdefg", "'", True, "abcdefg")
'AccUnit:Row("abc'defg", "'", False, "'abc''defg'")
'AccUnit:Row("abc""defg", "'", False, "'abc""defg'")
'AccUnit:Row(Null, "'", False, "NULL")
Public Function TextToSqlText_DefaultValueIfNullArg_ReturnsSqlText( _
                     ByVal Value As Variant, _
                     ByVal Delimiter As String, _
                     ByVal WithoutLeftRightDelim As Boolean, _
                     ByVal Expected As String) As String

   Dim Actual As String
   Actual = SqlTools.TextToSqlText(Value, Delimiter, WithoutLeftRightDelim)
   Assert.That Actual, Iz.EqualTo(Expected)

End Function


'
' DateToSqlText
'

'AccUnit:Row(Null, "", "NULL")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#yyyy-mm-dd\#", "#2000-01-31#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#mm\/dd\/yyyy\#", "#01/31/2000#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "'yyyymmdd'", "'20000131'")
Public Sub DateToSqlText_UseFormatArg_ReturnsSqlText(ByVal Value As Variant, _
                     ByVal FormatString As String, _
                     ByVal Expected As String)
   Dim Actual As String
   Actual = SqlTools.DateToSqlText(Value, FormatString)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(Null, "", "NULL")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#yyyy-mm-dd\#", "#2000-01-31#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#mm\/dd\/yyyy\#", "#01/31/2000#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "'yyyymmdd'", "'20000131'")
Public Sub DateToSqlText_UsePublicFormatVariable_ReturnsSqlText(ByVal Value As Variant, _
                     ByVal FormatString As String, _
                     ByVal Expected As String)
   Dim Actual As String
   SqlTools.SqlDateFormat = FormatString
   Actual = SqlTools.DateToSqlText(Value)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DateToSqlText_FormatVariableNotDefined_RaiseError()
   Dim Actual As String
   Assert.Throws -2147221503, "DateToSqlText"
   Actual = SqlTools.DateToSqlText(Date)
End Sub


'
' NumberToSqlText
'

'AccUnit:Row(Null, "NULL")
'AccUnit:Row(0, "0")
'AccUnit:Row(000, "0")
'AccUnit:Row(12345, "12345")
'AccUnit:Row(1.2345, "1.2345")
'AccUnit:Row(0.123, "0.123")
'AccUnit:Row(.123, "0.123")
'AccUnit:Row(1E2, "100")
'AccUnit:Row(1E-2, "0.01")
Public Sub NumberToSqlText_UseNumbers_ReturnsSqlText(ByVal Value As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = SqlTools.NumberToSqlText(Value)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NumberToSqlText_UseText_ReturnsSqlText()
   Dim Actual As String
   Assert.Throws 13
   Actual = SqlTools.NumberToSqlText("abc")
End Sub
