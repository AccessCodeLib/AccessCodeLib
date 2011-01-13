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
'  <use>data/SqlTools.bas</use>
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
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()


End Sub

Public Sub TearDown()

End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("abcdefg", "'", False, "'abcdefg'")
'AccUnit:Row("abc'defg", "'", False, "'abc''defg'")
'AccUnit:Row("abc""defg", "'", False, "'abc""defg'")
Public Function FormatTextToSqlText_JetSQL( _
                     ByVal vValue As Variant, _
                     ByVal sDelimiter As String, _
                     ByVal bWithoutLeftRightDelim As Boolean, _
                     ByVal expected As String) As String

   Dim actual As String
   actual = FormatTextToSqlText(vValue, sDelimiter, bWithoutLeftRightDelim)
   Assert.That actual, Iz.EqualTo(expected)

End Function

'AccUnit:Row("abcdefg", "'", False, "'abcdefg'")
'AccUnit:Row("abc'defg", "'", False, "'abc''defg'")
'AccUnit:Row("abc""defg", "'", False, "'abc""defg'")
Public Function FormatTextToSqlText_JetSQL2( _
                     ByVal vValue As Variant, _
                     ByVal sDelimiter As String, _
                     ByVal bWithoutLeftRightDelim As Boolean, _
                     ByVal expected As String) As String

   Dim actual As String
   actual = FormatTextToSqlText(vValue, sDelimiter, bWithoutLeftRightDelim)
   Assert.That actual, Iz.EqualTo(expected)

End Function

'AccUnit:Row("abc'defg", "'", False, "abc'defg")
Public Sub FormatTextToSqlText_JetSQL_FailedTest( _
                     ByVal vValue As Variant, _
                     ByVal sDelimiter As String, _
                     ByVal bWithoutLeftRightDelim As Boolean, _
                     ByVal sNotExpect As String)

   Dim actual As String
   actual = FormatTextToSqlText(vValue, sDelimiter, bWithoutLeftRightDelim)
   
   Assert.That actual, Iz.Not.EqualTo(sNotExpect)

End Sub
