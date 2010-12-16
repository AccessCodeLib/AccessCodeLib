VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "modSQL_Tools_FormatTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: modSQL_Tools_FormatTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' TEST_modSQL_Tools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/modSQL_Tools_FormatTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/modSQL_Tools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'AccUnit:TestClass

'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("abcdefg", "'", False, "'abcdefg'")
'AccUnit:Row("abc'defg", "'", False, "'abc''defg'")
'AccUnit:Row("abc""defg", "'", False, "'abc""defg'")
Public Function GetSQLString_Text_JetSQL( _
                     ByVal vValue As Variant, _
                     ByVal sDelimiter As String, _
                     ByVal bWithoutLeftRightDelim As Boolean, _
                     ByVal expected As String) As String

   Dim actual As String
   actual = GetSQLString_Text(vValue, sDelimiter, bWithoutLeftRightDelim)
   Assert.That actual, Iz.EqualTo(expected)

End Function

'AccUnit:Row("abc'defg", "'", False, "abc'defg")
Public Sub GetSQLString_Text_JetSQL_FailedTest( _
                     ByVal vValue As Variant, _
                     ByVal sDelimiter As String, _
                     ByVal bWithoutLeftRightDelim As Boolean, _
                     ByVal sNotExpect As String)

   Dim actual As String
   actual = GetSQLString_Text(vValue, sDelimiter, bWithoutLeftRightDelim)
   
   Assert.That actual, Iz.Not.EqualTo(sNotExpect)

End Sub
