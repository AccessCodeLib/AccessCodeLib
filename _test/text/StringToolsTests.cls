VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/text/StringToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>text/StringTools.bas</use>
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

'AccUnit:Row("abc", False)
'AccUnit:Row("", True)
'AccUnit:Row(" ", False).Name = "1 Space"
'AccUnit:Row(Null, True)
'AccUnit:Row(1, False)
Public Sub IsNullOrEmpty_WithoutIgnoreSpaceParameter_ReturnsTrueOrFalse(ByVal vValue As Variant, ByVal expected As Boolean)
   Dim actual As Boolean
   actual = StringTools.IsNullOrEmpty(vValue)
   Assert.AreEqual expected, actual
End Sub


'AccUnit:Row(" ", False, False).Name = "1 Space"
'AccUnit:Row(" ", True, True).Name = "1 Space + Ignore"
'AccUnit:Row("abc", False, False)
'AccUnit:Row("abc", True, False)
'AccUnit:Row("", False, True)
'AccUnit:Row("", True, True)
'AccUnit:Row(Null, True, True)
'AccUnit:Row(1, True, False)
Public Sub IsNullOrEmpty_WithIgnoreSpaceParameter_ReturnsTrueOrFalse(ByVal vValue As Variant, ByVal IgnoreSpaces As Boolean, ByVal expected As Boolean)
   Dim actual As Boolean
   actual = StringTools.IsNullOrEmpty(vValue, IgnoreSpaces)
   Assert.AreEqual expected, actual
End Sub
