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
'VBA-Inspector:Ignore

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
Public Sub IsNullOrEmpty_WithIgnoreSpaceParameter_ReturnsTrueOrFalse( _
               ByVal vValue As Variant, ByVal IgnoreSpaces As Boolean, ByVal expected As Boolean)
   Dim actual As Boolean
   actual = StringTools.IsNullOrEmpty(vValue, IgnoreSpaces)
   Assert.AreEqual expected, actual
End Sub



'AccUnit:Row("{0}", "A", "A")
'AccUnit:Row("a{0}", "A", "aA")
'AccUnit:Row("a{0}b", "A", "aAb")
'AccUnit:Row("{0}b", "A", "Ab")
'AccUnit:Row("ab", "A", "ab")
'AccUnit:Row("", "A", "")
Public Sub FormatText_1Arg_ReturnsFormattedText( _
               ByVal TextToFormat As String, ByVal Arg As Variant, ByVal expected As String)
   Dim actual As String
   actual = StringTools.FormatText(TextToFormat, Arg)
   Assert.AreEqual expected, actual
End Sub

'AccUnit:Row("{0}{1}", "A", "B",  "AB")
'AccUnit:Row("a{0}b{1}", "A", "B", "aAbB")
Public Sub FormatText_2Arg_ReturnsFormattedText( _
               ByVal TextToFormat As String, ByVal Arg1 As Variant, ByVal Arg2 As Variant, ByVal expected As String)
   Dim actual As String
   actual = StringTools.FormatText(TextToFormat, Arg1, Arg2)
   Assert.AreEqual expected, actual
End Sub
