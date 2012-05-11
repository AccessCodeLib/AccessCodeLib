VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' AccUnit:TestClass
'--------------------------------------------------------------------
'<codelib>
'  <file>_test/text/StringToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>text/StringTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'VBA-Inspector:Ignore

'--------------------------------------------------------------------
' AccUnit Infrastructure
'--------------------------------------------------------------------

Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'AccUnit:Row("abc", False)
'AccUnit:Row("", True)
'AccUnit:Row(" ", False).Name = "1 Space"
'AccUnit:Row(Null, True)
'AccUnit:Row(1, False)
Public Sub IsNullOrEmpty_WithoutIgnoreSpaceParameter_ReturnsTrueOrFalse(ByVal vValue As Variant, ByVal Expected As Boolean)
   Dim Actual As Boolean
   Actual = StringTools.IsNullOrEmpty(vValue)
   Assert.AreEqual Expected, Actual
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
               ByVal vValue As Variant, ByVal IgnoreSpaces As Boolean, ByVal Expected As Boolean)
   Dim Actual As Boolean
   Actual = StringTools.IsNullOrEmpty(vValue, IgnoreSpaces)
   Assert.AreEqual Expected, Actual
End Sub

'AccUnit:Row("{0}", "A", "A")
'AccUnit:Row("a{0}", "A", "aA")
'AccUnit:Row("a{0}b", "A", "aAb")
'AccUnit:Row("{0}b", "A", "Ab")
'AccUnit:Row("ab", "A", "ab")
'AccUnit:Row("", "A", "")
Public Sub FormatText_1Arg_ReturnsFormattedText( _
               ByVal TextToFormat As String, ByVal Arg As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = StringTools.FormatText(TextToFormat, Arg)
   Assert.AreEqual Expected, Actual
End Sub

'AccUnit:Row("{0}{1}", "A", "B",  "AB")
'AccUnit:Row("a{0}b{1}", "A", "B", "aAbB")
Public Sub FormatText_2Arg_ReturnsFormattedText( _
               ByVal TextToFormat As String, ByVal Arg1 As Variant, ByVal Arg2 As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = StringTools.FormatText(TextToFormat, Arg1, Arg2)
   Assert.AreEqual Expected, Actual
End Sub

'AccUnit:Row(#1/1/2000 15:30:45#, "hh:nn:ss", "15:30:45")
'AccUnit:Row(#1/1/2000 15:30:45#, "dd.mm.yyyy hh:nn:ss", "01.01.2000 15:30:45")
'AccUnit:Row(#12/31/1899 01:30:45#, "hh:nn:ss", "01:30:45")
'AccUnit:Row(#12/31/1899 01:30:45#, "[h]:nn:ss", "25:30:45")
'AccUnit:Row(#1/1/1900 01:30:45#, "[H]:nn:ss", "49:30:45")
'AccUnit:Row(#12/30/1899 15:30:45#, "[H]:nn:ss", "15:30:45")
'AccUnit:Row(#12/30/1899 01:30:45#, "[h]:nn:ss", "1:30:45")
'AccUnit:Row(#12/30/1899 10:30:45#, "[hh]:nn:ss", "10:30:45")
'AccUnit:Row(#12/30/1899 01:30:45#, "[hh]:nn:ss", "01:30:45")
'AccUnit:Row(#1/1/1900 12:30:15#, "[hh]:nn:ss", "60:30:15")
'AccUnit:Row(#12/29/1899 14:34:56#, "[hh]:nn:ss", "-38:34:56")
Public Sub Format_DateTime_FormattedDateTime(ByVal Expression As Variant, ByVal FormatString As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = StringTools.Format(Expression, FormatString)
   Assert.AreEqual Expected, Actual
End Sub

'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "__Test")
Public Sub PadLeft_String_LeftPadString(ByVal Value As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)
    Assert.That StringTools.PadLeft(Value, totalWidth, padChar), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("Test", 4, "Test")
'AccUnit:Row("Test", 6, "  Test")
Public Sub PadLeft_String_LeftPadStringDefaultParam(ByVal Value As String, ByVal totalWidth As Integer, ByVal Expected As String)
    Assert.That StringTools.PadLeft(Value, totalWidth), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "Test__")
Public Sub PadRight_String_RightPadString(ByVal Value As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)
    Assert.That StringTools.PadRight(Value, totalWidth, padChar), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("Test", 4, "Test")
'AccUnit:Row("Test", 6, "Test  ")
Public Sub PadRight_String_RightPadStringDefaultParam(ByVal Value As String, ByVal totalWidth As Integer, ByVal Expected As String)
    Assert.That StringTools.PadRight(Value, totalWidth), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("abc", "b", True)
'AccUnit:Row("abc", "B", True)
'AccUnit:Row("abc", "", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", False)
'AccUnit:Row("", "A", False)
Public Sub Contains_String_ContainsSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That StringTools.Contains(checkValue, searchValue), Iz.EqualTo(Expected)
End Sub

