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

'AccUnit:Row(#01/02/2000 14:34:56#, #01/01/2000 00:00:00#, "[hh]:nn:ss", "38:34:56")
'AccUnit:Row(2, 0, "[hh]:nn:ss", "48:00:00")
Public Sub Format_CalculatedDateTime_FormattedDateTime(ByVal Date1 As Variant, ByVal Date2 As Variant, ByVal FormatString As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = StringTools.Format(Date1 - Date2, FormatString)
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

'AccUnit:Row("abcd", "cd", True)
'AccUnit:Row("abcd", "CD", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_EndsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That StringTools.EndsWith(checkValue, searchValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("abcd", "ab", True)
'AccUnit:Row("abcd", "AB", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_StartsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That StringTools.StartsWith(checkValue, searchValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", 0)
'AccUnit:Row("1234567890", 10)
Public Sub Lenght_String_HasLenght(ByVal checkValue As String, Expected As Integer)
    Assert.That StringTools.Lenght(checkValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "", "")
'AccUnit:Row("A", "", "A")
'AccUnit:Row("", "B", "B")
'AccUnit:Row("A", "B", "AB")
Public Sub Concat_String_ConcatedString(ByVal ValueA As String, ValueB As String, Expected As String)
    Assert.That StringTools.Concat(ValueA, ValueB), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A")
Public Sub Trim_String_TrimedString_Both(ByVal checkValue As String, Expected As String)
    Assert.That StringTools.Trim(checkValue, TrimBoth), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A ")
Public Sub Trim_String_TrimedString_Start(ByVal checkValue As String, Expected As String)
    Assert.That StringTools.Trim(checkValue, TrimStart), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", " A")
Public Sub Trim_String_TrimedString_End(ByVal checkValue As String, Expected As String)
    Assert.That StringTools.Trim(checkValue, TrimEnd), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", 0,0, "")
'AccUnit:Row("abc", 1,1, "b")
'AccUnit:Row("abc", 2,1, "c")
'AccUnit:Row("abc", 3,0, "")
'AccUnit:Row("abc", 0,3, "abc")
Public Sub Substring_String_ExpectedSubstring(ByVal Value As String, ByVal startIndex As Long, ByVal length As Long, ByVal Expected As String)
    Assert.That StringTools.Substring(Value, startIndex, length), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "", 0, "")
'AccUnit:Row("A", "", 0, "A")
'AccUnit:Row("A", "", 1, "A")
'AccUnit:Row("", "A", 0, "A")
'AccUnit:Row(" ", "A", 1, " A")
'AccUnit:Row(" ", "A", 0, "A ")
'AccUnit:Row("AAAAA", "B", 0, "BAAAAA")
'AccUnit:Row("AAAAA", "BB", 1, "ABBAAAA")
'AccUnit:Row("AAAAA", "BB", 5, "AAAAABB")
Public Sub InsertAt_String_StringInsertedAtExpectedPos(ByVal Value As String, ByVal insertValue As String, ByVal pos As Long, ByVal Expected As String)
    Assert.AreEqual StringTools.InsertAt(Value, insertValue, pos), Expected
End Sub
