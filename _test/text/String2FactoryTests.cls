VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2FactoryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'AccUnit:TestClass

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


Public Sub Value_Get_String2DataType()
    Dim a As New String2
        a = "Test"
        
    Dim b As String2
    Set b = String2Factory.Value("Test")
        
    Assert.That a, Iz.EqualTo(b)
    
    Set a = Nothing
    Set b = Nothing
End Sub

'AccUnit:Row("abc", False)
'AccUnit:Row("", True)
'AccUnit:Row(" ", False).Name = "1 Space"
'AccUnit:Row(Null, True)
'AccUnit:Row(1, False)
Public Sub Value_IsNullOrEmpty_WithoutIgnoreSpaceParameter_ReturnsTrueOrFalse(ByVal vValue As Variant, ByVal Expected As Boolean)
   Dim Actual As Boolean
   Actual = String2Factory.Value(vValue).IsNullOrEmpty()
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
   Actual = String2Factory.Value(vValue).IsNullOrEmpty(IgnoreSpaces)
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
   Actual = String2Factory.Value(Expression).ToString(FormatString)
   Assert.AreEqual Expected, Actual
End Sub

'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "__Test")
Public Sub PadLeft_String_LeftPadString(ByVal testValue As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)
    Assert.AreEqual String2Factory.Value(testValue).PadLeft(totalWidth, padChar).ToString(), Expected
End Sub

'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "Test__")
Public Sub PadRight_String_RightPadString(ByVal testValue As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)
    Assert.AreEqual String2Factory.Value(testValue).PadRight(totalWidth, padChar).ToString(), Expected
End Sub

'AccUnit:Row("Test", 4, "Test")
'AccUnit:Row("Test", 6, "Test  ")
Public Sub PadRight_String_RightPadStringDefaultParam(ByVal testValue As String, ByVal totalWidth As Integer, ByVal Expected As String)
    Assert.AreEqual String2Factory.Value(testValue).PadRight(totalWidth).ToString(), Expected
End Sub

'AccUnit:Row("abc", "b", True)
'AccUnit:Row("abc", "B", True)
'AccUnit:Row("abc", "", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", False)
'AccUnit:Row("", "A", False)
Public Sub Contains_String_ContainsSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That String2Factory.Value(checkValue).Contains(searchValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("abcd", "cd", True)
'AccUnit:Row("abcd", "CD", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_EndsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That String2Factory.Value(checkValue).EndsWith(searchValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("abcd", "ab", True)
'AccUnit:Row("abcd", "AB", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_StartsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    Assert.That String2Factory.Value(checkValue).StartsWith(searchValue), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", 0)
'AccUnit:Row("1234567890", 10)
Public Sub Lenght_String_HasLenght(ByVal checkValue As String, Expected As Integer)
    Assert.That String2Factory.Value(checkValue).Lenght(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "", "")
'AccUnit:Row("A", "", "A")
'AccUnit:Row("", "B", "B")
'AccUnit:Row("A", "B", "AB")
Public Sub Concat_String_ConcatedString(ByVal ValueA As String, ValueB As String, Expected As String)
    Assert.That String2Factory.Value(ValueA).Concat(ValueB).ToString(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A")
Public Sub Trim_String_TrimedString_Both(ByVal checkValue As String, Expected As String)
    Assert.That String2Factory.Value(checkValue).Trim(TrimBoth).ToString(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A ")
Public Sub Trim_String_TrimedString_Start(ByVal checkValue As String, Expected As String)
    Assert.That String2Factory.Value(checkValue).Trim(TrimStart).ToString(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", " A")
Public Sub Trim_String_TrimedString_End(ByVal checkValue As String, Expected As String)
    Assert.That String2Factory.Value(checkValue).Trim(TrimEnd).ToString(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("", 0,0, "")
'AccUnit:Row("abc", 1,1, "b")
'AccUnit:Row("abc", 2,1, "c")
'AccUnit:Row("abc", 3,0, "")
'AccUnit:Row("abc", 0,3, "abc")
Public Sub Substring_String_ExpectedSubstring(ByVal Value As String, ByVal startIndex As Long, ByVal length As Long, ByVal Expected As String)
    Assert.That String2Factory.Value(Value).Substring(startIndex, length), Iz.EqualTo(Expected)
End Sub
