VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2Tests"
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


Public Sub Value_Init_DefaultValue()
    
    Dim a As New String2
    
    Assert.That a.Value, Iz.EqualTo(vbNullString)
    
    Set a = Nothing
End Sub

'AccUnit:Row("abc", False)
'AccUnit:Row("", True)
'AccUnit:Row(" ", False).Name = "1 Space"
'AccUnit:Row(Null, True)
'AccUnit:Row(1, False)
Public Sub IsNullOrEmpty_WithoutIgnoreSpaceParameter_ReturnsTrueOrFalse(ByVal vValue As Variant, ByVal Expected As Boolean)
   
   With New String2
        Assert.AreEqual Expected, .NewValue(vValue).IsNullOrEmpty
   End With
   
   Assert.AreEqual Expected, String2.NewValue(vValue).IsNullOrEmpty
   
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
   
   With New String2
        Assert.AreEqual Expected, .NewValue(vValue).IsNullOrEmpty(IgnoreSpaces)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(vValue).IsNullOrEmpty(IgnoreSpaces)
   
End Sub


'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "__Test")
Public Sub PadLeft_String_LeftPadString(ByVal Value As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)

   With New String2
        Assert.AreEqual Expected, .NewValue(Value).PadLeft(totalWidth, padChar)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).PadLeft(totalWidth, padChar)
    
End Sub

'AccUnit:Row("Test", 4, "Test")
'AccUnit:Row("Test", 6, "  Test")
Public Sub PadLeft_String_LeftPadStringDefaultParam(ByVal Value As String, ByVal totalWidth As Integer, ByVal Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(Value).PadLeft(totalWidth)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).PadLeft(totalWidth)
   
End Sub

'AccUnit:Row("Test", 4, "_", "Test")
'AccUnit:Row("Test", 6, "_", "Test__")
Public Sub PadRight_String_RightPadString(ByVal Value As String, ByVal totalWidth As Integer, ByVal padChar As String, ByVal Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(Value).PadRight(totalWidth, padChar)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).PadRight(totalWidth, padChar)
   
End Sub

'AccUnit:Row("Test", 4, "Test")
'AccUnit:Row("Test", 6, "Test  ")
Public Sub PadRight_String_RightPadStringDefaultParam(ByVal Value As String, ByVal totalWidth As Integer, ByVal Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(Value).PadRight(totalWidth)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).PadRight(totalWidth)
    
End Sub

'AccUnit:Row("abc", "b", True)
'AccUnit:Row("abc", "B", True)
'AccUnit:Row("abc", "", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", False)
'AccUnit:Row("", "A", False)
Public Sub Contains_String_ContainsSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).Contains(searchValue)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).Contains(searchValue)
   
End Sub

'AccUnit:Row("abcd", "cd", True)
'AccUnit:Row("abcd", "CD", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_EndsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).EndsWith(searchValue)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).EndsWith(searchValue)
   
End Sub

'AccUnit:Row("abcd", "ab", True)
'AccUnit:Row("abcd", "AB", True)
'AccUnit:Row("abc", "ABC", True)
'AccUnit:Row("abc", "abc", True)
'AccUnit:Row("", "", True)
'AccUnit:Row("", "A", False)
Public Sub EndsWith_String_StartsWithSearchString(ByVal checkValue As String, ByVal searchValue As String, Expected As Boolean)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).StartsWith(searchValue)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).StartsWith(searchValue)
   
End Sub

'AccUnit:Row("", 0)
'AccUnit:Row("1234567890", 10)
Public Sub Lenght_String_HasLenght(ByVal checkValue As String, Expected As Integer)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).Lenght
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).Lenght
   
End Sub

'AccUnit:Row("", "", "")
'AccUnit:Row("A", "", "A")
'AccUnit:Row("", "B", "B")
'AccUnit:Row("A", "B", "AB")
Public Sub Concat_String_AppendString(ByVal ValueA As String, ValueB As String, Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(ValueA).Append(ValueB)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(ValueA).Append(ValueB)
   
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A")
Public Sub Trim_String_TrimedString_Both(ByVal checkValue As String, Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).Trim(TrimBoth)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).Trim(TrimBoth)
   
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", "A ")
Public Sub Trim_String_TrimedString_Start(ByVal checkValue As String, Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).Trim(TrimStart)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).Trim(TrimStart)
   
End Sub

'AccUnit:Row("", "")
'AccUnit:Row(" ", "")
'AccUnit:Row("  ", "")
'AccUnit:Row(" A ", " A")
Public Sub Trim_String_TrimedString_End(ByVal checkValue As String, Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(checkValue).Trim(TrimEnd)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(checkValue).Trim(TrimEnd)
   
End Sub

'AccUnit:Row("", 0,0, "")
'AccUnit:Row("abc", 1,1, "b")
'AccUnit:Row("abc", 2,1, "c")
'AccUnit:Row("abc", 3,0, "")
'AccUnit:Row("abc", 0,3, "abc")
Public Sub Substring_String_ExpectedSubstring(ByVal Value As String, ByVal startIndex As Long, ByVal length As Long, ByVal Expected As String)
    
   With New String2
        Assert.AreEqual Expected, .NewValue(Value).Substring(startIndex, length)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).Substring(startIndex, length)
   
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
    
   With New String2
        Assert.AreEqual Expected, .NewValue(Value).InsertAt(insertValue, pos)
   End With
   
   Assert.AreEqual Expected, String2.NewValue(Value).InsertAt(insertValue, pos)
   
End Sub
