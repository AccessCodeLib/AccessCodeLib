VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Base64Tests"
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

Public Sub EncodeAndDecode_String_FromAndToUnicode()
    Assert.AreEqual Base64.Decode(Base64.Encode("foo")), "foo"
    Assert.AreEqual Base64.Decode(Base64.Encode("foo", FromUnicode), ToUnicode), "foo"
    Assert.AreEqual Base64.Decode(Base64.Encode("foo", ToUnicode), FromUnicode), "foo"
End Sub


'AccUnit:Row("", "")
'AccUnit:Row("A", "QQA=")
'AccUnit:Row("a", "YQA=")
'AccUnit:Row("test", "dABlAHMAdAA=")
Public Sub Encode_String_Base64EncodedString_None(ByVal Value As String, ByVal Expected As String)
    Assert.AreEqual Base64.Encode(Value, None), Expected
End Sub

'AccUnit:Row("", "")
'AccUnit:Row("A", "QQ==")
'AccUnit:Row("a", "YQ==")
'AccUnit:Row("test", "dGVzdA==")
Public Sub Encode_String_Base64EncodedString(ByVal Value As String, ByVal Expected As String)
    Assert.AreEqual Base64.Encode(Value, FromUnicode), Expected
End Sub

Public Sub Encode_MiscTests()
    Assert.AreEqual Base64.Encode(Chr$(0)), "AAA="
    Assert.AreEqual Base64.Encode(Chr$(9)), "CQA="
    Assert.AreEqual Base64.Encode(Chr$(13)), "DQA="
    
    Assert.AreEqual Base64.Encode(" ", FromUnicode), "IA=="
End Sub

'AccUnit:Row("", "")
'AccUnit:Row("QQA=", "A")
'AccUnit:Row("YQA=", "a")
'AccUnit:Row("dABlAHMAdAA=", "test")
Public Sub Decode_String_Base64DecodedString_None(ByVal Value As String, ByVal Expected As String)
    Assert.AreEqual Base64.Decode(Value, None), Expected
End Sub

'AccUnit:Row("", "")
'AccUnit:Row("IA==", " ")
'AccUnit:Row("QQ==", "A")
'AccUnit:Row("YQ==", "a")
'AccUnit:Row("dGVzdA==", "test")
Public Sub Decode_String_Base64DecodedString_ToUnicode(ByVal Value As String, ByVal Expected As String)
    Assert.AreEqual Base64.Decode(Value, ToUnicode), Expected
End Sub


Public Sub Decode_MiscTests()
    Assert.AreEqual Base64.Decode("AAA="), Chr$(0)
    Assert.AreEqual Base64.Decode("CQA="), Chr$(9)
    Assert.AreEqual Base64.Decode("DQA="), Chr$(13)
End Sub
