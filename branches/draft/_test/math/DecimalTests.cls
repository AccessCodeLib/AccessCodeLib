VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DecimalTests"
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

'AccUnit:Row(1, 2, 3)
Public Sub Value_Decimal_Addition(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant)
    Assert.That (Decimal.Value(a) + Decimal.Value(b)), Iz.EqualTo(Decimal.Value(c))
End Sub

'AccUnit:Row(1, 2, -1)
Public Sub Value_Decimal_Subtraktion(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant)
    Assert.That (Decimal.Value(a) - Decimal.Value(b)), Iz.EqualTo(Decimal.Value(c))
End Sub

'AccUnit:Row(2, 2, 4)
Public Sub Value_Decimal_Multiplication(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant)
    Assert.That (Decimal.Value(a) * Decimal.Value(b)), Iz.EqualTo(Decimal.Value(c))
End Sub

'AccUnit:Row(2, 2, 1)
Public Sub Value_Decimal_Division(ByVal a As Variant, ByVal b As Variant, ByVal c As Variant)
    Assert.That (Decimal.Value(a) / Decimal.Value(b)), Iz.EqualTo(Decimal.Value(c))
End Sub

'AccUnit:Row(2.5, 2)
Public Sub Value_Decimal_Round(ByVal a As Variant, ByVal b As Variant)
    Assert.That Decimal.Value(a).Round(0), Iz.EqualTo(Decimal.Value(b))
End Sub

'AccUnit:Row(1.025, "##,##0.00 €", 1.03, "1,03 €")
'AccUnit:Row(0.025, "0.00%", 1.03, "3,00%")
Public Sub Value_Decimal_RoundAndFormat(ByVal a As Variant, ByVal FormatString As String, ByVal b As Variant, ByVal c As String)
    Assert.That Decimal.Value(a).Round(2, AwayFromZero).ToString(FormatString), Iz.EqualTo(c)
End Sub

