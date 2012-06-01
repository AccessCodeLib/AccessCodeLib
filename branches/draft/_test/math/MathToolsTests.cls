VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MathToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' AccUnit:TestClass
'--------------------------------------------------------------------
'<codelib>
'  <file>_test/math/MathToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>math/MathTools.bas</use>
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

'AccUnit:Row(1.025, 0, 1)
'AccUnit:Row(1.025, 1, 1.0)
'AccUnit:Row(1.025, 2, 1.03)
'AccUnit:Row(1.25,1, 1.3)
'AccUNit:row(2.3500, 1, 2.4)
'AccUnit:Row(2.5, 0, 3)
'AccUnit:Row(2.5, 1, 2.5)
'AccUnit:Row(2.5, 2, 2.5)
'AccUnit:Row(2.2499, 1, 2.2)
'AccUnit:Row(2.2500, 1, 2.3)
'AccUnit:Row(2.2501, 1, 2.3)
'AccUnit:Row(2.505, 2, 2.51)
'AccUnit:Row(1250, -2, 1300)
Public Sub Round_Round_AwayFromZero(Number As Double, numDiggitsAfterDecimal As Long, Expected As Double)
    Assert.That MathTools.Round(Number, numDiggitsAfterDecimal, AwayFromZero), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(1.025, 0, 1)
'AccUnit:Row(1.025, 1, 1.0)
'AccUnit:Row(1.025, 2, 1.02)
'AccUnit:Row(1.25,1, 1.2)
'AccUnit:Row(2.5, 0, 2)
'AccUnit:Row(2.5, 1, 2.5)
'AccUnit:Row(2.5, 2, 2.5)
'AccUnit:Row(2.2499, 1, 2.2)
'AccUnit:Row(2.2500, 1, 2.2)
'AccUnit:Row(2.2501, 1, 2.3)
'AccUNit:row(2.3500, 1, 2.4)
'AccUnit:Row(2.505, 2, 2.500)
Public Sub Round_Round_ToEven(Number As Double, numDiggitsAfterDecimal As Long, Expected As Double)
    Assert.That MathTools.Round(Number, numDiggitsAfterDecimal, ToEven), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(-171, "Null")
'AccUnit:Row(-1, "Null")
'AccUnit:Row(0, 1)
'AccUnit:Row(1, 1)
'AccUnit:Row(2, 2)
'AccUnit:Row(3, 6)
'AccUnit:Row(171, "Null")
Public Sub Fact_Calculate_Fact(Number As Variant, Expected As Variant)
    If "Null" <> Expected Then
        Assert.AreEqual MathTools.Fact(Number), Expected
    Else
        Assert.That MathTools.Fact(Number), Iz.vbNull
    End If
End Sub
