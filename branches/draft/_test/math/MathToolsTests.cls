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
Public Sub Round_Round_AwayFromZero(number As Double, numDiggitsAfterDecimal As Long, Expected As Double)
    Assert.That MathTools.Round(number, numDiggitsAfterDecimal, AwayFromZero), Iz.EqualTo(Expected)
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
Public Sub Round_Round_ToEven(number As Double, numDiggitsAfterDecimal As Long, Expected As Double)
    Assert.That MathTools.Round(number, numDiggitsAfterDecimal, ToEven), Iz.EqualTo(Expected)
End Sub

'/*
' * TODO: Better Tests
' */

'AccUnit:Row(0.000000000000000000000000000000000000009, 0.000000000000000000000000000000000000009, 0.000000000000000000000000000000000000018)
Public Sub DecimalAddition_TestInputValues(ByVal valueA As Variant, ByVal valueB As Variant, Expected As Variant)
    Assert.That MathTools.DecimalAddition(valueA, valueB), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(0.000000000000000000000000000000000000009, 0.000000000000000000000000000000000000009, 0)
Public Sub DecimalSubtraction_TestInputValues(ByVal valueA As Variant, ByVal valueB As Variant, Expected As Variant)
    Assert.That MathTools.DecimalSubtraction(valueA, valueB), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(0.000000000000000000000000000000000000009, 0.000000000000000000000000000000000000081, 0)
Public Sub DecimalMultiplication_TestInputValues(ByVal valueA As Variant, ByVal valueB As Variant, Expected As Variant)
    Assert.That MathTools.DecimalMultiplication(valueA, valueB), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(1, 2, 0.5)
'AccUnit:Row(0.0000000000000000000000000009, 0.0000000000000000000000000001, 9)
Public Sub DecimalDivision_TestInputValues(ByVal valueA As Variant, ByVal valueB As Variant, Expected As Variant)
    Assert.That MathTools.DecimalDivision(valueA, valueB), Iz.EqualTo(Expected)
End Sub
