VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Decimal2Tests"
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


Public Sub value_Default_IsZero()
    
    Dim a As New Decimal2
        a = 0
        
    Dim b As New Decimal2 '0 by Default
        
    Assert.AreEqual a, b
        
    Set a = Nothing
    Set b = Nothing
    
End Sub

Public Sub value_Equal()
    
    Dim a As New Decimal2
    Dim b As New Decimal2
    
    a = 1
    b = 1
    
    Assert.AreEqual a, b
        
    Set a = Nothing
    Set b = Nothing
    
End Sub

Public Sub value_Simple_Add()
    
    Dim a As New Decimal2
    Dim b As New Decimal2
    Dim c As New Decimal2
    
    a = 1
    b = 1
    c = 2
    
    Assert.AreEqual a + b, c
    
    Set a = Nothing
    Set b = Nothing
    Set c = Nothing
    
End Sub

Public Sub Round_CallMethod_UsesMathToolsRoundFunction()

    Dim a As New Decimal2
    Dim b As New Decimal2
    Dim c As New Decimal2
    
    a = 1.025
    b = 1.02
    c = 1.03
    
    Assert.AreEqual a.Round(2, ToEven), b
    Assert.AreEqual a.Round(2, AwayFromZero), c
    
    Set a = Nothing
    Set b = Nothing
    Set c = Nothing
    
End Sub

Public Sub ToString_ReturnsString()

    Dim a As New Decimal2
        a = 1.01

    Assert.That a.ToString(), Iz.EqualTo("1,01")
    Assert.That a.ToString("##,##0.00 €"), Iz.EqualTo("1,01 €")
    
    Set a = Nothing
    
End Sub


Public Sub ToString2_GetValue_AsString2()
    
    Dim a As New Decimal2
        a = 1.01
    Dim s As New String2
        s = "1,01"
        
    Assert.AreEqual a.ToString2(), s
    
    Set a = Nothing
    Set s = Nothing
End Sub

