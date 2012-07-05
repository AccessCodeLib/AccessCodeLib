VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime2Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'<codelib>
'  <file>_test\datetime\DateTime2Tests.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>

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

'AccUnit:Row("01.01.2012", "31.01.2012")
'AccUnit:Row("15.02.2012", "29.02.2012")
'AccUnit:Row("15.02.2011", "28.02.2011")
'AccUnit:Row("31.12.2011", "31.12.2011")
Public Sub NewValue_LastDayOfMonth(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfMonth(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2011", "01.01.2011")
'AccUnit:Row("15.02.2012", "01.02.2012")
'AccUnit:Row("15.02.2011", "01.02.2011")
'AccUnit:Row("31.12.2011", "01.12.2011")
Public Sub NewValue_FirstDayOfMonth(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfMonth(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2010", "01.01.2010")
'AccUnit:Row("15.02.2012", "01.01.2012")
'AccUnit:Row("15.02.2011", "01.01.2011")
'AccUnit:Row("31.12.2013", "01.01.2013")
Public Sub NewValue_FirstDayOfYear(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2010", "31.12.2010")
'AccUnit:Row("15.02.2012", "31.12.2012")
'AccUnit:Row("15.02.2011", "31.12.2011")
'AccUnit:Row("31.12.2013", "31.12.2013")
Public Sub NewValue_LastDayOfYear(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfYear(), Iz.EqualTo(Expected)
End Sub
