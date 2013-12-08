VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorEnumsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorEnums
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorEnumsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorEnums.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------

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


Public Sub ErrorNummers_EnumCheck()

    Assert.That LibPdfCreatorErrorNumer.WarteSchlangeNichtLeer, Iz.EqualTo(6000)
    Assert.That LibPdfCreatorErrorNumer.UnexpectedPdfCreatorError, Iz.EqualTo(6001)
    Assert.That LibPdfCreatorErrorNumer.ZielDateiSchreibgeschützt, Iz.EqualTo(6002)
    Assert.That LibPdfCreatorErrorNumer.NotSupportedPdfCreatorFileFormat, Iz.EqualTo(6003)
    Assert.That LibPdfCreatorErrorNumer.ReportDoesNotExist, Iz.EqualTo(6004)
    Assert.That LibPdfCreatorErrorNumer.ReportNameIsEmpty, Iz.EqualTo(6005)

End Sub
