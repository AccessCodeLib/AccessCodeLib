VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorFactoryTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorFactoryTests
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorFactoryTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorFactory.cls</use>
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

Public Sub PdfCreatorInstance__NoError_After_Start_Multiple_Instances()

    Dim instance As LibPdfCreatorInstance
    Set instance = LibPdfCreatorFactory.PdfCreatorInstance()

    Assert.That instance.cError.Number, Iz.EqualTo(0)

    '2. instance, Error = 2 should be catched by Factory
    Set instance = LibPdfCreatorFactory.PdfCreatorInstance()

    Assert.That instance.cError.Number, Iz.EqualTo(0)

        instance.cClose
    Set instance = Nothing
End Sub


Public Sub PdfCreatorInstance_Correct_Type()

    Dim test As LibPdfCreatorInstance
    Set test = LibPdfCreatorFactory.PdfCreatorInstance()
    
    Dim b As Boolean: b = False
    
    If TypeOf test Is LibPdfCreatorInstance Then _
        b = True
    
    Assert.That b, Iz.True
        test.cClose
    Set test = Nothing

End Sub
