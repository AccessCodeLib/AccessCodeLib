VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorInstanceTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorInstanceTests
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorInstanceTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorInstance.cls</use>
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


Public Sub CreatorInstance_MultipleInstanceOfPdfCreatortError_HasErrorNr_2()
    
    Dim CreatorInstance As New LibPdfCreatorInstance
        CreatorInstance.cStart ("/NoProcessingAtStartup")
        CreatorInstance.cStart ("/NoProcessingAtStartup")
        
        Assert.That CreatorInstance.cError.Number, Iz.EqualTo(2)
    
        CreatorInstance.cErrorClear
    
        Assert.That CreatorInstance.cError.Number, Iz.EqualTo(0)
        
        CreatorInstance.cClose
        CreatorInstance.cStart ("/NoProcessingAtStartup")
        
        Assert.That CreatorInstance.cError.Number, Iz.EqualTo(0)
    
        CreatorInstance.cClose
    Set CreatorInstance = Nothing
    
    
End Sub
