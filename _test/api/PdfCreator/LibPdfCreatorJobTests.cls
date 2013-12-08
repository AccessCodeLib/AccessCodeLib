VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorJobTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorJobTests
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorJobTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorJob.cls</use>
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


Public Sub CreatorOptions_Option_IsStandardOption()
    
    'Test ob job.CreatorOptions tatsächlich die StandardOptions liefert
    
    Dim job As LibPdfCreatorJob
    Set job = New LibPdfCreatorJob
    
    Assert.That job.CreatorOptions.AutosaveDirectory, Iz.EqualTo("<MyFiles>\")
    Assert.That job.CreatorOptions.AutosaveFilename, Iz.EqualTo("<DateTime>")
    
    Set job = Nothing
    
End Sub
