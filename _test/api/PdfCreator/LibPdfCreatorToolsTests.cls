VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorToolsTests
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorTools.cls</use>
'  <use>_test/api/PdfCreator/TestReport.rpt</use>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

#Const DeleteTestFiles = True

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


Public Sub ConvertReportToPDF_Generate_PDF_ReportNotFound_ThrowsReportDoesNotExistExeption()
On Error GoTo Err_Test_

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("NotExistingReport", file, False, vbNullString)
    
    Assert.Fail "Expected: NotExistingReport-Error"

Err_Test_:
    Assert.That Err.Number, Iz.EqualTo(ReportDoesNotExist)
    Assert.Pass
End Sub

Public Sub ConvertReportToPDF_Generate_PDF_EmptyReportName_ThrowsReportNameIsEmptyExeption()
On Error GoTo Err_Test_

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF(vbNullString, file, False, vbNullString)
    
    Assert.Fail "Expected: NotExistingReport-Error"

Err_Test_:
    Assert.That Err.Number, Iz.EqualTo(ReportNameIsEmpty)
    Assert.Pass

End Sub

Public Sub ConvertReportToPDF_Generate_PDF()

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest_Generate_PDF.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString)
    
    Assert.That FileTools.FileExists(file), Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If

    Assert.That Result, Iz.True
       
End Sub

Public Sub ConvertReportToPDF_Generate_PDF_Metadata()

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest_Generate_PDF_Metadata.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString, _
                                                        True, "AccUnit", "A Title", "A Subject")
    
    Assert.That FileTools.FileExists(file), Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If

    Assert.That Result, Iz.True
       
End Sub

Public Sub ConvertReportToPDF_Generate_PDF_WithMaxSecurity()

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest_Generate_PDF_WithMaxSecurity.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString, _
                              True, , , , True, "AccLibPassword", , , , True, , , False, False)
    
    Assert.That FileTools.FileExists(file), Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If

    Assert.That Result, Iz.True
       
End Sub

Public Sub ConvertReportToPDF_Generate_PNG()
    
    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest.png"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString)
    
    Assert.That FileTools.FileExists(file), Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If
        
    Assert.That Result, Iz.True
       
End Sub

Public Sub ConvertReportToPDF_Generate_TIF()
    
    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest.tif"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString)
    
    Assert.That FileTools.FileExists(file), Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If
    
    Assert.That Result, Iz.True
       
End Sub

Public Sub ConvertReportToPDF_Generate_PDF_OverrideOutputFilefExists()

    Dim file As String
        file = CodeProject.Path & "\" & "AccUnitTest.pdf"
    
    Dim Result As Boolean
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString)
    
    Assert.That FileTools.FileExists(file), Iz.True
    Assert.That Result, Iz.True
    
    'File created, File exists, now override the File:
    
        Result = LibPdfCreatorTools.ConvertReportToPDF("TestReport", file, False, vbNullString)
    
    Assert.That FileTools.FileExists(file), Iz.True
    Assert.That Result, Iz.True
    
#If DeleteTestFiles Then
    Kill file
    Assert.That FileTools.FileExists(file), Iz.False
#End If
    
    
End Sub
