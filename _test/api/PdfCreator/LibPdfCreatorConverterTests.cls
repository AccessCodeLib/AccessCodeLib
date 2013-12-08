VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorConverterTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
'\short     AccUnit-Testklasse für LibPdfCreatorConverter
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/api/PdfCreator/LibPdfCreatorConverterTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>api/PdfCreator/LibPdfCreatorConverter.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'AccUnit:TestClass

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub


Public Sub ConvertPdfCreatorOption_Converts_ToLibPdfCreatorOptions()
    
#If UseEarlyBinding Then
    Dim instance As PDFCreator.clsPDFCreator
    Dim pdfOptions As PDFCreator.clsPDFCreatorOptions
#Else
    Dim instance As Object
    Dim pdfOptions As Object
#End If

    Set instance = LibPdfCreatorComFactory.clsPDFCreator
    Set pdfOptions = instance.cStandardOptions
        pdfOptions.AutosaveFilename = "UnitTest1234.pdf"
    
    Dim libOptions As LibPdfCreatorOptions
    Set libOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(pdfOptions)
    
    Assert.That libOptions.AutosaveFilename, Iz.EqualTo("UnitTest1234.pdf")
    
    Set pdfOptions = Nothing
    Set libOptions = Nothing
    Set instance = Nothing
    
End Sub

Public Sub ConvertPdfCreatorOption_UnConverts_PdfCreatorOptionsComObject()
    
    Dim instance As New LibPdfCreatorInstance
    
    Dim pdfOptions As LibPdfCreatorOptions
    Set pdfOptions = instance.cStandardOptions
        pdfOptions.AutosaveFilename = "UnitTest1234.pdf"
    
#If UseEarlyBinding Then
    Dim libOptions As PDFCreator.clsPDFCreatorOptions
#Else
    Dim libOptions As Object
#End If

    Set libOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToCOM(pdfOptions)
    
    Assert.That libOptions.AutosaveFilename, Iz.EqualTo("UnitTest1234.pdf")
    
    Set pdfOptions = Nothing
    Set libOptions = Nothing
    Set instance = Nothing
    
End Sub
