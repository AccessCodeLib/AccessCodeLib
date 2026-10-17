VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorComFactory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.pdfcreator.LibPdfCreatorComFactory
'---------------------------------------------------------------------------------------
'
' Factory for PDFCreator COM objects
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/pdfcreator/LibPdfCreatorComFactory.cls</file>
'  <license>_codelib/license.bas</license>
'  <description>Factory for the COM objects of PDFCreator 1.x (clsPDFCreator, clsPDFCreatorOptions, clsPDFCreatorError)</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

#If UseEarlyBinding Then

    Public Property Get clsPDFCreator() As PDFCreator.clsPDFCreator
        Set clsPDFCreator = New PDFCreator.clsPDFCreator
    End Property
    
    Public Property Get clsPDFCreatorOptions() As PDFCreator.clsPDFCreatorOptions
        Set clsPDFCreatorOptions = New PDFCreator.clsPDFCreatorOptions
    End Property
    
    Public Property Get clsPDFCreatorError() As PDFCreator.clsPDFCreatorError
        Set clsPDFCreatorError = New PDFCreator.clsPDFCreatorError
    End Property

#Else
    
    Public Property Get clsPDFCreator() As Object
        Set clsPDFCreator = VBA.Interaction.CreateObject("PDFCreator.clsPDFCreator")
    End Property
    
    Public Property Get clsPDFCreatorOptions() As Object
        Set clsPDFCreatorOptions = VBA.Interaction.CreateObject("PDFCreator.clsPDFCreatorOptions")
    End Property
    
    Public Property Get clsPDFCreatorError() As Object
        Set clsPDFCreatorError = VBA.Interaction.CreateObject("PDFCreator.clsPDFCreatorError")
    End Property

#End If
