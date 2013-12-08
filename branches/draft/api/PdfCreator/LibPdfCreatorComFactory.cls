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
' Class Module: LibPdfCreatorComFactory
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Factory für COM-Objekte von PDF-Creator
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorComFactory.cls</file>
'  <use></use>
'  <test></test>
'  <license>_codelib/license.bas</license>
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
