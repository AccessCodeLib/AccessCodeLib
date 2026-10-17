VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: api.pdfcreator.LibPdfCreatorError
'---------------------------------------------------------------------------------------
'
' Error object for PDFCreator
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/pdfcreator/LibPdfCreatorError.cls</file>
'  <license>_codelib/license.bas</license>
'  <description>Error object of the PdfCreator package (number, description)</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Description As String
Private m_Number As Long

Private Sub Class_Initialize()
    m_Number = 0
    m_Description = vbNullString
End Sub

Public Property Get Number() As Long
    Number = m_Number
End Property

Public Property Let Number(ByVal Value As Long)
    m_Number = Value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal Value As String)
    m_Description = Value
End Property
