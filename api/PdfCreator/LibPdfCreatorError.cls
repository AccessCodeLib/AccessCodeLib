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
' Class Module: LibPdfCreatorError
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Fehler-Objekt für PDF Creator
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorError.cls</file>
'  <license>_codelib/license.bas</license>
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

Public Property Let Number(ByVal value As Long)
    m_Number = value
End Property

Public Property Get Description() As String
    Description = m_Description
End Property

Public Property Let Description(ByVal value As String)
    m_Description = value
End Property
