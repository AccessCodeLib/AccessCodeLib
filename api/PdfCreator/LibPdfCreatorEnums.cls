VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorEnums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorEnums
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Enumerationen
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorEnums.cls</file>
'  <test>_test/api/PdfCreator/LibPdfCreatorEnumsTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

'Ausgabeformat
Public Enum LibPdfCreatorFileFormat
    LibPdfCreatorFileFormat_PDF = 0
    LibPdfCreatorFileFormat_PNG = 1
    LibPdfCreatorFileFormat_JPG = 2
    LibPdfCreatorFileFormat_BMP = 3
    LibPdfCreatorFileFormat_PCX = 4
    LibPdfCreatorFileFormat_TIF = 5
    LibPdfCreatorFileFormat_PS = 6
    LibPdfCreatorFileFormat_EPS = 7
    LibPdfCreatorFileFormat_TXT = 8
    LibPdfCreatorFileFormat_PDF_A = 9
    LibPdfCreatorFileFormat_PDF_X = 10
    LibPdfCreatorFileFormat_PSD = 11
    LibPdfCreatorFileFormat_PCL = 12
    LibPdfCreatorFileFormat_RAW = 13
End Enum

Public Enum LibPdfCreatorOptionFlag
    No = 0
    Yes = 1
End Enum

Public Enum LibPdfCreatorErrorNumer
    WarteSchlangeNichtLeer = 6000
    UnexpectedPdfCreatorError
    ZielDateiSchreibgeschützt
    NotSupportedPdfCreatorFileFormat
    ReportDoesNotExist
    ReportNameIsEmpty
End Enum
