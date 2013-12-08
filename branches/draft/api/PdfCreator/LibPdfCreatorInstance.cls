VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorInstance"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorInstance
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Repräsentiert die Instanz eines PDF-Creator Prozesses
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorInstance.cls</file>
'  <use>api/PdfCreator/LibPdfCreatorOptions</use>
'  <test>_test/api/PdfCreator/LibPdfCreatorInstanceTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

#If UseEarlyBinding Then
    Private m_clsPdfCreator As PDFCreator.clsPDFCreator
#Else
    Private m_clsPdfCreator As Object
#End If

Private Sub Class_Initialize()
    Set m_clsPdfCreator = LibPdfCreatorComFactory.clsPDFCreator
End Sub

Private Sub Class_Terminate()
    Set m_clsPdfCreator = Nothing
End Sub

'Functions

Public Function cAddPDFCreatorPrinter(ByVal PrinterName As String, Optional ByVal ProfileName As String = vbNullString) As Boolean
    cAddPDFCreatorPrinter = m_clsPdfCreator.cAddPDFCreatorPrinter(PrinterName, ProfileName)
End Function

Public Function cAddProfile(ByVal ProfileName As String, Optional ByVal Options1 As Object) As Boolean
    cAddProfile = m_clsPdfCreator.cAddProfile(ProfileName, Options1)
End Function

Public Function cClose() As Boolean
    cClose = m_clsPdfCreator.cClose
End Function

Public Function cDeletePDFCreatorPrinter(ByVal PrinterName As String) As Boolean
    cDeletePDFCreatorPrinter = m_clsPdfCreator.cDeletePDFCreatorPrinter(PrinterName)
End Function

Public Function cDeleteProfile(ByVal ProfileName As String) As Boolean
    cDeleteProfile = m_clsPdfCreator.cDeleteProfile(ProfileName)
End Function

Public Function cGetLogfile() As String
    cGetLogfile = m_clsPdfCreator.cGetLogfile
End Function

Public Function cGetPDFCreatorPrinters() As Collection
    Set cGetPDFCreatorPrinters = m_clsPdfCreator.cGetPDFCreatorPrinters
End Function

Public Function cGetPrinterProfiles() As Collection
    Set cGetPrinterProfiles = m_clsPdfCreator.cGetPrinterProfiles
End Function

Public Function cGetProfileNames() As Collection
    Set cGetProfileNames = m_clsPdfCreator.cGetProfileNames
End Function

Public Function cGhostscriptRun(Arguments() As Variant) As Boolean
    cGhostscriptRun = m_clsPdfCreator.cGhostscriptRun(Arguments())
End Function

Public Function cIsAdministrator() As Boolean
    cIsAdministrator = m_clsPdfCreator.cIsAdministrator
End Function

Public Function cIsPrintable(ByVal filename As String) As Boolean
    cIsPrintable = m_clsPdfCreator.cIsPrintable(filename)
End Function

Public Function cPrinterIsInstalled(ByVal PrinterName As String) As Boolean
    cPrinterIsInstalled = m_clsPdfCreator.cPrinterIsInstalled(PrinterName)
End Function

Public Function cProfileExists(ByVal ProfileName As String) As Boolean
    cProfileExists = m_clsPdfCreator.cProfileExists(ProfileName)
End Function

Public Function cReadOptions(Optional ByVal ProfileName As String = vbNullString) As LibPdfCreatorOptions
    cReadOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(m_clsPdfCreator.cReadOptions(ProfileName))
End Function

Public Function cReadOptionsFromFile(ByVal INIFilename As String) As LibPdfCreatorOptions
    cReadOptionsFromFile = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(m_clsPdfCreator.cReadOptionsFromFile(INIFilename))
End Function

Public Function cRenameProfile(ByVal OldProfileName As String, ByVal NewProfileName As String) As Boolean
    cRenameProfile = m_clsPdfCreator.cRenameProfile(OldProfileName, NewProfileName)
End Function

Public Function cStart(Optional ByVal Params As String, Optional ByVal ForceInitialize As Boolean = False) As Boolean
    cStart = m_clsPdfCreator.cStart(Params, ForceInitialize)
End Function

'Subroutines

Public Sub cAddPrintjob(ByVal filename As String) ' Supported file types: Postscript, JPEG, BMP
    m_clsPdfCreator.cAddPrintjob filename
End Sub

Public Sub cClearCache()
    m_clsPdfCreator.cClearCache
End Sub

Public Sub cClearLogfile()
    m_clsPdfCreator.cClearLogfile
End Sub

Public Sub cCombineAll()
    m_clsPdfCreator.cCombineAll
End Sub

' Allowed values for 'SubFormat': "PDF/A-1b" for "PDF/A-1b"files; "PDF/X" for "PDF/X"-files
Public Sub cConvertFile(ByVal InputFilename As String, _
                        ByVal OutputFilename As String, _
               Optional ByVal SubFormat As String = vbNullString)

    m_clsPdfCreator.cConvertFile InputFilename, OutputFilename, SubFormat
End Sub

' This function is deprecated! Use "cConvertFile()" instead.
Public Sub cConvertPostscriptfile(ByVal InputFilename As String, ByVal OutputFilename As String)
    m_clsPdfCreator.cConvertPostscriptfile InputFilename, OutputFilename
End Sub

Public Sub cDeletePrintjob(ByVal JobNumber As Long)
    m_clsPdfCreator.cDeletePrintjob JobNumber
End Sub

Public Sub cErrorClear()
    m_clsPdfCreator.cErrorClear
End Sub

Public Sub cMovePrintjobBottom(ByVal JobNumber As Long)
    m_clsPdfCreator.cMovePrintjobBottom JobNumber
End Sub

Public Sub cMovePrintjobDown(ByVal JobNumber As Long)
    m_clsPdfCreator.cMovePrintjobDown JobNumber
End Sub

Public Sub cMovePrintjobTop(ByVal JobNumber As Long)
    m_clsPdfCreator.cMovePrintjobTop JobNumber
End Sub

Public Sub cMovePrintjobUp(ByVal JobNumber As Long)
    m_clsPdfCreator.cMovePrintjobUp JobNumber
End Sub

Public Sub cPrintFile(ByVal filename As String)
    m_clsPdfCreator.cPrintFile filename
End Sub

Public Sub cPrintPDFCreatorTestpage()
    m_clsPdfCreator.cPrintPDFCreatorTestpage
End Sub

Public Sub cPrintPrinterTestpage(Optional ByVal PrinterName As String)
    m_clsPdfCreator.cPrintPrinterTestpage PrinterName
End Sub

Public Sub cPrintURL(ByVal URL As String, Optional ByVal TimeBetweenLoadAndPrint As Long = 0)
    m_clsPdfCreator.cPrintURL URL, TimeBetweenLoadAndPrint
End Sub

' Options1 must be a clsPDFCreatorOptions type
Public Sub cSaveOptions(Optional ByVal Options1 As Variant, Optional ByVal ProfileName As String = "")
    m_clsPdfCreator.cSaveOptions Options1, ProfileName
End Sub

' Options1 must be a clsPDFCreatorOptions type
Public Sub cSaveOptionsToFile(ByVal INIFilename As String, Optional ByVal Options1 As Variant)
    m_clsPdfCreator.cSaveOptionsToFile INIFilename, Options1
End Sub

Public Sub cSendMail(ByVal OutputFilename As String, ByVal Recipients As String)
    m_clsPdfCreator.cSendMail OutputFilename, Recipients
End Sub

Public Sub cShowLogfileDialog(ByVal value As Boolean)
    m_clsPdfCreator.cShowLogfileDialog value
End Sub

Public Sub cShowOptionsDialog(ByVal value As Boolean)
    m_clsPdfCreator.cShowOptionsDialog value
End Sub

Public Sub cTestEvent(ByVal EventName As String)
    m_clsPdfCreator.cTestEvent EventName
End Sub

Public Sub cWriteToLogfile(ByVal LogStr As String)
    m_clsPdfCreator.cWriteToLogfile LogStr
End Sub

'Properties

Public Property Get cCountOfPrintjobs() As Long
    cCountOfPrintjobs = m_clsPdfCreator.cCountOfPrintjobs
End Property

Public Property Get cDefaultPrinter() As String
    cDefaultPrinter = m_clsPdfCreator.cDefaultPrinter
End Property

Public Property Let cDefaultPrinter(ByVal PrinterName As String)
    m_clsPdfCreator.cDefaultPrinter = PrinterName
End Property

Public Property Get cError() As LibPdfCreatorError
    Set cError = LibPdfCreatorConverter.ConvertPdfCreatorError_ToLib(m_clsPdfCreator.cError)
End Property

Public Property Get cErrorDetail(ByVal PropertyName As String) As Variant
    cErrorDetail = m_clsPdfCreator.cErrorDetail(PropertyName)
End Property

Public Property Get cGhostscriptVersion() As String
    cGhostscriptVersion = m_clsPdfCreator.cGhostscriptVersion
End Property

Public Property Get cInstalledAsServer() As Boolean
    cInstalledAsServer = m_clsPdfCreator.cInstalledAsServer
End Property

Public Property Get cInstanceCounter() As Long
    cInstanceCounter = m_clsPdfCreator.cInstanceCounter
End Property

Public Property Get cIsClosed() As Boolean
    cIsClosed = m_clsPdfCreator.cIsClosed
End Property

Public Property Get cIsConverted() As Boolean
    cIsConverted = m_clsPdfCreator.cIsConverted
End Property

Public Property Let cIsConverted(ByVal value As Boolean)
    m_clsPdfCreator.cIsConverted = value
End Property

Public Property Get cIsLogfileDialogDisplayed() As Boolean
    cIsLogfileDialogDisplayed = m_clsPdfCreator.cIsLogfileDialogDisplayed
End Property

Public Property Get cIsOptionsDialogDisplayed() As Boolean
    cIsOptionsDialogDisplayed = m_clsPdfCreator.cIsOptionsDialogDisplayed
End Property

Public Property Get cOption(ByVal PropertyName As String) As Variant
    cOption = m_clsPdfCreator.cOption(PropertyName)
End Property

Public Property Let cOption(ByVal PropertyName As String, ByVal value As Variant)
    m_clsPdfCreator.cOption(PropertyName) = value
End Property

Public Property Get cOptionProfile(ByVal ProfileName As String, ByVal PropertyName As String) As Variant
    cOptionProfile = m_clsPdfCreator.cOptionProfile(ProfileName, PropertyName)
End Property

Public Property Let cOptionProfile(ByVal ProfileName As String, ByVal PropertyName As String, ByVal value As Variant)
    m_clsPdfCreator.cOptionProfile(ProfileName, PropertyName) = value
End Property

Public Property Get cOptions() As LibPdfCreatorOptions
    Set cOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(m_clsPdfCreator.cOptions)
End Property

Public Property Set cOptions(ByVal value As LibPdfCreatorOptions)
    Set m_clsPdfCreator.cOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToCOM(value)
End Property

Public Property Get cOptionsNames() As Collection
    Set cOptionsNames = m_clsPdfCreator.cOptionsNames()
End Property

Public Property Get cOptionsProfile(ByVal ProfileName As String) As LibPdfCreatorOptions
    Set cOptionsProfile = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(m_clsPdfCreator(ProfileName))
End Property

Public Property Set cOptionsProfile(ByVal ProfileName As String, ByVal Options1 As LibPdfCreatorOptions)
    Set m_clsPdfCreator.cOptionsProfile(ProfileName) = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToCOM(Options1)
End Property

Public Property Get cOutputFilename() As String
    cOutputFilename = m_clsPdfCreator.cOutputFilename
End Property

Public Property Get cPDFCreatorApplicationPath() As String
    cPDFCreatorApplicationPath = m_clsPdfCreator.cPDFCreatorApplicationPath
End Property

Public Property Get cPostscriptInfo(ByVal PostScriptFilename As String, ByVal PropertyName As String) As String
    cPostscriptInfo = m_clsPdfCreator(PostScriptFilename, PropertyName)
End Property

Public Property Get cPrinterProfile(ByVal PrinterName As String) As String
    cPrinterProfile = m_clsPdfCreator.cPrinterProfile(PrinterName)
End Property

Public Property Let cPrinterProfile(ByVal PrinterName As String, ByVal ProfileName As String)
    m_clsPdfCreator.cPrinterProfile(PrinterName) = ProfileName
End Property

Public Property Get cPrinterStop() As Boolean
    cPrinterStop = m_clsPdfCreator.cPrinterStop
End Property

Public Property Let cPrinterStop(ByVal value As Boolean)
    m_clsPdfCreator.cPrinterStop = value
End Property

Public Property Get cPrintjobFilename(ByVal JobNumber As Long) As String
    cPrintjobFilename = m_clsPdfCreator.cPrintjobFilename(JobNumber)
End Property

Public Property Get cProgramIsRunning() As Boolean
    cProgramIsRunning = m_clsPdfCreator.cProgramIsRunning
End Property

Public Property Get cProgramRelease(Optional ByVal WithBeta As Boolean = True) As String
    cProgramRelease = m_clsPdfCreator.cProgramRelease(WithBeta)
End Property

Public Property Get cStandardOption(ByVal PropertyName As String) As Variant
    cStandardOption = m_clsPdfCreator.cStandardOption(PropertyName)
End Property

Public Property Get cStandardOptions() As LibPdfCreatorOptions
    Set cStandardOptions = LibPdfCreatorConverter.ConvertPdfCreatorOption_ToLib(m_clsPdfCreator.cStandardOptions)
End Property

Public Property Get cStopURLPrinting() As Boolean
    cStopURLPrinting = m_clsPdfCreator.cStopURLPrinting
End Property

Public Property Let cStopURLPrinting(ByVal value As Boolean)
    m_clsPdfCreator.cStopURLPrinting = value
End Property

Public Property Get cVisible() As Boolean
    cVisible = m_clsPdfCreator.cVisible
End Property

Public Property Let cVisible(ByVal value As Boolean)
    m_clsPdfCreator.cVisible = value
End Property

Public Property Get cWindowState() As Long
    cWindowState = m_clsPdfCreator.cWindowState
End Property

' Value must be '0 = normal' or '1 = minimized' or '2 = maximized'
Public Property Let cWindowState(ByVal value As Long)
    m_clsPdfCreator.cWindowState = value
End Property
    
Public Property Get cWindowsVersion() As String
    cWindowsVersion = m_clsPdfCreator.cWindowsVersion
End Property
