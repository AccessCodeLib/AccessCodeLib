VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorJob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorJob
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Repräsentiert den PDF-Creator Druckauftrag
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorJob.cls</file>
'  <use>api/PdfCreator/LibPdfCreatorOptions</use>
'  <use>api/PdfCreator/LibPdfCreatorInstance</use>
'  <use>file/FileTools.bas</use>
'  <test>_test/api/PdfCreator/LibPdfCreatorJobTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_creator As LibPdfCreatorInstance
Attribute m_creator.VB_VarHelpID = -1
Private m_options As LibPdfCreatorOptions
Private m_options_backup As LibPdfCreatorOptions

Private m_defaultPrinter As String
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Sub Class_Initialize()
    Set m_creator = LibPdfCreatorFactory.PdfCreatorInstance()
    Set m_options_backup = m_creator.cOptions 'backup user settings
    Set m_options = m_creator.cStandardOptions 'set Standard-Options
    
    If m_creator.cError.Number <> 0 Then _
        Err.Raise LibPdfCreatorErrorNumer.UnexpectedPdfCreatorError, "PDFCreatorError", m_creator.cError.Description & " - " & m_creator.cError.Number
    
    If m_creator.cCountOfPrintjobs > 0 Then _
        Err.Raise LibPdfCreatorErrorNumer.WarteSchlangeNichtLeer, , "Es befinden sich noch nicht abgeschlossene Druckaufträge in der Warteschlange"
    
    m_defaultPrinter = m_creator.cDefaultPrinter 'backup windows default printer
    m_creator.cDefaultPrinter = LibPdfCreatorTools.pdfCreatorPrinter(m_creator) 'set pdfCreator to default printer

End Sub

Private Sub Class_Terminate()
    Set m_creator.cOptions = m_options_backup 'restore user settings
        m_creator.cDefaultPrinter = m_defaultPrinter 'restore windows default printer
        m_creator.cClose
    Set m_options = Nothing
    Set m_creator = Nothing
End Sub

Public Property Get creator() As LibPdfCreatorInstance
    Set creator = m_creator
End Property

Public Property Get CreatorOptions() As LibPdfCreatorOptions
    Set CreatorOptions = m_options
End Property

''Event: Wenn PdfCreator den Druckauftrag abgeschlossen hat
'Public Sub m_creator_eReady()
'    m_isReady = FileTools.FileExists(m_creator.cOutputFilename)
'End Sub

'Event: Wenn während der Verarbeitung ein Fehler auftritt (ungetestet)
'Public Sub m_creator_eError()
'    Set m_creator_error = m_creator.cError
'    Err.Raise LibPdfCreatorErrorNumer.UnexpectedPdfCreatorError, "LibPdfCreatorJob", m_creator_error.Description & " - " & m_creator_error.Number
'End Sub

Public Property Get IsReady() As Boolean
    IsReady = FileTools.FileExists(m_creator.cOutputFilename)
End Property

Public Sub SetOptions(ByRef options As LibPdfCreatorOptions)
    Set m_creator.cOptions = options
End Sub

Public Function StartPrintJob(ByVal RptName As String, ByVal ReportFilter As String) As Boolean
On Error GoTo Err_

    StartPrintJob = False
    m_creator.cClearCache
    
    DoCmd.OpenReport RptName, acViewNormal, , ReportFilter, acHidden
    
    Do Until m_creator.cCountOfPrintjobs = 1 'Warten bis der Druckauftrag in der Warteschlange angekommen ist
        DoEvents
        Sleep 100
    Loop
    
    m_creator.cPrinterStop = False

    'TODO:  Was ist wenn eine andere Anwendung auch druckt?
    Do Until m_creator.cCountOfPrintjobs = 0 And Me.IsReady 'Warten bis der Druckauftrag abgebarbeitet ist
        DoEvents
        Sleep 400
    Loop
    
    If m_creator.cError.Number <> 0 Then _
        Err.Raise LibPdfCreatorErrorNumer.UnexpectedPdfCreatorError, "PDFCreatorError", m_creator.cError.Description & " - Error: " & Me.creator.cError.Number
    
    StartPrintJob = Me.IsReady
Err_:
    Select Case Err.Number
        Case 0:
        Case 2103:
            Err.Raise LibPdfCreatorErrorNumer.ReportDoesNotExist, "LibPdfCreatorJob", "Der Report " & RptName & " konnte nicht gefunden werden."
        Case Else:
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select

End Function
