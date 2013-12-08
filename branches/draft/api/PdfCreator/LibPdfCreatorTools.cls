VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorTools
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Diverse Hilfsmittel
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorTools.cls</file>
'  <use>api/PdfCreator/LibPdfCreatorEnums</use>
'  <use>api/PdfCreator/LibPdfCreatorOptions</use>
'  <use>api/PdfCreator/LibPdfCreatorJob.cls</use>
'  <use>file/FileTools.bas</use>
'  <use>text/StringTools.bas</use>
'  <test>_test/api/PdfCreator/LibPdfCreatorToolsTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Function ConvertReportToPDF( _
                                     ByVal RptName As String, _
                                     ByVal OutputPath As String, _
                            Optional ByVal StartPDFViewer As Boolean = True, _
                            Optional ByVal ReportFilter As String = vbNullString, _
                            Optional ByVal SetMetaData As Boolean = False, _
                            Optional ByVal Author As String = vbNullString, _
                            Optional ByVal Title As String = vbNullString, _
                            Optional ByVal Subject As String = vbNullString, _
                            Optional ByVal SetSecurity As Boolean = False, _
                            Optional ByVal OwnerPasswordString As String = ")(/js)/&T&qR&T/a%F(D(U~Z)=I`*'Ä7z", _
                            Optional ByVal DisallowCopy As Boolean = True, _
                            Optional ByVal DisallowModifyAnnotations As Boolean = True, _
                            Optional ByVal DisallowModifyContents As Boolean = True, _
                            Optional ByVal DisallowPrinting As Boolean = False, _
                            Optional ByVal AllowAssembly As Boolean = False, _
                            Optional ByVal AllowDegradedPrinting As Boolean = False, _
                            Optional ByVal PDFAllowFillIn As Boolean = True, _
                            Optional ByVal AllowScreenReaders As Boolean = True, _
                            Optional ByVal Aes128Encryption As Boolean = True _
                            ) As Boolean
 

On Error GoTo ERROR_HANDLER

    ConvertReportToPDF = False
    
    If Len(RptName & vbNullString) = 0 Then _
        Err.Raise ReportNameIsEmpty, "LibPdfCreatorTools", "ReportName ist leer oder wurde nicht übergeben."
 
    If FileTools.FileExists(OutputPath) Then _
        Kill OutputPath
 
    Dim filePathFolder As String
        filePathFolder = FileTools.GetDirFromFullFileName(OutputPath)
    Dim filePathFileName As String
        filePathFileName = FileTools.FileNameWithoutPath(OutputPath)
    Dim filePathFileExtension As String
        filePathFileExtension = VBA.Strings.LCase$(FileTools.GetFileExtension(filePathFileName))
  
    Dim outputFormat As LibPdfCreatorFileFormat
        outputFormat = LibPdfCreatorTools.GetPdfCreatorFileFormatFromFile(filePathFileName)
    
    Dim job As LibPdfCreatorJob
    Set job = New LibPdfCreatorJob
   
    Dim opt As LibPdfCreatorOptions
    Set opt = job.CreatorOptions
    
    opt.UseAutosave = LibPdfCreatorOptionFlag.Yes
    opt.UseAutosaveDirectory = LibPdfCreatorOptionFlag.Yes
    opt.AutosaveDirectory = filePathFolder
    opt.AutosaveFilename = VBA.Strings.Replace(filePathFileName, filePathFileExtension, vbNullString) '(Filename ohne Dateiendung)
    opt.AutosaveFormat = outputFormat
    opt.AutosaveStartStandardProgram = LibPdfCreatorTools.GetPDFCreatorOptionFlag(StartPDFViewer)
        
    'opt.DisableUpdateCheck = PDFCreatorOptionFlag.Yes 'not in 1.2, works in 1.5
    'opt.EditWithPDFArchitect = PDFCreatorOptionFlag.No 'not in 1.2, works in 1.5
    
    If SetMetaData Then _
        SetPdfMetaData opt, Author, Title, Subject
        
    If SetSecurity Then _
        SetPdfSecurity opt, OwnerPasswordString, DisallowCopy, DisallowModifyAnnotations, _
                       DisallowModifyContents, DisallowPrinting, AllowAssembly, AllowDegradedPrinting, _
                       PDFAllowFillIn, AllowScreenReaders, Aes128Encryption

    'Options zurückschreiben
    job.SetOptions opt
    
    ConvertReportToPDF = job.StartPrintJob(RptName, ReportFilter)

EXIT_HANDLER:
    On Error Resume Next
    Set job = Nothing
    Exit Function

ERROR_HANDLER:
    Select Case Err.Number
        Case 70:
            Err.Raise LibPdfCreatorErrorNumer.ZielDateiSchreibgeschützt, "LibPdfCreatorTools", "Die Zieldatei (" & FileTools.FileNameWithoutPath(OutputPath) & ") konnte nicht gelöscht oder überschrieben werden, da sie schreibgeschützt oder in einem anderen Programm geöffnet ist."
        Case Else:
            Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    End Select
    Resume EXIT_HANDLER
    
End Function

'PDF-Methadaten
'As PDFCreator.clsPDFCreatorOptions
Private Sub SetPdfMetaData(ByRef LibPdfCreatorOptions As LibPdfCreatorOptions, _
                          ByVal Author As String, _
                          ByVal Title As String, _
                          ByVal Subject As String)

    LibPdfCreatorOptions.UseStandardAuthor = LibPdfCreatorOptionFlag.Yes
    LibPdfCreatorOptions.StandardAuthor = Author
    LibPdfCreatorOptions.StandardTitle = Title
    LibPdfCreatorOptions.StandardSubject = Subject
    
End Sub

'PDF-Security
'As PDFCreator.clsPDFCreatorOptions
Private Sub SetPdfSecurity(ByRef LibPdfCreatorOptions As LibPdfCreatorOptions, _
                           ByVal PasswordString As String, _
                  Optional ByVal DisallowCopy As Boolean = True, _
                  Optional ByVal DisallowModifyAnnotations As Boolean = True, _
                  Optional ByVal DisallowModifyContents As Boolean = True, _
                  Optional ByVal DisallowPrinting As Boolean = False, _
                  Optional ByVal AllowAssembly As Boolean = False, _
                  Optional ByVal AllowDegradedPrinting As Boolean = False, _
                  Optional ByVal PDFAllowFillIn As Boolean = True, _
                  Optional ByVal AllowScreenReaders As Boolean = True, _
                  Optional ByVal Aes128Encryption As Boolean = True)

    LibPdfCreatorOptions.PDFUseSecurity = LibPdfCreatorOptionFlag.Yes
    'LibPdfCreatorOptions.PDFOwnerPass = LibPdfCreatorOptionFlag.Yes
    LibPdfCreatorOptions.PDFOwnerPasswordString = PasswordString
    LibPdfCreatorOptions.PDFDisallowCopy = GetPDFCreatorOptionFlag(DisallowCopy)
    LibPdfCreatorOptions.PDFDisallowModifyAnnotations = GetPDFCreatorOptionFlag(DisallowModifyAnnotations)
    LibPdfCreatorOptions.PDFDisallowModifyContents = GetPDFCreatorOptionFlag(DisallowModifyContents)
    LibPdfCreatorOptions.PDFDisallowPrinting = GetPDFCreatorOptionFlag(DisallowPrinting)
    LibPdfCreatorOptions.PDFAllowAssembly = GetPDFCreatorOptionFlag(AllowAssembly)
    LibPdfCreatorOptions.PDFAllowDegradedPrinting = GetPDFCreatorOptionFlag(AllowDegradedPrinting)
    LibPdfCreatorOptions.PDFAllowFillIn = GetPDFCreatorOptionFlag(PDFAllowFillIn)
    LibPdfCreatorOptions.PDFAllowScreenReaders = GetPDFCreatorOptionFlag(AllowScreenReaders)
    '1: Öffnen ab Acrobat 7, '0: Öffnen ab Acrobat 3 (nicht alle Sicherheitsfeatures einstellbar)
    LibPdfCreatorOptions.PDFAes128Encryption = GetPDFCreatorOptionFlag(Aes128Encryption)
End Sub

'Durchsucht alle Windows-Drucker nach dem PDFCreator-Drucker.
'Der Name des Druckers muss den Bezeichner "PDFCreator" enthalten.
'Sind mehrere PDF-Creator-Drucker vorhanden, wird der erste Zurückgegeben, es sei denn
'es wird der korrekte Name (printerName) übergeben.
Public Property Get pdfCreatorPrinter(ByRef creator As LibPdfCreatorInstance, _
                             Optional ByVal PrinterName As String = "PDFCreator") As String

'TODO: Contains ist an dieser Stelle nicht ganz korrekt
    Dim printer As Variant
    For Each printer In creator.cGetPDFCreatorPrinters
        If StringTools.Contains(printer, PrinterName) Then
            pdfCreatorPrinter = printer
            Exit For
        End If
    Next
    
End Property

'Ermittelt das PDFCreatorFileFormat anhand eines Dateinamens bzw. Dateipfads
Public Property Get GetPdfCreatorFileFormatFromFile(ByVal filePath As String) As LibPdfCreatorFileFormat

    GetPdfCreatorFileFormatFromFile = LibPdfCreatorTools.GetPdfCreatorFileFormatFromFileExtension( _
                                      FileTools.GetFileExtension(filePath))

End Property

'Ermittelt das PDFCreatorFileFormat anhand der Dateiendung
Public Property Get GetPdfCreatorFileFormatFromFileExtension(ByVal filePathFileExtension As String) As LibPdfCreatorFileFormat
    
    Select Case VBA.Strings.LCase$(filePathFileExtension)
        Case ".pdf"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PDF
        Case ".png"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PNG
        Case ".jpg"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_JPG
        Case ".bmp"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_BMP
        Case ".pcx"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PCX
        Case ".tif"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_TIF
        Case ".ps"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PS
        Case ".eps"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_EPS
        Case ".txt"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_TXT
        Case ".psd"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PSD
        Case ".pcl"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_PCL
        Case ".raw"
            GetPdfCreatorFileFormatFromFileExtension = LibPdfCreatorFileFormat.LibPdfCreatorFileFormat_RAW
        Case Else
            Err.Raise LibPdfCreatorErrorNumer.NotSupportedPdfCreatorFileFormat, "LibPdfCreatorTools", "Das Dateiformat " & filePathFileExtension & " wird von PdfCreator nicht unterstützt"
    End Select

End Property

Public Function GetPDFCreatorOptionFlag(ByVal boolOption As Boolean) As LibPdfCreatorOptionFlag
    If boolOption Then _
        GetPDFCreatorOptionFlag = LibPdfCreatorOptionFlag.Yes _
    Else _
        GetPDFCreatorOptionFlag = LibPdfCreatorOptionFlag.No
End Function
