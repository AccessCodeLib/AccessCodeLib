VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorOptions
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Repräsentiert die PDF-Creator Einstellungen (Druckoptionen)
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorOptions.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_AdditionalGhostscriptParameters As String
Private m_AdditionalGhostscriptSearchpath As String
Private m_AddWindowsFontpath As Long
Private m_AllowSpecialGSCharsInFilenames As Long
Private m_AutosaveDirectory As String
Private m_AutosaveFilename As String
Private m_AutosaveFormat As Long
Private m_AutosaveStartStandardProgram As Long
Private m_BMPColorscount As Long
Private m_BMPResolution As Long
Private m_ClientComputerResolveIPAddress As Long
Private m_Counter As Currency
Private m_DeviceHeightPoints As Double
Private m_DeviceWidthPoints As Double
Private m_DirectoryGhostscriptBinaries As String
Private m_DirectoryGhostscriptFonts As String
Private m_DirectoryGhostscriptLibraries As String
Private m_DirectoryGhostscriptResource As String
Private m_DisableEmail As Long
Private m_DisableUpdateCheck As Long
Private m_DontUseDocumentSettings As Long
Private m_EditWithPDFArchitect As Long
Private m_EPSLanguageLevel As Long
Private m_FilenameSubstitutions As String
Private m_FilenameSubstitutionsOnlyInTitle As Long
Private m_JPEGColorscount As Long
Private m_JPEGQuality As Long
Private m_JPEGResolution As Long
Private m_Language As String
Private m_LastSaveDirectory As String
Private m_LastUpdateCheck As String
Private m_Logging As Long
Private m_LogLines As Long
Private m_MaximumCountOfPDFArchitectToolTip As Long
Private m_NoConfirmMessageSwitchingDefaultprinter As Long
Private m_NoProcessingAtStartup As Long
Private m_NoPSCheck As Long
Private m_OneFilePerPage As Long
Private m_OpenOutputFile As Long
Private m_OptionsDesign As Long
Private m_OptionsEnabled As Long
Private m_OptionsVisible As Long
Private m_Papersize As String
Private m_PCLColorsCount As Long
Private m_PCLResolution As Long
Private m_PCXColorscount As Long
Private m_PCXResolution As Long
Private m_PDFAes128Encryption As Long
Private m_PDFAllowAssembly As Long
Private m_PDFAllowDegradedPrinting As Long
Private m_PDFAllowFillIn As Long
Private m_PDFAllowScreenReaders As Long
Private m_PDFColorsCMYKToRGB As Long
Private m_PDFColorsColorModel As Long
Private m_PDFColorsPreserveHalftone As Long
Private m_PDFColorsPreserveOverprint As Long
Private m_PDFColorsPreserveTransfer As Long
Private m_PDFCompressionColorCompression As Long
Private m_PDFCompressionColorCompressionChoice As Long
Private m_PDFCompressionColorCompressionJPEGHighFactor As Double
Private m_PDFCompressionColorCompressionJPEGLowFactor As Double
Private m_PDFCompressionColorCompressionJPEGManualFactor As Double
Private m_PDFCompressionColorCompressionJPEGMaximumFactor As Double
Private m_PDFCompressionColorCompressionJPEGMediumFactor As Double
Private m_PDFCompressionColorCompressionJPEGMinimumFactor As Double
Private m_PDFCompressionColorResample As Long
Private m_PDFCompressionColorResampleChoice As Long
Private m_PDFCompressionColorResolution As Long
Private m_PDFCompressionGreyCompression As Long
Private m_PDFCompressionGreyCompressionChoice As Long
Private m_PDFCompressionGreyCompressionJPEGHighFactor As Double
Private m_PDFCompressionGreyCompressionJPEGLowFactor As Double
Private m_PDFCompressionGreyCompressionJPEGManualFactor As Double
Private m_PDFCompressionGreyCompressionJPEGMaximumFactor As Double
Private m_PDFCompressionGreyCompressionJPEGMediumFactor As Double
Private m_PDFCompressionGreyCompressionJPEGMinimumFactor As Double
Private m_PDFCompressionGreyResample As Long
Private m_PDFCompressionGreyResampleChoice As Long
Private m_PDFCompressionGreyResolution As Long
Private m_PDFCompressionMonoCompression As Long
Private m_PDFCompressionMonoCompressionChoice As Long
Private m_PDFCompressionMonoResample As Long
Private m_PDFCompressionMonoResampleChoice As Long
Private m_PDFCompressionMonoResolution As Long
Private m_PDFCompressionTextCompression As Long
Private m_PDFDisallowCopy As Long
Private m_PDFDisallowModifyAnnotations As Long
Private m_PDFDisallowModifyContents As Long
Private m_PDFDisallowPrinting As Long
Private m_PDFEncryptor As Long
Private m_PDFFontsEmbedAll As Long
Private m_PDFFontsSubSetFonts As Long
Private m_PDFFontsSubSetFontsPercent As Long
Private m_PDFGeneralASCII85 As Long
Private m_PDFGeneralAutorotate As Long
Private m_PDFGeneralCompatibility As Long
Private m_PDFGeneralDefault As Long
Private m_PDFGeneralOverprint As Long
Private m_PDFHighEncryption As Long
Private m_PDFLowEncryption As Long
Private m_PDFOptimize As Long
Private m_PDFOwnerPass As Long
Private m_PDFOwnerPasswordString As String
Private m_PDFPageLayout As Long
Private m_PDFPageMode As Long
Private m_PDFSigningMultiSignature As Long
Private m_PDFSigningPFXFile As String
Private m_PDFSigningPFXFilePassword As String
Private m_PDFSigningSignatureContact As String
Private m_PDFSigningSignatureLeftX As Double
Private m_PDFSigningSignatureLeftY As Double
Private m_PDFSigningSignatureLocation As String
Private m_PDFSigningSignatureOnPage As Long
Private m_PDFSigningSignatureReason As String
Private m_PDFSigningSignatureRightX As Double
Private m_PDFSigningSignatureRightY As Double
Private m_PDFSigningSignatureVisible As Long
Private m_PDFSigningSignPDF As Long
Private m_PDFSigningTimeServerUrl As String
Private m_PDFStartPage As Long
Private m_PDFUpdateMetadata As Long
Private m_PDFUserPass As Long
Private m_PDFUserPasswordString As String
Private m_PDFUseSecurity As Long
Private m_PNGColorscount As Long
Private m_PNGResolution As Long
Private m_PrintAfterSaving As Long
Private m_PrintAfterSavingBitsPerPixel As Long
Private m_PrintAfterSavingDuplex As Long
Private m_PrintAfterSavingMaxResolution As Long
Private m_PrintAfterSavingMaxResolutionEnabled As Long
Private m_PrintAfterSavingNoCancel As Long
Private m_PrintAfterSavingPrinter As String
Private m_PrintAfterSavingQueryUser As Long
Private m_PrintAfterSavingTumble As Long
Private m_PrinterStop As Long
Private m_ProcessPriority As Long
Private m_ProgramFont As String
Private m_ProgramFontCharset As Long
Private m_ProgramFontSize As Long
Private m_PSDColorsCount As Long
Private m_PSDResolution As Long
Private m_PSLanguageLevel As Long
Private m_RAWColorsCount As Long
Private m_RAWResolution As Long
Private m_RemoveAllKnownFileExtensions As Long
Private m_RemoveSpaces As Long
Private m_RunProgramAfterSaving As Long
Private m_RunProgramAfterSavingProgramname As String
Private m_RunProgramAfterSavingProgramParameters As String
Private m_RunProgramAfterSavingWaitUntilReady As Long
Private m_RunProgramAfterSavingWindowstyle As Long
Private m_RunProgramBeforeSaving As Long
Private m_RunProgramBeforeSavingProgramname As String
Private m_RunProgramBeforeSavingProgramParameters As String
Private m_RunProgramBeforeSavingWindowstyle As Long
Private m_SaveFilename As String
Private m_SendEmailAfterAutoSaving As Long
Private m_SendMailMethod As Long
Private m_ShowAnimation As Long
Private m_StampFontColor As String
Private m_StampFontname As String
Private m_StampFontsize As Long
Private m_StampOutlineFontthickness As Long
Private m_StampString As String
Private m_StampUseOutlineFont As Long
Private m_StandardAuthor As String
Private m_StandardCreationdate As String
Private m_StandardDateformat As String
Private m_StandardKeywords As String
Private m_StandardMailDomain As String
Private m_StandardModifydate As String
Private m_StandardSaveformat As Long
Private m_StandardSubject As String
Private m_StandardTitle As String
Private m_SVGResolution As Long
Private m_TIFFColorscount As Long
Private m_TIFFResolution As Long
Private m_Toolbars As Long
Private m_UpdateInterval As Long
Private m_UseAutosave As Long
Private m_UseAutosaveDirectory As Long
Private m_UseCreationDateNow As Long
Private m_UseCustomPaperSize As String
Private m_UseFixPapersize As Long
Private m_UseStandardAuthor As Long

Public Property Get AdditionalGhostscriptParameters() As String
       AdditionalGhostscriptParameters = m_AdditionalGhostscriptParameters
End Property

Public Property Let AdditionalGhostscriptParameters(ByVal value As String)
        m_AdditionalGhostscriptParameters = value
End Property

Public Property Get AdditionalGhostscriptSearchpath() As String
       AdditionalGhostscriptSearchpath = m_AdditionalGhostscriptSearchpath
End Property

Public Property Let AdditionalGhostscriptSearchpath(ByVal value As String)
        m_AdditionalGhostscriptSearchpath = value
End Property

Public Property Get AddWindowsFontpath() As Variant
       AddWindowsFontpath = m_AddWindowsFontpath
End Property

Public Property Let AddWindowsFontpath(ByVal value As Variant)
        m_AddWindowsFontpath = value
End Property

Public Property Get AllowSpecialGSCharsInFilenames() As Variant
       AllowSpecialGSCharsInFilenames = m_AllowSpecialGSCharsInFilenames
End Property

Public Property Let AllowSpecialGSCharsInFilenames(ByVal value As Variant)
        m_AllowSpecialGSCharsInFilenames = value
End Property

Public Property Get AutosaveDirectory() As String
       AutosaveDirectory = m_AutosaveDirectory
End Property

Public Property Let AutosaveDirectory(ByVal value As String)
        m_AutosaveDirectory = value
End Property

Public Property Get AutosaveFilename() As String
       AutosaveFilename = m_AutosaveFilename
End Property

Public Property Let AutosaveFilename(ByVal value As String)
        m_AutosaveFilename = value
End Property

Public Property Get AutosaveFormat() As Variant
       AutosaveFormat = m_AutosaveFormat
End Property

Public Property Let AutosaveFormat(ByVal value As Variant)
        m_AutosaveFormat = value
End Property

Public Property Get AutosaveStartStandardProgram() As Variant
       AutosaveStartStandardProgram = m_AutosaveStartStandardProgram
End Property

Public Property Let AutosaveStartStandardProgram(ByVal value As Variant)
        m_AutosaveStartStandardProgram = value
End Property

Public Property Get BMPColorscount() As Variant
       BMPColorscount = m_BMPColorscount
End Property

Public Property Let BMPColorscount(ByVal value As Variant)
        m_BMPColorscount = value
End Property

Public Property Get BMPResolution() As Variant
       BMPResolution = m_BMPResolution
End Property

Public Property Let BMPResolution(ByVal value As Variant)
        m_BMPResolution = value
End Property

Public Property Get ClientComputerResolveIPAddress() As Variant
       ClientComputerResolveIPAddress = m_ClientComputerResolveIPAddress
End Property

Public Property Let ClientComputerResolveIPAddress(ByVal value As Variant)
        m_ClientComputerResolveIPAddress = value
End Property

Public Property Get Counter() As Variant
       Counter = m_Counter
End Property

Public Property Let Counter(ByVal value As Variant)
        m_Counter = value
End Property

Public Property Get DeviceHeightPoints() As Variant
       DeviceHeightPoints = m_DeviceHeightPoints
End Property

Public Property Let DeviceHeightPoints(ByVal value As Variant)
        m_DeviceHeightPoints = value
End Property

Public Property Get DeviceWidthPoints() As Variant
       DeviceWidthPoints = m_DeviceWidthPoints
End Property

Public Property Let DeviceWidthPoints(ByVal value As Variant)
        m_DeviceWidthPoints = value
End Property

Public Property Get DirectoryGhostscriptBinaries() As Variant
       DirectoryGhostscriptBinaries = m_DirectoryGhostscriptBinaries
End Property

Public Property Let DirectoryGhostscriptBinaries(ByVal value As Variant)
        m_DirectoryGhostscriptBinaries = value
End Property

Public Property Get DirectoryGhostscriptFonts() As Variant
       DirectoryGhostscriptFonts = m_DirectoryGhostscriptFonts
End Property

Public Property Let DirectoryGhostscriptFonts(ByVal value As Variant)
        m_DirectoryGhostscriptFonts = value
End Property

Public Property Get DirectoryGhostscriptLibraries() As Variant
       DirectoryGhostscriptLibraries = m_DirectoryGhostscriptLibraries
End Property

Public Property Let DirectoryGhostscriptLibraries(ByVal value As Variant)
        m_DirectoryGhostscriptLibraries = value
End Property

Public Property Get DirectoryGhostscriptResource() As Variant
       DirectoryGhostscriptResource = m_DirectoryGhostscriptResource
End Property

Public Property Let DirectoryGhostscriptResource(ByVal value As Variant)
        m_DirectoryGhostscriptResource = value
End Property

Public Property Get DisableEmail() As Variant
       DisableEmail = m_DisableEmail
End Property

Public Property Let DisableEmail(ByVal value As Variant)
        m_DisableEmail = value
End Property

Public Property Get DisableUpdateCheck() As Variant
       DisableUpdateCheck = m_DisableUpdateCheck
End Property

Public Property Let DisableUpdateCheck(ByVal value As Variant)
        m_DisableUpdateCheck = value
End Property

Public Property Get DontUseDocumentSettings() As Variant
       DontUseDocumentSettings = m_DontUseDocumentSettings
End Property

Public Property Let DontUseDocumentSettings(ByVal value As Variant)
        m_DontUseDocumentSettings = value
End Property

Public Property Get EditWithPDFArchitect() As Variant
       EditWithPDFArchitect = m_EditWithPDFArchitect
End Property

Public Property Let EditWithPDFArchitect(ByVal value As Variant)
        m_EditWithPDFArchitect = value
End Property

Public Property Get EPSLanguageLevel() As Variant
       EPSLanguageLevel = m_EPSLanguageLevel
End Property

Public Property Let EPSLanguageLevel(ByVal value As Variant)
        m_EPSLanguageLevel = value
End Property

Public Property Get FilenameSubstitutions() As Variant
       FilenameSubstitutions = m_FilenameSubstitutions
End Property

Public Property Let FilenameSubstitutions(ByVal value As Variant)
        m_FilenameSubstitutions = value
End Property

Public Property Get FilenameSubstitutionsOnlyInTitle() As Variant
       FilenameSubstitutionsOnlyInTitle = m_FilenameSubstitutionsOnlyInTitle
End Property

Public Property Let FilenameSubstitutionsOnlyInTitle(ByVal value As Variant)
        m_FilenameSubstitutionsOnlyInTitle = value
End Property

Public Property Get JPEGColorscount() As Variant
       JPEGColorscount = m_JPEGColorscount
End Property

Public Property Let JPEGColorscount(ByVal value As Variant)
        m_JPEGColorscount = value
End Property

Public Property Get JPEGQuality() As Variant
       JPEGQuality = m_JPEGQuality
End Property

Public Property Let JPEGQuality(ByVal value As Variant)
        m_JPEGQuality = value
End Property

Public Property Get JPEGResolution() As Variant
       JPEGResolution = m_JPEGResolution
End Property

Public Property Let JPEGResolution(ByVal value As Variant)
        m_JPEGResolution = value
End Property

Public Property Get Language() As Variant
       Language = m_Language
End Property

Public Property Let Language(ByVal value As Variant)
        m_Language = value
End Property

Public Property Get LastSaveDirectory() As Variant
       LastSaveDirectory = m_LastSaveDirectory
End Property

Public Property Let LastSaveDirectory(ByVal value As Variant)
        m_LastSaveDirectory = value
End Property

Public Property Get LastUpdateCheck() As Variant
       LastUpdateCheck = m_LastUpdateCheck
End Property

Public Property Let LastUpdateCheck(ByVal value As Variant)
        m_LastUpdateCheck = value
End Property

Public Property Get Logging() As Variant
       Logging = m_Logging
End Property

Public Property Let Logging(ByVal value As Variant)
        m_Logging = value
End Property

Public Property Get LogLines() As Variant
       LogLines = m_LogLines
End Property

Public Property Let LogLines(ByVal value As Variant)
        m_LogLines = value
End Property

Public Property Get MaximumCountOfPDFArchitectToolTip() As Variant
       MaximumCountOfPDFArchitectToolTip = m_MaximumCountOfPDFArchitectToolTip
End Property

Public Property Let MaximumCountOfPDFArchitectToolTip(ByVal value As Variant)
        m_MaximumCountOfPDFArchitectToolTip = value
End Property

Public Property Get NoConfirmMessageSwitchingDefaultprinter() As Variant
       NoConfirmMessageSwitchingDefaultprinter = m_NoConfirmMessageSwitchingDefaultprinter
End Property

Public Property Let NoConfirmMessageSwitchingDefaultprinter(ByVal value As Variant)
        m_NoConfirmMessageSwitchingDefaultprinter = value
End Property

Public Property Get NoProcessingAtStartup() As Variant
       NoProcessingAtStartup = m_NoProcessingAtStartup
End Property

Public Property Let NoProcessingAtStartup(ByVal value As Variant)
        m_NoProcessingAtStartup = value
End Property

Public Property Get NoPSCheck() As Variant
       NoPSCheck = m_NoPSCheck
End Property

Public Property Let NoPSCheck(ByVal value As Variant)
        m_NoPSCheck = value
End Property

Public Property Get OneFilePerPage() As Variant
       OneFilePerPage = m_OneFilePerPage
End Property

Public Property Let OneFilePerPage(ByVal value As Variant)
        m_OneFilePerPage = value
End Property

Public Property Get OpenOutputFile() As Variant
       OpenOutputFile = m_OpenOutputFile
End Property

Public Property Let OpenOutputFile(ByVal value As Variant)
        m_OpenOutputFile = value
End Property

Public Property Get OptionsDesign() As Variant
       OptionsDesign = m_OptionsDesign
End Property

Public Property Let OptionsDesign(ByVal value As Variant)
        m_OptionsDesign = value
End Property

Public Property Get OptionsEnabled() As Variant
       OptionsEnabled = m_OptionsEnabled
End Property

Public Property Let OptionsEnabled(ByVal value As Variant)
        m_OptionsEnabled = value
End Property

Public Property Get OptionsVisible() As Variant
       OptionsVisible = m_OptionsVisible
End Property

Public Property Let OptionsVisible(ByVal value As Variant)
        m_OptionsVisible = value
End Property

Public Property Get Papersize() As Variant
       Papersize = m_Papersize
End Property

Public Property Let Papersize(ByVal value As Variant)
        m_Papersize = value
End Property

Public Property Get PCLColorsCount() As Variant
       PCLColorsCount = m_PCLColorsCount
End Property

Public Property Let PCLColorsCount(ByVal value As Variant)
        m_PCLColorsCount = value
End Property

Public Property Get PCLResolution() As Variant
       PCLResolution = m_PCLResolution
End Property

Public Property Let PCLResolution(ByVal value As Variant)
        m_PCLResolution = value
End Property

Public Property Get PCXColorscount() As Variant
       PCXColorscount = m_PCXColorscount
End Property

Public Property Let PCXColorscount(ByVal value As Variant)
        m_PCXColorscount = value
End Property

Public Property Get PCXResolution() As Variant
       PCXResolution = m_PCXResolution
End Property

Public Property Let PCXResolution(ByVal value As Variant)
        m_PCXResolution = value
End Property

Public Property Get PDFAes128Encryption() As Variant
       PDFAes128Encryption = m_PDFAes128Encryption
End Property

Public Property Let PDFAes128Encryption(ByVal value As Variant)
        m_PDFAes128Encryption = value
End Property

Public Property Get PDFAllowAssembly() As Variant
       PDFAllowAssembly = m_PDFAllowAssembly
End Property

Public Property Let PDFAllowAssembly(ByVal value As Variant)
        m_PDFAllowAssembly = value
End Property

Public Property Get PDFAllowDegradedPrinting() As Variant
       PDFAllowDegradedPrinting = m_PDFAllowDegradedPrinting
End Property

Public Property Let PDFAllowDegradedPrinting(ByVal value As Variant)
        m_PDFAllowDegradedPrinting = value
End Property

Public Property Get PDFAllowFillIn() As Variant
       PDFAllowFillIn = m_PDFAllowFillIn
End Property

Public Property Let PDFAllowFillIn(ByVal value As Variant)
        m_PDFAllowFillIn = value
End Property

Public Property Get PDFAllowScreenReaders() As Variant
       PDFAllowScreenReaders = m_PDFAllowScreenReaders
End Property

Public Property Let PDFAllowScreenReaders(ByVal value As Variant)
        m_PDFAllowScreenReaders = value
End Property

Public Property Get PDFColorsCMYKToRGB() As Variant
       PDFColorsCMYKToRGB = m_PDFColorsCMYKToRGB
End Property

Public Property Let PDFColorsCMYKToRGB(ByVal value As Variant)
        m_PDFColorsCMYKToRGB = value
End Property

Public Property Get PDFColorsColorModel() As Variant
       PDFColorsColorModel = m_PDFColorsColorModel
End Property

Public Property Let PDFColorsColorModel(ByVal value As Variant)
        m_PDFColorsColorModel = value
End Property

Public Property Get PDFColorsPreserveHalftone() As Variant
       PDFColorsPreserveHalftone = m_PDFColorsPreserveHalftone
End Property

Public Property Let PDFColorsPreserveHalftone(ByVal value As Variant)
        m_PDFColorsPreserveHalftone = value
End Property

Public Property Get PDFColorsPreserveOverprint() As Variant
       PDFColorsPreserveOverprint = m_PDFColorsPreserveOverprint
End Property

Public Property Let PDFColorsPreserveOverprint(ByVal value As Variant)
        m_PDFColorsPreserveOverprint = value
End Property

Public Property Get PDFColorsPreserveTransfer() As Variant
       PDFColorsPreserveTransfer = m_PDFColorsPreserveTransfer
End Property

Public Property Let PDFColorsPreserveTransfer(ByVal value As Variant)
        m_PDFColorsPreserveTransfer = value
End Property

Public Property Get PDFCompressionColorCompression() As Variant
       PDFCompressionColorCompression = m_PDFCompressionColorCompression
End Property

Public Property Let PDFCompressionColorCompression(ByVal value As Variant)
        m_PDFCompressionColorCompression = value
End Property

Public Property Get PDFCompressionColorCompressionChoice() As Variant
       PDFCompressionColorCompressionChoice = m_PDFCompressionColorCompressionChoice
End Property

Public Property Let PDFCompressionColorCompressionChoice(ByVal value As Variant)
        m_PDFCompressionColorCompressionChoice = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGHighFactor() As Variant
       PDFCompressionColorCompressionJPEGHighFactor = m_PDFCompressionColorCompressionJPEGHighFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGHighFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGHighFactor = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGLowFactor() As Variant
       PDFCompressionColorCompressionJPEGLowFactor = m_PDFCompressionColorCompressionJPEGLowFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGLowFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGLowFactor = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGManualFactor() As Variant
       PDFCompressionColorCompressionJPEGManualFactor = m_PDFCompressionColorCompressionJPEGManualFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGManualFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGManualFactor = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGMaximumFactor() As Variant
       PDFCompressionColorCompressionJPEGMaximumFactor = m_PDFCompressionColorCompressionJPEGMaximumFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGMaximumFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGMaximumFactor = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGMediumFactor() As Variant
       PDFCompressionColorCompressionJPEGMediumFactor = m_PDFCompressionColorCompressionJPEGMediumFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGMediumFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGMediumFactor = value
End Property

Public Property Get PDFCompressionColorCompressionJPEGMinimumFactor() As Variant
       PDFCompressionColorCompressionJPEGMinimumFactor = m_PDFCompressionColorCompressionJPEGMinimumFactor
End Property

Public Property Let PDFCompressionColorCompressionJPEGMinimumFactor(ByVal value As Variant)
        m_PDFCompressionColorCompressionJPEGMinimumFactor = value
End Property

Public Property Get PDFCompressionColorResample() As Variant
       PDFCompressionColorResample = m_PDFCompressionColorResample
End Property

Public Property Let PDFCompressionColorResample(ByVal value As Variant)
        m_PDFCompressionColorResample = value
End Property

Public Property Get PDFCompressionColorResampleChoice() As Variant
       PDFCompressionColorResampleChoice = m_PDFCompressionColorResampleChoice
End Property

Public Property Let PDFCompressionColorResampleChoice(ByVal value As Variant)
        m_PDFCompressionColorResampleChoice = value
End Property

Public Property Get PDFCompressionColorResolution() As Variant
       PDFCompressionColorResolution = m_PDFCompressionColorResolution
End Property

Public Property Let PDFCompressionColorResolution(ByVal value As Variant)
        m_PDFCompressionColorResolution = value
End Property

Public Property Get PDFCompressionGreyCompression() As Variant
       PDFCompressionGreyCompression = m_PDFCompressionGreyCompression
End Property

Public Property Let PDFCompressionGreyCompression(ByVal value As Variant)
        m_PDFCompressionGreyCompression = value
End Property

Public Property Get PDFCompressionGreyCompressionChoice() As Variant
       PDFCompressionGreyCompressionChoice = m_PDFCompressionGreyCompressionChoice
End Property

Public Property Let PDFCompressionGreyCompressionChoice(ByVal value As Variant)
        m_PDFCompressionGreyCompressionChoice = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGHighFactor() As Variant
       PDFCompressionGreyCompressionJPEGHighFactor = m_PDFCompressionGreyCompressionJPEGHighFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGHighFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGHighFactor = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGLowFactor() As Variant
       PDFCompressionGreyCompressionJPEGLowFactor = m_PDFCompressionGreyCompressionJPEGLowFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGLowFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGLowFactor = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGManualFactor() As Variant
       PDFCompressionGreyCompressionJPEGManualFactor = m_PDFCompressionGreyCompressionJPEGManualFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGManualFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGManualFactor = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGMaximumFactor() As Variant
       PDFCompressionGreyCompressionJPEGMaximumFactor = m_PDFCompressionGreyCompressionJPEGMaximumFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGMaximumFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGMaximumFactor = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGMediumFactor() As Variant
       PDFCompressionGreyCompressionJPEGMediumFactor = m_PDFCompressionGreyCompressionJPEGMediumFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGMediumFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGMediumFactor = value
End Property

Public Property Get PDFCompressionGreyCompressionJPEGMinimumFactor() As Variant
       PDFCompressionGreyCompressionJPEGMinimumFactor = m_PDFCompressionGreyCompressionJPEGMinimumFactor
End Property

Public Property Let PDFCompressionGreyCompressionJPEGMinimumFactor(ByVal value As Variant)
        m_PDFCompressionGreyCompressionJPEGMinimumFactor = value
End Property

Public Property Get PDFCompressionGreyResample() As Variant
       PDFCompressionGreyResample = m_PDFCompressionGreyResample
End Property

Public Property Let PDFCompressionGreyResample(ByVal value As Variant)
        m_PDFCompressionGreyResample = value
End Property

Public Property Get PDFCompressionGreyResampleChoice() As Variant
       PDFCompressionGreyResampleChoice = m_PDFCompressionGreyResampleChoice
End Property

Public Property Let PDFCompressionGreyResampleChoice(ByVal value As Variant)
        m_PDFCompressionGreyResampleChoice = value
End Property

Public Property Get PDFCompressionGreyResolution() As Variant
       PDFCompressionGreyResolution = m_PDFCompressionGreyResolution
End Property

Public Property Let PDFCompressionGreyResolution(ByVal value As Variant)
        m_PDFCompressionGreyResolution = value
End Property

Public Property Get PDFCompressionMonoCompression() As Variant
       PDFCompressionMonoCompression = m_PDFCompressionMonoCompression
End Property

Public Property Let PDFCompressionMonoCompression(ByVal value As Variant)
        m_PDFCompressionMonoCompression = value
End Property

Public Property Get PDFCompressionMonoCompressionChoice() As Variant
       PDFCompressionMonoCompressionChoice = m_PDFCompressionMonoCompressionChoice
End Property

Public Property Let PDFCompressionMonoCompressionChoice(ByVal value As Variant)
        m_PDFCompressionMonoCompressionChoice = value
End Property

Public Property Get PDFCompressionMonoResample() As Variant
       PDFCompressionMonoResample = m_PDFCompressionMonoResample
End Property

Public Property Let PDFCompressionMonoResample(ByVal value As Variant)
        m_PDFCompressionMonoResample = value
End Property

Public Property Get PDFCompressionMonoResampleChoice() As Variant
       PDFCompressionMonoResampleChoice = m_PDFCompressionMonoResampleChoice
End Property

Public Property Let PDFCompressionMonoResampleChoice(ByVal value As Variant)
        m_PDFCompressionMonoResampleChoice = value
End Property

Public Property Get PDFCompressionMonoResolution() As Variant
       PDFCompressionMonoResolution = m_PDFCompressionMonoResolution
End Property

Public Property Let PDFCompressionMonoResolution(ByVal value As Variant)
        m_PDFCompressionMonoResolution = value
End Property

Public Property Get PDFCompressionTextCompression() As Variant
       PDFCompressionTextCompression = m_PDFCompressionTextCompression
End Property

Public Property Let PDFCompressionTextCompression(ByVal value As Variant)
        m_PDFCompressionTextCompression = value
End Property

Public Property Get PDFDisallowCopy() As Variant
       PDFDisallowCopy = m_PDFDisallowCopy
End Property

Public Property Let PDFDisallowCopy(ByVal value As Variant)
        m_PDFDisallowCopy = value
End Property

Public Property Get PDFDisallowModifyAnnotations() As Variant
       PDFDisallowModifyAnnotations = m_PDFDisallowModifyAnnotations
End Property

Public Property Let PDFDisallowModifyAnnotations(ByVal value As Variant)
        m_PDFDisallowModifyAnnotations = value
End Property

Public Property Get PDFDisallowModifyContents() As Variant
       PDFDisallowModifyContents = m_PDFDisallowModifyContents
End Property

Public Property Let PDFDisallowModifyContents(ByVal value As Variant)
        m_PDFDisallowModifyContents = value
End Property

Public Property Get PDFDisallowPrinting() As Variant
       PDFDisallowPrinting = m_PDFDisallowPrinting
End Property

Public Property Let PDFDisallowPrinting(ByVal value As Variant)
        m_PDFDisallowPrinting = value
End Property

Public Property Get PDFEncryptor() As Variant
       PDFEncryptor = m_PDFEncryptor
End Property

Public Property Let PDFEncryptor(ByVal value As Variant)
        m_PDFEncryptor = value
End Property

Public Property Get PDFFontsEmbedAll() As Variant
       PDFFontsEmbedAll = m_PDFFontsEmbedAll
End Property

Public Property Let PDFFontsEmbedAll(ByVal value As Variant)
        m_PDFFontsEmbedAll = value
End Property

Public Property Get PDFFontsSubSetFonts() As Variant
       PDFFontsSubSetFonts = m_PDFFontsSubSetFonts
End Property

Public Property Let PDFFontsSubSetFonts(ByVal value As Variant)
        m_PDFFontsSubSetFonts = value
End Property

Public Property Get PDFFontsSubSetFontsPercent() As Variant
       PDFFontsSubSetFontsPercent = m_PDFFontsSubSetFontsPercent
End Property

Public Property Let PDFFontsSubSetFontsPercent(ByVal value As Variant)
        m_PDFFontsSubSetFontsPercent = value
End Property

Public Property Get PDFGeneralASCII85() As Variant
       PDFGeneralASCII85 = m_PDFGeneralASCII85
End Property

Public Property Let PDFGeneralASCII85(ByVal value As Variant)
        m_PDFGeneralASCII85 = value
End Property

Public Property Get PDFGeneralAutorotate() As Variant
       PDFGeneralAutorotate = m_PDFGeneralAutorotate
End Property

Public Property Let PDFGeneralAutorotate(ByVal value As Variant)
        m_PDFGeneralAutorotate = value
End Property

Public Property Get PDFGeneralCompatibility() As Variant
       PDFGeneralCompatibility = m_PDFGeneralCompatibility
End Property

Public Property Let PDFGeneralCompatibility(ByVal value As Variant)
        m_PDFGeneralCompatibility = value
End Property

Public Property Get PDFGeneralDefault() As Variant
       PDFGeneralDefault = m_PDFGeneralDefault
End Property

Public Property Let PDFGeneralDefault(ByVal value As Variant)
        m_PDFGeneralDefault = value
End Property

Public Property Get PDFGeneralOverprint() As Variant
       PDFGeneralOverprint = m_PDFGeneralOverprint
End Property

Public Property Let PDFGeneralOverprint(ByVal value As Variant)
        m_PDFGeneralOverprint = value
End Property

Public Property Get PDFHighEncryption() As Variant
       PDFHighEncryption = m_PDFHighEncryption
End Property

Public Property Let PDFHighEncryption(ByVal value As Variant)
        m_PDFHighEncryption = value
End Property

Public Property Get PDFLowEncryption() As Variant
       PDFLowEncryption = m_PDFLowEncryption
End Property

Public Property Let PDFLowEncryption(ByVal value As Variant)
        m_PDFLowEncryption = value
End Property

Public Property Get PDFOptimize() As Variant
       PDFOptimize = m_PDFOptimize
End Property

Public Property Let PDFOptimize(ByVal value As Variant)
        m_PDFOptimize = value
End Property

'Public Property Get PDFOwnerPass() As Variant
'       PDFOwnerPass = m_PDFOwnerPass
'End Property

'Public Property Let PDFOwnerPass(ByVal value As Variant)
'        m_PDFOwnerPass = value
'End Property

Public Property Get PDFOwnerPasswordString() As Variant
       PDFOwnerPasswordString = m_PDFOwnerPasswordString
End Property

Public Property Let PDFOwnerPasswordString(ByVal value As Variant)
        m_PDFOwnerPasswordString = value
End Property

Public Property Get PDFPageLayout() As Variant
       PDFPageLayout = m_PDFPageLayout
End Property

Public Property Let PDFPageLayout(ByVal value As Variant)
        m_PDFPageLayout = value
End Property

Public Property Get PDFPageMode() As Variant
       PDFPageMode = m_PDFPageMode
End Property

Public Property Let PDFPageMode(ByVal value As Variant)
        m_PDFPageMode = value
End Property

Public Property Get PDFSigningMultiSignature() As Variant
       PDFSigningMultiSignature = m_PDFSigningMultiSignature
End Property

Public Property Let PDFSigningMultiSignature(ByVal value As Variant)
        m_PDFSigningMultiSignature = value
End Property

Public Property Get PDFSigningPFXFile() As Variant
       PDFSigningPFXFile = m_PDFSigningPFXFile
End Property

Public Property Let PDFSigningPFXFile(ByVal value As Variant)
        m_PDFSigningPFXFile = value
End Property

Public Property Get PDFSigningPFXFilePassword() As Variant
       PDFSigningPFXFilePassword = m_PDFSigningPFXFilePassword
End Property

Public Property Let PDFSigningPFXFilePassword(ByVal value As Variant)
        m_PDFSigningPFXFilePassword = value
End Property

Public Property Get PDFSigningSignatureContact() As Variant
       PDFSigningSignatureContact = m_PDFSigningSignatureContact
End Property

Public Property Let PDFSigningSignatureContact(ByVal value As Variant)
        m_PDFSigningSignatureContact = value
End Property

Public Property Get PDFSigningSignatureLeftX() As Variant
       PDFSigningSignatureLeftX = m_PDFSigningSignatureLeftX
End Property

Public Property Let PDFSigningSignatureLeftX(ByVal value As Variant)
        m_PDFSigningSignatureLeftX = value
End Property

Public Property Get PDFSigningSignatureLeftY() As Variant
       PDFSigningSignatureLeftY = m_PDFSigningSignatureLeftY
End Property

Public Property Let PDFSigningSignatureLeftY(ByVal value As Variant)
        m_PDFSigningSignatureLeftY = value
End Property

Public Property Get PDFSigningSignatureLocation() As Variant
       PDFSigningSignatureLocation = m_PDFSigningSignatureLocation
End Property

Public Property Let PDFSigningSignatureLocation(ByVal value As Variant)
        m_PDFSigningSignatureLocation = value
End Property

Public Property Get PDFSigningSignatureOnPage() As Variant
       PDFSigningSignatureOnPage = m_PDFSigningSignatureOnPage
End Property

Public Property Let PDFSigningSignatureOnPage(ByVal value As Variant)
        m_PDFSigningSignatureOnPage = value
End Property

Public Property Get PDFSigningSignatureReason() As Variant
       PDFSigningSignatureReason = m_PDFSigningSignatureReason
End Property

Public Property Let PDFSigningSignatureReason(ByVal value As Variant)
        m_PDFSigningSignatureReason = value
End Property

Public Property Get PDFSigningSignatureRightX() As Variant
       PDFSigningSignatureRightX = m_PDFSigningSignatureRightX
End Property

Public Property Let PDFSigningSignatureRightX(ByVal value As Variant)
        m_PDFSigningSignatureRightX = value
End Property

Public Property Get PDFSigningSignatureRightY() As Variant
       PDFSigningSignatureRightY = m_PDFSigningSignatureRightY
End Property

Public Property Let PDFSigningSignatureRightY(ByVal value As Variant)
        m_PDFSigningSignatureRightY = value
End Property

Public Property Get PDFSigningSignatureVisible() As Variant
       PDFSigningSignatureVisible = m_PDFSigningSignatureVisible
End Property

Public Property Let PDFSigningSignatureVisible(ByVal value As Variant)
        m_PDFSigningSignatureVisible = value
End Property

Public Property Get PDFSigningSignPDF() As Variant
       PDFSigningSignPDF = m_PDFSigningSignPDF
End Property

Public Property Let PDFSigningSignPDF(ByVal value As Variant)
        m_PDFSigningSignPDF = value
End Property

Public Property Get PDFSigningTimeServerUrl() As Variant
       PDFSigningTimeServerUrl = m_PDFSigningTimeServerUrl
End Property

Public Property Let PDFSigningTimeServerUrl(ByVal value As Variant)
        m_PDFSigningTimeServerUrl = value
End Property

Public Property Get PDFStartPage() As Variant
       PDFStartPage = m_PDFStartPage
End Property

Public Property Let PDFStartPage(ByVal value As Variant)
        m_PDFStartPage = value
End Property

Public Property Get PDFUpdateMetadata() As Variant
       PDFUpdateMetadata = m_PDFUpdateMetadata
End Property

Public Property Let PDFUpdateMetadata(ByVal value As Variant)
        m_PDFUpdateMetadata = value
End Property

Public Property Get PDFUserPass() As Variant
       PDFUserPass = m_PDFUserPass
End Property

Public Property Let PDFUserPass(ByVal value As Variant)
        m_PDFUserPass = value
End Property

Public Property Get PDFUserPasswordString() As Variant
       PDFUserPasswordString = m_PDFUserPasswordString
End Property

Public Property Let PDFUserPasswordString(ByVal value As Variant)
        m_PDFUserPasswordString = value
End Property

Public Property Get PDFUseSecurity() As Variant
       PDFUseSecurity = m_PDFUseSecurity
End Property

Public Property Let PDFUseSecurity(ByVal value As Variant)
        m_PDFUseSecurity = value
End Property

Public Property Get PNGColorscount() As Variant
       PNGColorscount = m_PNGColorscount
End Property

Public Property Let PNGColorscount(ByVal value As Variant)
        m_PNGColorscount = value
End Property

Public Property Get PNGResolution() As Variant
       PNGResolution = m_PNGResolution
End Property

Public Property Let PNGResolution(ByVal value As Variant)
        m_PNGResolution = value
End Property

Public Property Get PrintAfterSaving() As Variant
       PrintAfterSaving = m_PrintAfterSaving
End Property

Public Property Let PrintAfterSaving(ByVal value As Variant)
        m_PrintAfterSaving = value
End Property

Public Property Get PrintAfterSavingBitsPerPixel() As Variant
       PrintAfterSavingBitsPerPixel = m_PrintAfterSavingBitsPerPixel
End Property

Public Property Let PrintAfterSavingBitsPerPixel(ByVal value As Variant)
        m_PrintAfterSavingBitsPerPixel = value
End Property

Public Property Get PrintAfterSavingDuplex() As Variant
       PrintAfterSavingDuplex = m_PrintAfterSavingDuplex
End Property

Public Property Let PrintAfterSavingDuplex(ByVal value As Variant)
        m_PrintAfterSavingDuplex = value
End Property

Public Property Get PrintAfterSavingMaxResolution() As Variant
       PrintAfterSavingMaxResolution = m_PrintAfterSavingMaxResolution
End Property

Public Property Let PrintAfterSavingMaxResolution(ByVal value As Variant)
        m_PrintAfterSavingMaxResolution = value
End Property

Public Property Get PrintAfterSavingMaxResolutionEnabled() As Variant
       PrintAfterSavingMaxResolutionEnabled = m_PrintAfterSavingMaxResolutionEnabled
End Property

Public Property Let PrintAfterSavingMaxResolutionEnabled(ByVal value As Variant)
        m_PrintAfterSavingMaxResolutionEnabled = value
End Property

Public Property Get PrintAfterSavingNoCancel() As Variant
       PrintAfterSavingNoCancel = m_PrintAfterSavingNoCancel
End Property

Public Property Let PrintAfterSavingNoCancel(ByVal value As Variant)
        m_PrintAfterSavingNoCancel = value
End Property

Public Property Get PrintAfterSavingPrinter() As Variant
       PrintAfterSavingPrinter = m_PrintAfterSavingPrinter
End Property

Public Property Let PrintAfterSavingPrinter(ByVal value As Variant)
        m_PrintAfterSavingPrinter = value
End Property

Public Property Get PrintAfterSavingQueryUser() As Variant
       PrintAfterSavingQueryUser = m_PrintAfterSavingQueryUser
End Property

Public Property Let PrintAfterSavingQueryUser(ByVal value As Variant)
        m_PrintAfterSavingQueryUser = value
End Property

Public Property Get PrintAfterSavingTumble() As Variant
       PrintAfterSavingTumble = m_PrintAfterSavingTumble
End Property

Public Property Let PrintAfterSavingTumble(ByVal value As Variant)
        m_PrintAfterSavingTumble = value
End Property

Public Property Get PrinterStop() As Variant
       PrinterStop = m_PrinterStop
End Property

Public Property Let PrinterStop(ByVal value As Variant)
        m_PrinterStop = value
End Property

Public Property Get ProcessPriority() As Variant
       ProcessPriority = m_ProcessPriority
End Property

Public Property Let ProcessPriority(ByVal value As Variant)
        m_ProcessPriority = value
End Property

Public Property Get ProgramFont() As Variant
       ProgramFont = m_ProgramFont
End Property

Public Property Let ProgramFont(ByVal value As Variant)
        m_ProgramFont = value
End Property

Public Property Get ProgramFontCharset() As Variant
       ProgramFontCharset = m_ProgramFontCharset
End Property

Public Property Let ProgramFontCharset(ByVal value As Variant)
        m_ProgramFontCharset = value
End Property

Public Property Get ProgramFontSize() As Variant
       ProgramFontSize = m_ProgramFontSize
End Property

Public Property Let ProgramFontSize(ByVal value As Variant)
        m_ProgramFontSize = value
End Property

Public Property Get PSDColorsCount() As Variant
       PSDColorsCount = m_PSDColorsCount
End Property

Public Property Let PSDColorsCount(ByVal value As Variant)
        m_PSDColorsCount = value
End Property

Public Property Get PSDResolution() As Variant
       PSDResolution = m_PSDResolution
End Property

Public Property Let PSDResolution(ByVal value As Variant)
        m_PSDResolution = value
End Property

Public Property Get PSLanguageLevel() As Variant
       PSLanguageLevel = m_PSLanguageLevel
End Property

Public Property Let PSLanguageLevel(ByVal value As Variant)
        m_PSLanguageLevel = value
End Property

Public Property Get RAWColorsCount() As Variant
       RAWColorsCount = m_RAWColorsCount
End Property

Public Property Let RAWColorsCount(ByVal value As Variant)
        m_RAWColorsCount = value
End Property

Public Property Get RAWResolution() As Variant
       RAWResolution = m_RAWResolution
End Property

Public Property Let RAWResolution(ByVal value As Variant)
        m_RAWResolution = value
End Property

Public Property Get RemoveAllKnownFileExtensions() As Variant
       RemoveAllKnownFileExtensions = m_RemoveAllKnownFileExtensions
End Property

Public Property Let RemoveAllKnownFileExtensions(ByVal value As Variant)
        m_RemoveAllKnownFileExtensions = value
End Property

Public Property Get RemoveSpaces() As Variant
       RemoveSpaces = m_RemoveSpaces
End Property

Public Property Let RemoveSpaces(ByVal value As Variant)
        m_RemoveSpaces = value
End Property

Public Property Get RunProgramAfterSaving() As Variant
       RunProgramAfterSaving = m_RunProgramAfterSaving
End Property

Public Property Let RunProgramAfterSaving(ByVal value As Variant)
        m_RunProgramAfterSaving = value
End Property

Public Property Get RunProgramAfterSavingProgramname() As Variant
       RunProgramAfterSavingProgramname = m_RunProgramAfterSavingProgramname
End Property

Public Property Let RunProgramAfterSavingProgramname(ByVal value As Variant)
        m_RunProgramAfterSavingProgramname = value
End Property

Public Property Get RunProgramAfterSavingProgramParameters() As Variant
       RunProgramAfterSavingProgramParameters = m_RunProgramAfterSavingProgramParameters
End Property

Public Property Let RunProgramAfterSavingProgramParameters(ByVal value As Variant)
        m_RunProgramAfterSavingProgramParameters = value
End Property

Public Property Get RunProgramAfterSavingWaitUntilReady() As Variant
       RunProgramAfterSavingWaitUntilReady = m_RunProgramAfterSavingWaitUntilReady
End Property

Public Property Let RunProgramAfterSavingWaitUntilReady(ByVal value As Variant)
        m_RunProgramAfterSavingWaitUntilReady = value
End Property

Public Property Get RunProgramAfterSavingWindowstyle() As Variant
       RunProgramAfterSavingWindowstyle = m_RunProgramAfterSavingWindowstyle
End Property

Public Property Let RunProgramAfterSavingWindowstyle(ByVal value As Variant)
        m_RunProgramAfterSavingWindowstyle = value
End Property

Public Property Get RunProgramBeforeSaving() As Variant
       RunProgramBeforeSaving = m_RunProgramBeforeSaving
End Property

Public Property Let RunProgramBeforeSaving(ByVal value As Variant)
        m_RunProgramBeforeSaving = value
End Property

Public Property Get RunProgramBeforeSavingProgramname() As Variant
       RunProgramBeforeSavingProgramname = m_RunProgramBeforeSavingProgramname
End Property

Public Property Let RunProgramBeforeSavingProgramname(ByVal value As Variant)
        m_RunProgramBeforeSavingProgramname = value
End Property

Public Property Get RunProgramBeforeSavingProgramParameters() As Variant
       RunProgramBeforeSavingProgramParameters = m_RunProgramBeforeSavingProgramParameters
End Property

Public Property Let RunProgramBeforeSavingProgramParameters(ByVal value As Variant)
        m_RunProgramBeforeSavingProgramParameters = value
End Property

Public Property Get RunProgramBeforeSavingWindowstyle() As Variant
       RunProgramBeforeSavingWindowstyle = m_RunProgramBeforeSavingWindowstyle
End Property

Public Property Let RunProgramBeforeSavingWindowstyle(ByVal value As Variant)
        m_RunProgramBeforeSavingWindowstyle = value
End Property

Public Property Get SaveFilename() As Variant
       SaveFilename = m_SaveFilename
End Property

Public Property Let SaveFilename(ByVal value As Variant)
        m_SaveFilename = value
End Property

Public Property Get SendEmailAfterAutoSaving() As Variant
       SendEmailAfterAutoSaving = m_SendEmailAfterAutoSaving
End Property

Public Property Let SendEmailAfterAutoSaving(ByVal value As Variant)
        m_SendEmailAfterAutoSaving = value
End Property

Public Property Get SendMailMethod() As Variant
       SendMailMethod = m_SendMailMethod
End Property

Public Property Let SendMailMethod(ByVal value As Variant)
        m_SendMailMethod = value
End Property

Public Property Get ShowAnimation() As Variant
       ShowAnimation = m_ShowAnimation
End Property

Public Property Let ShowAnimation(ByVal value As Variant)
        m_ShowAnimation = value
End Property

Public Property Get StampFontColor() As Variant
       StampFontColor = m_StampFontColor
End Property

Public Property Let StampFontColor(ByVal value As Variant)
        m_StampFontColor = value
End Property

Public Property Get StampFontname() As Variant
       StampFontname = m_StampFontname
End Property

Public Property Let StampFontname(ByVal value As Variant)
        m_StampFontname = value
End Property

Public Property Get StampFontsize() As Variant
       StampFontsize = m_StampFontsize
End Property

Public Property Let StampFontsize(ByVal value As Variant)
        m_StampFontsize = value
End Property

Public Property Get StampOutlineFontthickness() As Variant
       StampOutlineFontthickness = m_StampOutlineFontthickness
End Property

Public Property Let StampOutlineFontthickness(ByVal value As Variant)
        m_StampOutlineFontthickness = value
End Property

Public Property Get StampString() As Variant
       StampString = m_StampString
End Property

Public Property Let StampString(ByVal value As Variant)
        m_StampString = value
End Property

Public Property Get StampUseOutlineFont() As Variant
       StampUseOutlineFont = m_StampUseOutlineFont
End Property

Public Property Let StampUseOutlineFont(ByVal value As Variant)
        m_StampUseOutlineFont = value
End Property

Public Property Get StandardAuthor() As Variant
       StandardAuthor = m_StandardAuthor
End Property

Public Property Let StandardAuthor(ByVal value As Variant)
        m_StandardAuthor = value
End Property

Public Property Get StandardCreationdate() As Variant
       StandardCreationdate = m_StandardCreationdate
End Property

Public Property Let StandardCreationdate(ByVal value As Variant)
        m_StandardCreationdate = value
End Property

Public Property Get StandardDateformat() As Variant
       StandardDateformat = m_StandardDateformat
End Property

Public Property Let StandardDateformat(ByVal value As Variant)
        m_StandardDateformat = value
End Property

Public Property Get StandardKeywords() As Variant
       StandardKeywords = m_StandardKeywords
End Property

Public Property Let StandardKeywords(ByVal value As Variant)
        m_StandardKeywords = value
End Property

Public Property Get StandardMailDomain() As Variant
       StandardMailDomain = m_StandardMailDomain
End Property

Public Property Let StandardMailDomain(ByVal value As Variant)
        m_StandardMailDomain = value
End Property

Public Property Get StandardModifydate() As Variant
       StandardModifydate = m_StandardModifydate
End Property

Public Property Let StandardModifydate(ByVal value As Variant)
        m_StandardModifydate = value
End Property

Public Property Get StandardSaveformat() As Variant
       StandardSaveformat = m_StandardSaveformat
End Property

Public Property Let StandardSaveformat(ByVal value As Variant)
        m_StandardSaveformat = value
End Property

Public Property Get StandardSubject() As Variant
       StandardSubject = m_StandardSubject
End Property

Public Property Let StandardSubject(ByVal value As Variant)
        m_StandardSubject = value
End Property

Public Property Get StandardTitle() As Variant
       StandardTitle = m_StandardTitle
End Property

Public Property Let StandardTitle(ByVal value As Variant)
        m_StandardTitle = value
End Property

Public Property Get SVGResolution() As Variant
       SVGResolution = m_SVGResolution
End Property

Public Property Let SVGResolution(ByVal value As Variant)
        m_SVGResolution = value
End Property

Public Property Get TIFFColorscount() As Variant
       TIFFColorscount = m_TIFFColorscount
End Property

Public Property Let TIFFColorscount(ByVal value As Variant)
        m_TIFFColorscount = value
End Property

Public Property Get TIFFResolution() As Variant
       TIFFResolution = m_TIFFResolution
End Property

Public Property Let TIFFResolution(ByVal value As Variant)
        m_TIFFResolution = value
End Property

Public Property Get Toolbars() As Variant
       Toolbars = m_Toolbars
End Property

Public Property Let Toolbars(ByVal value As Variant)
        m_Toolbars = value
End Property

Public Property Get UpdateInterval() As Long
       UpdateInterval = m_UpdateInterval
End Property

Public Property Let UpdateInterval(ByVal value As Long)
        m_UpdateInterval = value
End Property

Public Property Get UseAutosave() As Long
       UseAutosave = m_UseAutosave
End Property

Public Property Let UseAutosave(ByVal value As Long)
        m_UseAutosave = value
End Property

Public Property Get UseAutosaveDirectory() As Long
       UseAutosaveDirectory = m_UseAutosaveDirectory
End Property

Public Property Let UseAutosaveDirectory(ByVal value As Long)
        m_UseAutosaveDirectory = value
End Property

Public Property Get UseCreationDateNow() As Long
       UseCreationDateNow = m_UseCreationDateNow
End Property

Public Property Let UseCreationDateNow(ByVal value As Long)
        m_UseCreationDateNow = value
End Property

Public Property Get UseCustomPaperSize() As Long
       UseCustomPaperSize = m_UseCustomPaperSize
End Property

Public Property Let UseCustomPaperSize(ByVal value As Long)
        m_UseCustomPaperSize = value
End Property

Public Property Get UseFixPapersize() As Long
       UseFixPapersize = m_UseFixPapersize
End Property

Public Property Let UseFixPapersize(ByVal value As Long)
        m_UseFixPapersize = value
End Property

Public Property Get UseStandardAuthor() As Long
       UseStandardAuthor = m_UseStandardAuthor
End Property

Public Property Let UseStandardAuthor(ByVal value As Long)
        m_UseStandardAuthor = value
End Property
