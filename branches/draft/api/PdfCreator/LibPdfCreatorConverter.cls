VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibPdfCreatorConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibPdfCreatorConverter
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Umwandlung zwischen Lib- und COM-Objekten
' \ingroup PdfCreator
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/PdfCreator/LibPdfCreatorConverter.cls</file>
'  <use>api/PdfCreator/LibPdfCreatorOptions</use>
'  <use>api/PdfCreator/LibPdfCreatorError</use>
'  <test>_test/api/PdfCreator/LibPdfCreatorConverterTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Function ConvertPdfCreatorError_ToLib(ByRef pdfCreatorError As Object) As LibPdfCreatorError
    Dim Result As New LibPdfCreatorError
    
    If Not pdfCreatorError Is Nothing Then
        Result.Description = pdfCreatorError.Description
        Result.Number = pdfCreatorError.Number
    End If
    
    Set ConvertPdfCreatorError_ToLib = Result
    Set Result = Nothing
End Function

#If UseEarlyBinding Then
Public Function ConvertPdfCreatorError_ToCOM(ByRef pdfCreatorError As LibPdfCreatorError) As PDFCreator.clsPDFCreatorError
#Else
Public Function ConvertPdfCreatorError_ToCOM(ByRef pdfCreatorError As LibPdfCreatorError) As Object
#End If
    
    Dim Result As Object
    Set Result = LibPdfCreatorComFactory.clsPDFCreatorError
        Result.Description = pdfCreatorError.Description
        Result.Number = pdfCreatorError.Number
    
    Set ConvertPdfCreatorError_ToCOM = Result
    Set Result = Nothing
End Function

Public Function ConvertPdfCreatorOption_ToLib(ByRef pdfCreatorOption As Object) As LibPdfCreatorOptions
    
    Dim Result As New LibPdfCreatorOptions

        Result.AdditionalGhostscriptParameters = pdfCreatorOption.AdditionalGhostscriptParameters
        Result.AdditionalGhostscriptSearchpath = pdfCreatorOption.AdditionalGhostscriptSearchpath
        Result.AddWindowsFontpath = pdfCreatorOption.AddWindowsFontpath
        Result.AllowSpecialGSCharsInFilenames = pdfCreatorOption.AllowSpecialGSCharsInFilenames
        Result.AutosaveDirectory = pdfCreatorOption.AutosaveDirectory
        Result.AutosaveFilename = pdfCreatorOption.AutosaveFilename
        Result.AutosaveFormat = pdfCreatorOption.AutosaveFormat
        Result.AutosaveStartStandardProgram = pdfCreatorOption.AutosaveStartStandardProgram
        Result.BMPColorscount = pdfCreatorOption.BMPColorscount
        Result.BMPResolution = pdfCreatorOption.BMPResolution
        Result.ClientComputerResolveIPAddress = pdfCreatorOption.ClientComputerResolveIPAddress
        Result.Counter = pdfCreatorOption.Counter
        Result.DeviceHeightPoints = pdfCreatorOption.DeviceHeightPoints
        Result.DeviceWidthPoints = pdfCreatorOption.DeviceWidthPoints
        Result.DirectoryGhostscriptBinaries = pdfCreatorOption.DirectoryGhostscriptBinaries
        Result.DirectoryGhostscriptFonts = pdfCreatorOption.DirectoryGhostscriptFonts
        Result.DirectoryGhostscriptLibraries = pdfCreatorOption.DirectoryGhostscriptLibraries
        Result.DirectoryGhostscriptResource = pdfCreatorOption.DirectoryGhostscriptResource
        Result.DisableEmail = pdfCreatorOption.DisableEmail
        Result.DisableUpdateCheck = pdfCreatorOption.DisableUpdateCheck
        Result.DontUseDocumentSettings = pdfCreatorOption.DontUseDocumentSettings
        Result.EditWithPDFArchitect = pdfCreatorOption.EditWithPDFArchitect
        Result.EPSLanguageLevel = pdfCreatorOption.EPSLanguageLevel
        Result.FilenameSubstitutions = pdfCreatorOption.FilenameSubstitutions
        Result.FilenameSubstitutionsOnlyInTitle = pdfCreatorOption.FilenameSubstitutionsOnlyInTitle
        Result.JPEGColorscount = pdfCreatorOption.JPEGColorscount
        Result.JPEGQuality = pdfCreatorOption.JPEGQuality
        Result.JPEGResolution = pdfCreatorOption.JPEGResolution
        Result.Language = pdfCreatorOption.Language
        Result.LastSaveDirectory = pdfCreatorOption.LastSaveDirectory
        Result.LastUpdateCheck = pdfCreatorOption.LastUpdateCheck
        Result.Logging = pdfCreatorOption.Logging
        Result.LogLines = pdfCreatorOption.LogLines
        Result.MaximumCountOfPDFArchitectToolTip = pdfCreatorOption.MaximumCountOfPDFArchitectToolTip
        Result.NoConfirmMessageSwitchingDefaultprinter = pdfCreatorOption.NoConfirmMessageSwitchingDefaultprinter
        Result.NoProcessingAtStartup = pdfCreatorOption.NoProcessingAtStartup
        Result.NoPSCheck = pdfCreatorOption.NoPSCheck
        Result.OneFilePerPage = pdfCreatorOption.OneFilePerPage
        Result.OpenOutputFile = pdfCreatorOption.OpenOutputFile
        Result.OptionsDesign = pdfCreatorOption.OptionsDesign
        Result.OptionsEnabled = pdfCreatorOption.OptionsEnabled
        Result.OptionsVisible = pdfCreatorOption.OptionsVisible
        Result.Papersize = pdfCreatorOption.Papersize
        Result.PCLColorsCount = pdfCreatorOption.PCLColorsCount
        Result.PCLResolution = pdfCreatorOption.PCLResolution
        Result.PCXColorscount = pdfCreatorOption.PCXColorscount
        Result.PCXResolution = pdfCreatorOption.PCXResolution
        Result.PDFAes128Encryption = pdfCreatorOption.PDFAes128Encryption
        Result.PDFAllowAssembly = pdfCreatorOption.PDFAllowAssembly
        Result.PDFAllowDegradedPrinting = pdfCreatorOption.PDFAllowDegradedPrinting
        Result.PDFAllowFillIn = pdfCreatorOption.PDFAllowFillIn
        Result.PDFAllowScreenReaders = pdfCreatorOption.PDFAllowScreenReaders
        Result.PDFColorsCMYKToRGB = pdfCreatorOption.PDFColorsCMYKToRGB
        Result.PDFColorsColorModel = pdfCreatorOption.PDFColorsColorModel
        Result.PDFColorsPreserveHalftone = pdfCreatorOption.PDFColorsPreserveHalftone
        Result.PDFColorsPreserveOverprint = pdfCreatorOption.PDFColorsPreserveOverprint
        Result.PDFColorsPreserveTransfer = pdfCreatorOption.PDFColorsPreserveTransfer
        Result.PDFCompressionColorCompression = pdfCreatorOption.PDFCompressionColorCompression
        Result.PDFCompressionColorCompressionChoice = pdfCreatorOption.PDFCompressionColorCompressionChoice
        Result.PDFCompressionColorCompressionJPEGHighFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGHighFactor
        Result.PDFCompressionColorCompressionJPEGLowFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGLowFactor
        Result.PDFCompressionColorCompressionJPEGManualFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGManualFactor
        Result.PDFCompressionColorCompressionJPEGMaximumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMaximumFactor
        Result.PDFCompressionColorCompressionJPEGMediumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMediumFactor
        Result.PDFCompressionColorCompressionJPEGMinimumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMinimumFactor
        Result.PDFCompressionColorResample = pdfCreatorOption.PDFCompressionColorResample
        Result.PDFCompressionColorResampleChoice = pdfCreatorOption.PDFCompressionColorResampleChoice
        Result.PDFCompressionColorResolution = pdfCreatorOption.PDFCompressionColorResolution
        Result.PDFCompressionGreyCompression = pdfCreatorOption.PDFCompressionGreyCompression
        Result.PDFCompressionGreyCompressionChoice = pdfCreatorOption.PDFCompressionGreyCompressionChoice
        Result.PDFCompressionGreyCompressionJPEGHighFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGHighFactor
        Result.PDFCompressionGreyCompressionJPEGLowFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGLowFactor
        Result.PDFCompressionGreyCompressionJPEGManualFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGManualFactor
        Result.PDFCompressionGreyCompressionJPEGMaximumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMaximumFactor
        Result.PDFCompressionGreyCompressionJPEGMediumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMediumFactor
        Result.PDFCompressionGreyCompressionJPEGMinimumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMinimumFactor
        Result.PDFCompressionGreyResample = pdfCreatorOption.PDFCompressionGreyResample
        Result.PDFCompressionGreyResampleChoice = pdfCreatorOption.PDFCompressionGreyResampleChoice
        Result.PDFCompressionGreyResolution = pdfCreatorOption.PDFCompressionGreyResolution
        Result.PDFCompressionMonoCompression = pdfCreatorOption.PDFCompressionMonoCompression
        Result.PDFCompressionMonoCompressionChoice = pdfCreatorOption.PDFCompressionMonoCompressionChoice
        Result.PDFCompressionMonoResample = pdfCreatorOption.PDFCompressionMonoResample
        Result.PDFCompressionMonoResampleChoice = pdfCreatorOption.PDFCompressionMonoResampleChoice
        Result.PDFCompressionMonoResolution = pdfCreatorOption.PDFCompressionMonoResolution
        Result.PDFCompressionTextCompression = pdfCreatorOption.PDFCompressionTextCompression
        Result.PDFDisallowCopy = pdfCreatorOption.PDFDisallowCopy
        Result.PDFDisallowModifyAnnotations = pdfCreatorOption.PDFDisallowModifyAnnotations
        Result.PDFDisallowModifyContents = pdfCreatorOption.PDFDisallowModifyContents
        Result.PDFDisallowPrinting = pdfCreatorOption.PDFDisallowPrinting
        Result.PDFEncryptor = pdfCreatorOption.PDFEncryptor
        Result.PDFFontsEmbedAll = pdfCreatorOption.PDFFontsEmbedAll
        Result.PDFFontsSubSetFonts = pdfCreatorOption.PDFFontsSubSetFonts
        Result.PDFFontsSubSetFontsPercent = pdfCreatorOption.PDFFontsSubSetFontsPercent
        Result.PDFGeneralASCII85 = pdfCreatorOption.PDFGeneralASCII85
        Result.PDFGeneralAutorotate = pdfCreatorOption.PDFGeneralAutorotate
        Result.PDFGeneralCompatibility = pdfCreatorOption.PDFGeneralCompatibility
        Result.PDFGeneralDefault = pdfCreatorOption.PDFGeneralDefault
        Result.PDFGeneralOverprint = pdfCreatorOption.PDFGeneralOverprint
        Result.PDFHighEncryption = pdfCreatorOption.PDFHighEncryption
        Result.PDFLowEncryption = pdfCreatorOption.PDFLowEncryption
        Result.PDFOptimize = pdfCreatorOption.PDFOptimize
        'Result.PDFOwnerPass = pdfCreatorOption.PDFOwnerPass
        Result.PDFOwnerPasswordString = pdfCreatorOption.PDFOwnerPasswordString
        Result.PDFPageLayout = pdfCreatorOption.PDFPageLayout
        Result.PDFPageMode = pdfCreatorOption.PDFPageMode
        Result.PDFSigningMultiSignature = pdfCreatorOption.PDFSigningMultiSignature
        Result.PDFSigningPFXFile = pdfCreatorOption.PDFSigningPFXFile
        Result.PDFSigningPFXFilePassword = pdfCreatorOption.PDFSigningPFXFilePassword
        Result.PDFSigningSignatureContact = pdfCreatorOption.PDFSigningSignatureContact
        Result.PDFSigningSignatureLeftX = pdfCreatorOption.PDFSigningSignatureLeftX
        Result.PDFSigningSignatureLeftY = pdfCreatorOption.PDFSigningSignatureLeftY
        Result.PDFSigningSignatureLocation = pdfCreatorOption.PDFSigningSignatureLocation
        Result.PDFSigningSignatureOnPage = pdfCreatorOption.PDFSigningSignatureOnPage
        Result.PDFSigningSignatureReason = pdfCreatorOption.PDFSigningSignatureReason
        Result.PDFSigningSignatureRightX = pdfCreatorOption.PDFSigningSignatureRightX
        Result.PDFSigningSignatureRightY = pdfCreatorOption.PDFSigningSignatureRightY
        Result.PDFSigningSignatureVisible = pdfCreatorOption.PDFSigningSignatureVisible
        Result.PDFSigningSignPDF = pdfCreatorOption.PDFSigningSignPDF
        Result.PDFSigningTimeServerUrl = pdfCreatorOption.PDFSigningTimeServerUrl
        Result.PDFStartPage = pdfCreatorOption.PDFStartPage
        Result.PDFUpdateMetadata = pdfCreatorOption.PDFUpdateMetadata
        Result.PDFUserPass = pdfCreatorOption.PDFUserPass
        Result.PDFUserPasswordString = pdfCreatorOption.PDFUserPasswordString
        Result.PDFUseSecurity = pdfCreatorOption.PDFUseSecurity
        Result.PNGColorscount = pdfCreatorOption.PNGColorscount
        Result.PNGResolution = pdfCreatorOption.PNGResolution
        Result.PrintAfterSaving = pdfCreatorOption.PrintAfterSaving
        Result.PrintAfterSavingBitsPerPixel = pdfCreatorOption.PrintAfterSavingBitsPerPixel
        Result.PrintAfterSavingDuplex = pdfCreatorOption.PrintAfterSavingDuplex
        Result.PrintAfterSavingMaxResolution = pdfCreatorOption.PrintAfterSavingMaxResolution
        Result.PrintAfterSavingMaxResolutionEnabled = pdfCreatorOption.PrintAfterSavingMaxResolutionEnabled
        Result.PrintAfterSavingNoCancel = pdfCreatorOption.PrintAfterSavingNoCancel
        Result.PrintAfterSavingPrinter = pdfCreatorOption.PrintAfterSavingPrinter
        Result.PrintAfterSavingQueryUser = pdfCreatorOption.PrintAfterSavingQueryUser
        Result.PrintAfterSavingTumble = pdfCreatorOption.PrintAfterSavingTumble
        Result.PrinterStop = pdfCreatorOption.PrinterStop
        Result.ProcessPriority = pdfCreatorOption.ProcessPriority
        Result.ProgramFont = pdfCreatorOption.ProgramFont
        Result.ProgramFontCharset = pdfCreatorOption.ProgramFontCharset
        Result.ProgramFontSize = pdfCreatorOption.ProgramFontSize
        Result.PSDColorsCount = pdfCreatorOption.PSDColorsCount
        Result.PSDResolution = pdfCreatorOption.PSDResolution
        Result.PSLanguageLevel = pdfCreatorOption.PSLanguageLevel
        Result.RAWColorsCount = pdfCreatorOption.RAWColorsCount
        Result.RAWResolution = pdfCreatorOption.RAWResolution
        Result.RemoveAllKnownFileExtensions = pdfCreatorOption.RemoveAllKnownFileExtensions
        Result.RemoveSpaces = pdfCreatorOption.RemoveSpaces
        Result.RunProgramAfterSaving = pdfCreatorOption.RunProgramAfterSaving
        Result.RunProgramAfterSavingProgramname = pdfCreatorOption.RunProgramAfterSavingProgramname
        Result.RunProgramAfterSavingProgramParameters = pdfCreatorOption.RunProgramAfterSavingProgramParameters
        Result.RunProgramAfterSavingWaitUntilReady = pdfCreatorOption.RunProgramAfterSavingWaitUntilReady
        Result.RunProgramAfterSavingWindowstyle = pdfCreatorOption.RunProgramAfterSavingWindowstyle
        Result.RunProgramBeforeSaving = pdfCreatorOption.RunProgramBeforeSaving
        Result.RunProgramBeforeSavingProgramname = pdfCreatorOption.RunProgramBeforeSavingProgramname
        Result.RunProgramBeforeSavingProgramParameters = pdfCreatorOption.RunProgramBeforeSavingProgramParameters
        Result.RunProgramBeforeSavingWindowstyle = pdfCreatorOption.RunProgramBeforeSavingWindowstyle
        Result.SaveFilename = pdfCreatorOption.SaveFilename
        Result.SendEmailAfterAutoSaving = pdfCreatorOption.SendEmailAfterAutoSaving
        Result.SendMailMethod = pdfCreatorOption.SendMailMethod
        Result.ShowAnimation = pdfCreatorOption.ShowAnimation
        Result.StampFontColor = pdfCreatorOption.StampFontColor
        Result.StampFontname = pdfCreatorOption.StampFontname
        Result.StampFontsize = pdfCreatorOption.StampFontsize
        Result.StampOutlineFontthickness = pdfCreatorOption.StampOutlineFontthickness
        Result.StampString = pdfCreatorOption.StampString
        Result.StampUseOutlineFont = pdfCreatorOption.StampUseOutlineFont
        Result.StandardAuthor = pdfCreatorOption.StandardAuthor
        Result.StandardCreationdate = pdfCreatorOption.StandardCreationdate
        Result.StandardDateformat = pdfCreatorOption.StandardDateformat
        Result.StandardKeywords = pdfCreatorOption.StandardKeywords
        Result.StandardMailDomain = pdfCreatorOption.StandardMailDomain
        Result.StandardModifydate = pdfCreatorOption.StandardModifydate
        Result.StandardSaveformat = pdfCreatorOption.StandardSaveformat
        Result.StandardSubject = pdfCreatorOption.StandardSubject
        Result.StandardTitle = pdfCreatorOption.StandardTitle
        Result.SVGResolution = pdfCreatorOption.SVGResolution
        Result.TIFFColorscount = pdfCreatorOption.TIFFColorscount
        Result.TIFFResolution = pdfCreatorOption.TIFFResolution
        Result.Toolbars = pdfCreatorOption.Toolbars
        Result.UpdateInterval = pdfCreatorOption.UpdateInterval
        Result.UseAutosave = pdfCreatorOption.UseAutosave
        Result.UseAutosaveDirectory = pdfCreatorOption.UseAutosaveDirectory
        Result.UseCreationDateNow = pdfCreatorOption.UseCreationDateNow
        Result.UseCustomPaperSize = pdfCreatorOption.UseCustomPaperSize
        Result.UseFixPapersize = pdfCreatorOption.UseFixPapersize
        Result.UseStandardAuthor = pdfCreatorOption.UseStandardAuthor

    
    Set ConvertPdfCreatorOption_ToLib = Result
    Set Result = Nothing

End Function

#If UseEarlyBinding Then
Public Function ConvertPdfCreatorOption_ToCOM(ByRef pdfCreatorOption As LibPdfCreatorOptions) As PDFCreator.clsPDFCreatorOptions
#Else
Public Function ConvertPdfCreatorOption_ToCOM(ByRef pdfCreatorOption As LibPdfCreatorOptions) As Object
#End If

    Dim Result As Object
    Set Result = LibPdfCreatorComFactory.clsPDFCreatorOptions

        Result.AdditionalGhostscriptParameters = pdfCreatorOption.AdditionalGhostscriptParameters
        Result.AdditionalGhostscriptSearchpath = pdfCreatorOption.AdditionalGhostscriptSearchpath
        Result.AddWindowsFontpath = pdfCreatorOption.AddWindowsFontpath
        Result.AllowSpecialGSCharsInFilenames = pdfCreatorOption.AllowSpecialGSCharsInFilenames
        Result.AutosaveDirectory = pdfCreatorOption.AutosaveDirectory
        Result.AutosaveFilename = pdfCreatorOption.AutosaveFilename
        Result.AutosaveFormat = pdfCreatorOption.AutosaveFormat
        Result.AutosaveStartStandardProgram = pdfCreatorOption.AutosaveStartStandardProgram
        Result.BMPColorscount = pdfCreatorOption.BMPColorscount
        Result.BMPResolution = pdfCreatorOption.BMPResolution
        Result.ClientComputerResolveIPAddress = pdfCreatorOption.ClientComputerResolveIPAddress
        Result.Counter = pdfCreatorOption.Counter
        Result.DeviceHeightPoints = pdfCreatorOption.DeviceHeightPoints
        Result.DeviceWidthPoints = pdfCreatorOption.DeviceWidthPoints
        Result.DirectoryGhostscriptBinaries = pdfCreatorOption.DirectoryGhostscriptBinaries
        Result.DirectoryGhostscriptFonts = pdfCreatorOption.DirectoryGhostscriptFonts
        Result.DirectoryGhostscriptLibraries = pdfCreatorOption.DirectoryGhostscriptLibraries
        Result.DirectoryGhostscriptResource = pdfCreatorOption.DirectoryGhostscriptResource
        Result.DisableEmail = pdfCreatorOption.DisableEmail
        Result.DisableUpdateCheck = pdfCreatorOption.DisableUpdateCheck
        Result.DontUseDocumentSettings = pdfCreatorOption.DontUseDocumentSettings
        Result.EditWithPDFArchitect = pdfCreatorOption.EditWithPDFArchitect
        Result.EPSLanguageLevel = pdfCreatorOption.EPSLanguageLevel
        Result.FilenameSubstitutions = pdfCreatorOption.FilenameSubstitutions
        Result.FilenameSubstitutionsOnlyInTitle = pdfCreatorOption.FilenameSubstitutionsOnlyInTitle
        Result.JPEGColorscount = pdfCreatorOption.JPEGColorscount
        Result.JPEGQuality = pdfCreatorOption.JPEGQuality
        Result.JPEGResolution = pdfCreatorOption.JPEGResolution
        Result.Language = pdfCreatorOption.Language
        Result.LastSaveDirectory = pdfCreatorOption.LastSaveDirectory
        Result.LastUpdateCheck = pdfCreatorOption.LastUpdateCheck
        Result.Logging = pdfCreatorOption.Logging
        Result.LogLines = pdfCreatorOption.LogLines
        Result.MaximumCountOfPDFArchitectToolTip = pdfCreatorOption.MaximumCountOfPDFArchitectToolTip
        Result.NoConfirmMessageSwitchingDefaultprinter = pdfCreatorOption.NoConfirmMessageSwitchingDefaultprinter
        Result.NoProcessingAtStartup = pdfCreatorOption.NoProcessingAtStartup
        Result.NoPSCheck = pdfCreatorOption.NoPSCheck
        Result.OneFilePerPage = pdfCreatorOption.OneFilePerPage
        Result.OpenOutputFile = pdfCreatorOption.OpenOutputFile
        Result.OptionsDesign = pdfCreatorOption.OptionsDesign
        Result.OptionsEnabled = pdfCreatorOption.OptionsEnabled
        Result.OptionsVisible = pdfCreatorOption.OptionsVisible
        Result.Papersize = pdfCreatorOption.Papersize
        Result.PCLColorsCount = pdfCreatorOption.PCLColorsCount
        Result.PCLResolution = pdfCreatorOption.PCLResolution
        Result.PCXColorscount = pdfCreatorOption.PCXColorscount
        Result.PCXResolution = pdfCreatorOption.PCXResolution
        Result.PDFAes128Encryption = pdfCreatorOption.PDFAes128Encryption
        Result.PDFAllowAssembly = pdfCreatorOption.PDFAllowAssembly
        Result.PDFAllowDegradedPrinting = pdfCreatorOption.PDFAllowDegradedPrinting
        Result.PDFAllowFillIn = pdfCreatorOption.PDFAllowFillIn
        Result.PDFAllowScreenReaders = pdfCreatorOption.PDFAllowScreenReaders
        Result.PDFColorsCMYKToRGB = pdfCreatorOption.PDFColorsCMYKToRGB
        Result.PDFColorsColorModel = pdfCreatorOption.PDFColorsColorModel
        Result.PDFColorsPreserveHalftone = pdfCreatorOption.PDFColorsPreserveHalftone
        Result.PDFColorsPreserveOverprint = pdfCreatorOption.PDFColorsPreserveOverprint
        Result.PDFColorsPreserveTransfer = pdfCreatorOption.PDFColorsPreserveTransfer
        Result.PDFCompressionColorCompression = pdfCreatorOption.PDFCompressionColorCompression
        Result.PDFCompressionColorCompressionChoice = pdfCreatorOption.PDFCompressionColorCompressionChoice
        Result.PDFCompressionColorCompressionJPEGHighFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGHighFactor
        Result.PDFCompressionColorCompressionJPEGLowFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGLowFactor
        Result.PDFCompressionColorCompressionJPEGManualFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGManualFactor
        Result.PDFCompressionColorCompressionJPEGMaximumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMaximumFactor
        Result.PDFCompressionColorCompressionJPEGMediumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMediumFactor
        Result.PDFCompressionColorCompressionJPEGMinimumFactor = pdfCreatorOption.PDFCompressionColorCompressionJPEGMinimumFactor
        Result.PDFCompressionColorResample = pdfCreatorOption.PDFCompressionColorResample
        Result.PDFCompressionColorResampleChoice = pdfCreatorOption.PDFCompressionColorResampleChoice
        Result.PDFCompressionColorResolution = pdfCreatorOption.PDFCompressionColorResolution
        Result.PDFCompressionGreyCompression = pdfCreatorOption.PDFCompressionGreyCompression
        Result.PDFCompressionGreyCompressionChoice = pdfCreatorOption.PDFCompressionGreyCompressionChoice
        Result.PDFCompressionGreyCompressionJPEGHighFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGHighFactor
        Result.PDFCompressionGreyCompressionJPEGLowFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGLowFactor
        Result.PDFCompressionGreyCompressionJPEGManualFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGManualFactor
        Result.PDFCompressionGreyCompressionJPEGMaximumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMaximumFactor
        Result.PDFCompressionGreyCompressionJPEGMediumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMediumFactor
        Result.PDFCompressionGreyCompressionJPEGMinimumFactor = pdfCreatorOption.PDFCompressionGreyCompressionJPEGMinimumFactor
        Result.PDFCompressionGreyResample = pdfCreatorOption.PDFCompressionGreyResample
        Result.PDFCompressionGreyResampleChoice = pdfCreatorOption.PDFCompressionGreyResampleChoice
        Result.PDFCompressionGreyResolution = pdfCreatorOption.PDFCompressionGreyResolution
        Result.PDFCompressionMonoCompression = pdfCreatorOption.PDFCompressionMonoCompression
        Result.PDFCompressionMonoCompressionChoice = pdfCreatorOption.PDFCompressionMonoCompressionChoice
        Result.PDFCompressionMonoResample = pdfCreatorOption.PDFCompressionMonoResample
        Result.PDFCompressionMonoResampleChoice = pdfCreatorOption.PDFCompressionMonoResampleChoice
        Result.PDFCompressionMonoResolution = pdfCreatorOption.PDFCompressionMonoResolution
        Result.PDFCompressionTextCompression = pdfCreatorOption.PDFCompressionTextCompression
        Result.PDFDisallowCopy = pdfCreatorOption.PDFDisallowCopy
        Result.PDFDisallowModifyAnnotations = pdfCreatorOption.PDFDisallowModifyAnnotations
        Result.PDFDisallowModifyContents = pdfCreatorOption.PDFDisallowModifyContents
        Result.PDFDisallowPrinting = pdfCreatorOption.PDFDisallowPrinting
        Result.PDFEncryptor = pdfCreatorOption.PDFEncryptor
        Result.PDFFontsEmbedAll = pdfCreatorOption.PDFFontsEmbedAll
        Result.PDFFontsSubSetFonts = pdfCreatorOption.PDFFontsSubSetFonts
        Result.PDFFontsSubSetFontsPercent = pdfCreatorOption.PDFFontsSubSetFontsPercent
        Result.PDFGeneralASCII85 = pdfCreatorOption.PDFGeneralASCII85
        Result.PDFGeneralAutorotate = pdfCreatorOption.PDFGeneralAutorotate
        Result.PDFGeneralCompatibility = pdfCreatorOption.PDFGeneralCompatibility
        Result.PDFGeneralDefault = pdfCreatorOption.PDFGeneralDefault
        Result.PDFGeneralOverprint = pdfCreatorOption.PDFGeneralOverprint
        Result.PDFHighEncryption = pdfCreatorOption.PDFHighEncryption
        Result.PDFLowEncryption = pdfCreatorOption.PDFLowEncryption
        Result.PDFOptimize = pdfCreatorOption.PDFOptimize
        'Result.PDFOwnerPass = pdfCreatorOption.PDFOwnerPass
        Result.PDFOwnerPasswordString = pdfCreatorOption.PDFOwnerPasswordString
        Result.PDFPageLayout = pdfCreatorOption.PDFPageLayout
        Result.PDFPageMode = pdfCreatorOption.PDFPageMode
        Result.PDFSigningMultiSignature = pdfCreatorOption.PDFSigningMultiSignature
        Result.PDFSigningPFXFile = pdfCreatorOption.PDFSigningPFXFile
        Result.PDFSigningPFXFilePassword = pdfCreatorOption.PDFSigningPFXFilePassword
        Result.PDFSigningSignatureContact = pdfCreatorOption.PDFSigningSignatureContact
        Result.PDFSigningSignatureLeftX = pdfCreatorOption.PDFSigningSignatureLeftX
        Result.PDFSigningSignatureLeftY = pdfCreatorOption.PDFSigningSignatureLeftY
        Result.PDFSigningSignatureLocation = pdfCreatorOption.PDFSigningSignatureLocation
        Result.PDFSigningSignatureOnPage = pdfCreatorOption.PDFSigningSignatureOnPage
        Result.PDFSigningSignatureReason = pdfCreatorOption.PDFSigningSignatureReason
        Result.PDFSigningSignatureRightX = pdfCreatorOption.PDFSigningSignatureRightX
        Result.PDFSigningSignatureRightY = pdfCreatorOption.PDFSigningSignatureRightY
        Result.PDFSigningSignatureVisible = pdfCreatorOption.PDFSigningSignatureVisible
        Result.PDFSigningSignPDF = pdfCreatorOption.PDFSigningSignPDF
        Result.PDFSigningTimeServerUrl = pdfCreatorOption.PDFSigningTimeServerUrl
        Result.PDFStartPage = pdfCreatorOption.PDFStartPage
        Result.PDFUpdateMetadata = pdfCreatorOption.PDFUpdateMetadata
        Result.PDFUserPass = pdfCreatorOption.PDFUserPass
        Result.PDFUserPasswordString = pdfCreatorOption.PDFUserPasswordString
        Result.PDFUseSecurity = pdfCreatorOption.PDFUseSecurity
        Result.PNGColorscount = pdfCreatorOption.PNGColorscount
        Result.PNGResolution = pdfCreatorOption.PNGResolution
        Result.PrintAfterSaving = pdfCreatorOption.PrintAfterSaving
        Result.PrintAfterSavingBitsPerPixel = pdfCreatorOption.PrintAfterSavingBitsPerPixel
        Result.PrintAfterSavingDuplex = pdfCreatorOption.PrintAfterSavingDuplex
        Result.PrintAfterSavingMaxResolution = pdfCreatorOption.PrintAfterSavingMaxResolution
        Result.PrintAfterSavingMaxResolutionEnabled = pdfCreatorOption.PrintAfterSavingMaxResolutionEnabled
        Result.PrintAfterSavingNoCancel = pdfCreatorOption.PrintAfterSavingNoCancel
        Result.PrintAfterSavingPrinter = pdfCreatorOption.PrintAfterSavingPrinter
        Result.PrintAfterSavingQueryUser = pdfCreatorOption.PrintAfterSavingQueryUser
        Result.PrintAfterSavingTumble = pdfCreatorOption.PrintAfterSavingTumble
        Result.PrinterStop = pdfCreatorOption.PrinterStop
        Result.ProcessPriority = pdfCreatorOption.ProcessPriority
        Result.ProgramFont = pdfCreatorOption.ProgramFont
        Result.ProgramFontCharset = pdfCreatorOption.ProgramFontCharset
        Result.ProgramFontSize = pdfCreatorOption.ProgramFontSize
        Result.PSDColorsCount = pdfCreatorOption.PSDColorsCount
        Result.PSDResolution = pdfCreatorOption.PSDResolution
        Result.PSLanguageLevel = pdfCreatorOption.PSLanguageLevel
        Result.RAWColorsCount = pdfCreatorOption.RAWColorsCount
        Result.RAWResolution = pdfCreatorOption.RAWResolution
        Result.RemoveAllKnownFileExtensions = pdfCreatorOption.RemoveAllKnownFileExtensions
        Result.RemoveSpaces = pdfCreatorOption.RemoveSpaces
        Result.RunProgramAfterSaving = pdfCreatorOption.RunProgramAfterSaving
        Result.RunProgramAfterSavingProgramname = pdfCreatorOption.RunProgramAfterSavingProgramname
        Result.RunProgramAfterSavingProgramParameters = pdfCreatorOption.RunProgramAfterSavingProgramParameters
        Result.RunProgramAfterSavingWaitUntilReady = pdfCreatorOption.RunProgramAfterSavingWaitUntilReady
        Result.RunProgramAfterSavingWindowstyle = pdfCreatorOption.RunProgramAfterSavingWindowstyle
        Result.RunProgramBeforeSaving = pdfCreatorOption.RunProgramBeforeSaving
        Result.RunProgramBeforeSavingProgramname = pdfCreatorOption.RunProgramBeforeSavingProgramname
        Result.RunProgramBeforeSavingProgramParameters = pdfCreatorOption.RunProgramBeforeSavingProgramParameters
        Result.RunProgramBeforeSavingWindowstyle = pdfCreatorOption.RunProgramBeforeSavingWindowstyle
        Result.SaveFilename = pdfCreatorOption.SaveFilename
        Result.SendEmailAfterAutoSaving = pdfCreatorOption.SendEmailAfterAutoSaving
        Result.SendMailMethod = pdfCreatorOption.SendMailMethod
        Result.ShowAnimation = pdfCreatorOption.ShowAnimation
        Result.StampFontColor = pdfCreatorOption.StampFontColor
        Result.StampFontname = pdfCreatorOption.StampFontname
        Result.StampFontsize = pdfCreatorOption.StampFontsize
        Result.StampOutlineFontthickness = pdfCreatorOption.StampOutlineFontthickness
        Result.StampString = pdfCreatorOption.StampString
        Result.StampUseOutlineFont = pdfCreatorOption.StampUseOutlineFont
        Result.StandardAuthor = pdfCreatorOption.StandardAuthor
        Result.StandardCreationdate = pdfCreatorOption.StandardCreationdate
        Result.StandardDateformat = pdfCreatorOption.StandardDateformat
        Result.StandardKeywords = pdfCreatorOption.StandardKeywords
        Result.StandardMailDomain = pdfCreatorOption.StandardMailDomain
        Result.StandardModifydate = pdfCreatorOption.StandardModifydate
        Result.StandardSaveformat = pdfCreatorOption.StandardSaveformat
        Result.StandardSubject = pdfCreatorOption.StandardSubject
        Result.StandardTitle = pdfCreatorOption.StandardTitle
        Result.SVGResolution = pdfCreatorOption.SVGResolution
        Result.TIFFColorscount = pdfCreatorOption.TIFFColorscount
        Result.TIFFResolution = pdfCreatorOption.TIFFResolution
        Result.Toolbars = pdfCreatorOption.Toolbars
        Result.UpdateInterval = pdfCreatorOption.UpdateInterval
        Result.UseAutosave = pdfCreatorOption.UseAutosave
        Result.UseAutosaveDirectory = pdfCreatorOption.UseAutosaveDirectory
        Result.UseCreationDateNow = pdfCreatorOption.UseCreationDateNow
        Result.UseCustomPaperSize = pdfCreatorOption.UseCustomPaperSize
        Result.UseFixPapersize = pdfCreatorOption.UseFixPapersize
        Result.UseStandardAuthor = pdfCreatorOption.UseStandardAuthor

    
    Set ConvertPdfCreatorOption_ToCOM = Result
    Set Result = Nothing

End Function
