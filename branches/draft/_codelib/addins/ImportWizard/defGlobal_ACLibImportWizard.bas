Attribute VB_Name = "defGlobal_ACLibImportWizard"
'---------------------------------------------------------------------------------------
' Modul: AcLib_defGlobal (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beispiel für Anwendungskonfiguration
' </summary>
' <remarks>
' Indiviuell gestaltete Config-Module nicht in das Repositiory laden.
' </remarks>
' \ingroup ACLibAddInImportWizard
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/ImportWizard/defGlobal_ACLibImportWizard.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/ImportWizard/ACLibFileManager.cls</use>
'  <use>_codelib/addins/shared/ACLibConfiguration.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_ACLibFilemanager As String = "ACLibFileManager"
Private Const EXTENSION_KEY_ACLibConfiguration As String = "ACLibConfiguration"


Public Enum CodeLibElementType  'angelehnt an Enum vbext_ComponentType
   clet_StdModule = 1           ' = vbext_ComponentType.vbext_ct_StdModule
   clet_ClassModule = 2         ' = vbext_ComponentType.vbext_ct_ClassModule
   clet_Form = 101              ' = vbext_ComponentType.vbext_ct_Document + 1
   clet_Report = 102            ' = vbext_ComponentType.vbext_ct_Document + 2
End Enum

Public Enum CodeLibImportMode
   clim_ImportMissingItems = 0  ' überschreibt keine vorhandene Access-Objekte in der Anwendung
   clim_ImportSelectedOnly = 1  ' nur die ausgewählte Datei wird importiert (keine Abhängigkeistprüfung)
   clim_ImportAllUsedItems = 2  ' auch vorhandene Access-Objekte werden überschrieben
End Enum

Public Type CodeLibInfoReference
   Name As String
   Major As Long
   Minor As Long
   GUID As String
End Type

Public Type CodeLibInfo
   Name As String
   Type As CodeLibElementType
   RepositoryFile As String
   LocalFile As String
   RepositoryFileReplacement As String
   Dependency() As String
   References() As CodeLibInfoReference
   TestFiles() As String
   LicenseFile As String
End Type


'Standard-Icon
Public ACLibIconFileName As String 'Nur Dateiname inkl. Dateierweiterung, aber ohne vollständigen Pfad

Public Property Get CurrentACLibFileManager() As ACLibFileManager

On Error GoTo HandleErr

   Set CurrentACLibFileManager = CurrentApplication.Extensions(EXTENSION_KEY_ACLibFilemanager)

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentACLibFileManager", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Public Property Get CurrentACLibConfiguration() As ACLibConfiguration

On Error GoTo HandleErr

   Set CurrentACLibConfiguration = CurrentApplication.Extensions(EXTENSION_KEY_ACLibConfiguration)

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentACLibConfiguration", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

'---------------------------------------------------------------------------------------
' Function: GetACLibFileManager
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die Filemanager-Referenz nach außen weiter
' </summary>
' <returns>ACLibFileManager</returns>
' <remarks>
' Über diese Function können andere Add-Ins oder die Anwendung
' auf den Filemanager des Import-Wizard zugreifen.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetACLibFileManager() As ACLibFileManager

On Error GoTo HandleErr

   Set GetACLibFileManager = CurrentACLibFileManager

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetACLibFileManager", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Public Function RefreshAllCodeLibAccessObjects( _
         Optional ByVal ImportMode As CodeLibImportMode = CodeLibImportMode.clim_ImportAllUsedItems) As Variant
On Error GoTo HandleErr

   CurrentACLibFileManager.RefreshAll ImportMode, True

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "RefreshAllCodeLibAccessObjects", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Function

Public Function RefreshAllCodeLibAccessModules( _
         Optional ByVal ImportMode As CodeLibImportMode = CodeLibImportMode.clim_ImportAllUsedItems) As Variant
         
On Error GoTo HandleErr

   CurrentACLibFileManager.RefreshAllModules ImportMode, True

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "RefreshAllCodeLibAccessModules", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Function

Public Function ExportAllCodeLibElements() As Variant

On Error GoTo HandleErr

   CurrentACLibFileManager.ExportAll
   MsgBox "Export abgeschlossen", vbInformation

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ExportAllCodeLibElements", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function
