VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_NEWCLASS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'###############################################################
'##                                                           ##
'##  %EXTENSION_KEY% durch Kennung der Erweiterung ersetzen!  ##
'##                                                           ##
'###############################################################
'
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_%EXTENSION_KEY% (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>AddApplicationHandlerExtension New ApplicationHandler_%EXTENSION_KEY%</code>
' </remarks>
'\ingroup templates
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/templates/ApplicationHandler_%EXTENSION_KEY%.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "%EXTENSION_KEY%"

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(objRef As ApplicationHandler)
   Set m_ApplicationHandler = objRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ExtensionReference As Object)

On Error GoTo HandleErr

   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtAppFile.m_ApplicationHandler_ExtensionLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
                     
On Error GoTo HandleErr

   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtAppFile.m_ApplicationHandler_ExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_%EXTENSION_KEY%
'---------------------------------------------------------------------------------------
'
Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   
On Error GoTo HandleErr

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName
      'case ...
	  
      Case Else 'Property wurde nicht erkannt
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
		 
   End Select

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "Ext%EXTENSION_KEY%.getExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

