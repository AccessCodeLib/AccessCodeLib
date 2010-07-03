VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_Tracing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_Tracing (Josef Pötzl, 2010-07-02)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse: Ablaufverfolgung
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_Tracing</code>
' </remarks>
'\ingroup base.log
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/log/ApplicationHandler_Tracing.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <use>base/ApplicationHandler.cls</use>
'  <test>_test/base/log/ApplicationHandler_TracingTest.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "Tracing"

Public Enum TracingOutputMode
   [_TracingOutputMode_EventOnly] = 0
   TracingOutputMode_Debug = 1
   TracingOutputMode_Memory = 2
   TracingOutputMode_String = 4
   TracingOutputMode_File = 8
End Enum

Public Event NewTraceLog(ByVal Msg As String, ByVal Args As Variant)

Private m_OutputMode As TracingOutputMode
Private m_TraceLogCollection As Collection
Private m_OutputString As String

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As ApplicationHandler)

On Error GoTo HandleErr

   Set m_ApplicationHandler = objRef

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandlerRef", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Property

Public Property Get ExtensionKey() As String
On Error Resume Next
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
On Error Resume Next
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)

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

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
On Error Resume Next
   Set m_ApplicationHandler = Nothing
End Sub



'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_Tracing
'---------------------------------------------------------------------------------------
'
Public Property Get TraceLogCollection() As Collection
   If m_TraceLogCollection Is Nothing Then
      Set m_TraceLogCollection = New Collection
   End If
   Set TraceLogCollection = m_TraceLogCollection
End Property

Public Property Get OutputMode() As TracingOutputMode
   OutputMode = m_OutputMode
End Property

Public Property Let OutputMode(ByVal newValue As TracingOutputMode)
   m_OutputMode = newValue
End Property

Public Sub Log(ByVal Msg As String, ParamArray Args() As Variant)
   writeTraceLog Msg, Args
End Sub

Private Sub writeTraceLog(ByRef Msg As String, ByVal Args As Variant)

   RaiseEvent NewTraceLog(Msg, Args)
   
   If (m_OutputMode And TracingOutputMode.TracingOutputMode_Debug) = TracingOutputMode.TracingOutputMode_Debug Then
      traceLog2DebugWindow Msg, Args
   End If
   
   If (m_OutputMode And TracingOutputMode.TracingOutputMode_Memory) = TracingOutputMode.TracingOutputMode_Memory Then
      traceLog2Memory Msg, Args
   End If
   
   If (m_OutputMode And TracingOutputMode.TracingOutputMode_String) = TracingOutputMode.TracingOutputMode_String Then
      traceLog2String Msg, Args, m_OutputString
   End If
   
   If (m_OutputMode And TracingOutputMode.TracingOutputMode_File) = TracingOutputMode.TracingOutputMode_File Then
      traceLog2File Msg, Args
   End If
   
End Sub

Private Sub traceLog2DebugWindow(ByRef Msg As String, ByRef Args As Variant)

   Dim tempString As String
   traceLog2String Msg, Args, tempString
   Debug.Print tempString
      
End Sub

Private Sub traceLog2Memory(ByRef Msg As String, ByRef Args As Variant)
   TraceLogCollection.Add Array(Msg, Args)
End Sub

Private Sub traceLog2String(ByRef Msg As String, ByRef Args As Variant, ByRef StringRef As String)

   Dim tempString As String
   
   If Not IsMissing(Args) Then
      If IsArray(Args) Then
         Dim arg As Variant
         For Each arg In Args
            tempString = tempString & "," & arg
         Next
         If Len(tempString) > 1 Then
            tempString = "(" & Mid(tempString, 2) & ")"
         End If
      Else
         tempString = tempString & " (" & Args & ")"
      End If
   End If
   
   
   tempString = Now() & ": " & Msg & tempString
 
   If Len(StringRef) > 0 Then
      tempString = vbNewLine & tempString
   End If
   StringRef = StringRef & tempString
   
End Sub

Private Sub traceLog2File(ByRef Msg As String, ByRef Args As Variant)
   '/**
   ' * @todo Ausgabe in Datei
   '**/
   Err.Raise vbObjectError, "Ausgabe in Datei fehlt noch!"
End Sub

'---------------------------------------------------------------------------------------
' Ereignisbehandlung von m_ApplicationHandler
'
' NewAppLog
Private Sub m_ApplicationHandler_NewAppLog(ByVal LogType As ApplicationHandlerLogType, _
                                           ByVal Msg As String, ByVal Args As Variant, _
                                           ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                           ByRef ResumeMessage As Variant)

   If LogType = ApplicationHandlerLogType.AppLogType_Tracing Then
      writeTraceLog Msg, Args
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
   End If
   
End Sub
