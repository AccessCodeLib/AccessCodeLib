Attribute VB_Name = "modApplication"
Attribute VB_Description = "Standard-Prozeduren für die Arbeit mit ApplicationHandler"
'---------------------------------------------------------------------------------------
' Module: modApplication (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Standard-Prozeduren für die Arbeit mit ApplicationHandler
' </summary>
' <remarks>
' </remarks>
' \ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/modApplication.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/_config_Application.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' Instanz der Hauptsteuerung
Private m_ApplicationHandler As ApplicationHandler

' Erweiterungen zu ApplicationHandler (Ansteuerung erfolgt über Ereignisse von ApplicationHandler)
Private m_Extension As Collection


'---------------------------------------------------------------------------------------
' Property: CurrentApplication (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Property für ApplicationHandler-Instanz (diese Property im Code verwenden)
' </summary>
' <returns>aktuelle Instanz von ApplicationHandler</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentApplication() As ApplicationHandler
On Error GoTo HandleErr

   If m_ApplicationHandler Is Nothing Then
      initApplication
   End If
   Set CurrentApplication = m_ApplicationHandler

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentApplication", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Property

Public Sub AddApplicationHandlerExtension(ByRef objRef As Object)

On Error GoTo HandleErr

   Set objRef.ApplicationHandlerRef = CurrentApplication
   m_Extension.Add objRef, objRef.ExtensionKey

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "AddApplicationHandlerExtension", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub initApplication()

   ' Hauptinstanz erzeugen
On Error GoTo HandleErr

   Set m_ApplicationHandler = New ApplicationHandler
   
   ' Extension-Collection neu setzen
   Set m_Extension = New Collection
   
   'Einstellungen initialisieren
   Call InitConfig(m_ApplicationHandler)

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "initApplication", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


'---------------------------------------------------------------------------------------
' Sub: DisposeCurrentApplicationHandler (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Instanz von ApplicationHandler und den Erweiterungen zerstören
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub DisposeCurrentApplicationHandler()

   Dim lngCheckCnt As Long, maxCnt As Long

On Error Resume Next
   
   If Not m_ApplicationHandler Is Nothing Then
      m_ApplicationHandler.Dispose
   End If

   If Not (m_Extension Is Nothing) Then
      maxCnt = m_Extension.Count * 2 'nur zur Sicherheit falls wider Erwarten m_Extension.Remove eine Endlosschleife bringen würde
      Do While m_Extension.Count > 0 Or lngCheckCnt > maxCnt
         m_Extension.Remove 1
         lngCheckCnt = lngCheckCnt + 1
      Loop
      Set m_Extension = Nothing
   End If
   
   Set m_ApplicationHandler = Nothing
   
End Sub


'---------------------------------------------------------------------------------------
'
' Hilfsprozeduren
'
'Public Property Get ApplicationTitlebar() As String
'   ApplicationTitlebar = CurrentApplication.Titelbar
'End Property

Public Sub WriteApplicationLogEntry(ByVal msg As String)

On Error GoTo HandleErr

   CurrentApplication.WriteApplicationLogEntry msg

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "WriteApplicationLogEntry", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Property Get PublicPath() As String

On Error GoTo HandleErr

   PublicPath = CurrentApplication.PublicPath

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "PublicPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property
