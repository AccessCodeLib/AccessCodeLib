Attribute VB_Name = "modApplication"
Attribute VB_Description = "Standard-Prozeduren für die Arbeit mit ApplicationHandler"
'---------------------------------------------------------------------------------------
' Module: modApplication
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
Option Compare Text
Option Explicit
Option Private Module

' Instanz der Hauptsteuerung
Private m_ApplicationHandler As ApplicationHandler

' Erweiterungen zu ApplicationHandler (Ansteuerung erfolgt über Ereignisse von ApplicationHandler)
Private m_Extension As Collection

'---------------------------------------------------------------------------------------
' Property: CurrentApplication
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
   If m_ApplicationHandler Is Nothing Then
      InitApplication
   End If
   Set CurrentApplication = m_ApplicationHandler
End Property

'---------------------------------------------------------------------------------------
' Sub: AddApplicationHandlerExtension (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung zu Collection hinzufügen
' </summary>
' <param name="objRef">Referenz auf Instanz der Erweiterung</param>
' <remarks>
' Referenz wird in Collection abgelegt, damit keine zusätzliche (manuelle)
' Referenzspeicherung notwendig ist.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddApplicationHandlerExtension(ByRef ObjRef As Object)
   Set ObjRef.ApplicationHandlerRef = CurrentApplication
   m_Extension.Add ObjRef, ObjRef.ExtensionKey
End Sub


'---------------------------------------------------------------------------------------
' Sub: TraceLog
'---------------------------------------------------------------------------------------
'/**
' <summary>
' TraceLog
' </summary>
' <param name="Param"></param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub TraceLog(ByRef Msg As String, ParamArray Args() As Variant)
   CurrentApplication.WriteLog Msg, ApplicationHandlerLogType.AppLogType_Tracing, Args, False
End Sub

Private Sub InitApplication()

   ' Hauptinstanz erzeugen
   Set m_ApplicationHandler = New ApplicationHandler
   
   ' Extension-Collection neu setzen
   Set m_Extension = New Collection
   
   'Einstellungen initialisieren
   Call InitConfig(m_ApplicationHandler)

End Sub


'---------------------------------------------------------------------------------------
' Sub: DisposeCurrentApplicationHandler
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

   Dim CheckCnt As Long, MaxCnt As Long

On Error Resume Next
   
   If Not m_ApplicationHandler Is Nothing Then
      m_ApplicationHandler.Dispose
   End If

   If Not (m_Extension Is Nothing) Then
      MaxCnt = m_Extension.Count * 2 'nur zur Sicherheit falls wider Erwarten m_Extension.Remove eine Endlosschleife bringen würde
      Do While m_Extension.Count > 0 Or CheckCnt > MaxCnt
         m_Extension.Remove 1
         CheckCnt = CheckCnt + 1
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

Public Sub WriteApplicationLogEntry(ByVal Msg As String)
   CurrentApplication.WriteApplicationLogEntry Msg
End Sub

Public Property Get PublicPath() As String
   PublicPath = CurrentApplication.PublicPath
End Property
