VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_Settings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_Settings
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse: Einstellungen
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_DbConnection</code>
' oder mittels ExtensionCollection:\n
' <code>ExtensionsCollectionReferenz.Add New ApplicationHandler_DbConnection</code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_Settings.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "Settings"
Private m_Disposed As Boolean

Private m_Settings As Scripting.Dictionary

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef ObjRef As ApplicationHandler)
   Set m_ApplicationHandler = ObjRef
   m_Disposed = False
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

Private Sub m_ApplicationHandler_AfterDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   Set m_ApplicationHandler = Nothing
End Sub

Private Sub m_ApplicationHandler_BeforeDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef Exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then Exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

Private Sub m_ApplicationHandler_ExtensionPropertyLookup(ByVal Key As String, ByVal PropertyName As String, ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   
   If Key <> m_EXTENSION_KEY Then
      Exit Sub
   End If
   
   With Settings
      If .Exists(PropertyName) Then
         ResumeMode = AppResumeMode_Completed
         ResumeMessage = .Item(PropertyName)
      Else
         ResumeMode = AppResumeMode_Error
         ResumeMessage = PropertyName & " ist nicht vorhanden"
      End If
   End With
   
   
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Public Sub Dispose()

On Error Resume Next

   If m_Disposed Then Exit Sub
   
   Set m_Settings = Nothing
   
   m_Disposed = (0 = Err.Number)

End Sub


'---------------------------------------------------------------------------------------
' Property: Settings
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Referenz auf Scripting.Dictionary für die Einstellungen
' </summary>
' <returns>Scripting.Dictionary</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Settings() As Scripting.Dictionary
   If m_Settings Is Nothing Then
      Set m_Settings = New Scripting.Dictionary
      m_Disposed = False
   End If
   Set Settings = m_Settings
End Property

'---------------------------------------------------------------------------------------
' Property: Settings
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Einstellung anfügen
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Add(ByVal Key As String, ByVal Value As Variant)
   Settings.Add Key, Value
End Sub
