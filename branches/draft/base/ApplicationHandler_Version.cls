VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_Version"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_Version (Josef Pötzl, 2015-11-08)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse: Anwendungsversion prüfen
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:\n
' <code>modApplication.AddApplicationHandlerExtension New ApplicationHandler_AppFile</code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/ApplicationHandler_Version.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Enum AppVersionState
   VerState_Older = 1
   VerState_Equal = 0
   VerState_Newer = -1
End Enum

Private Const m_EXTENSION_KEY As String = "AppVersion"

Private m_XmlVersionCheckFile As String

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As ApplicationHandler)
   Set m_ApplicationHandler = objRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

Private Sub m_ApplicationHandler_ExtensionProcedureCall(ByVal sExtensionKey As String, ByVal ProcedureName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant, ProcParams() As Variant)

   If sExtensionKey <> m_EXTENSION_KEY Then
      Exit Sub
   End If
   
On Error GoTo HandleErr
   
   ResumeMode = AppResumeMode_Completed
   
   Select Case ProcedureName
      Case "NewVersionExists"
         
         ResumeMessage = NewVersionExists(ProcParams(0), ProcParams(1))
         'ProcParams = VerParams
      Case Else
         ResumeMode = AppResumeMode_Error
   End Select

   Exit Sub
   
HandleErr:
   ResumeMode = AppResumeMode_Error
   ResumeMessage = Err.Description
   
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub

Public Function NewVersionExists(Optional ByRef NewVersion As Variant, Optional ByRef VersionRemarks As Variant) As Boolean

   NewVersionExists = (CheckVersion(NewVersion, VersionRemarks) = VerState_Newer)
   
End Function


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_AppVersion
'---------------------------------------------------------------------------------------
'

Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed

   Select Case PropertyName
   
      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         
   End Select

End Sub

'---------------------------------------------------------------------------------------
' Function: CheckVersion
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Version überprüfen
' </summary>
' <param name="NewVersion"></param>
' <param name="VersionRemarks"></param>
' <returns>AppVersionState</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckVersion(Optional ByRef NewVersion As Variant, Optional ByRef VersionRemarks As Variant) As AppVersionState

   Dim DownloadVersion As String
   Dim DownloadVersionRemarks As String
   
   DownloadVersion = GetVersionFromDownloadSource(XmlVersionCheckFile, DownloadVersionRemarks)
   
   Debug.Print CurrentApplication.Version, DownloadVersion
   Debug.Print CompareVersions(CurrentApplication.Version, DownloadVersion)

   CheckVersion = CompareVersions(CurrentApplication.Version, DownloadVersion)
   If CheckVersion = VerState_Newer Then
      NewVersion = DownloadVersion
      VersionRemarks = DownloadVersionRemarks
   End If

End Function


Public Property Get XmlVersionCheckFile() As String
   XmlVersionCheckFile = m_XmlVersionCheckFile
End Property

Public Property Let XmlVersionCheckFile(ByVal NewValue As String)
   m_XmlVersionCheckFile = NewValue
End Property

Private Function GetVersionFromDownloadSource(ByVal xmlFile As String, ByRef VersionRemarks As String) As String

   Dim xmlDoc As Object ' MSXML2.DOMDocument60
   
   Set xmlDoc = CreateObject("MSXML.DOMDocument") 'New MSXML2.DOMDocument60
   xmlDoc.async = False
   If Not xmlDoc.Load(xmlFile) Then
      Err.Raise xmlDoc.parseError.errorCode, , xmlDoc.parseError.reason
   End If
   
   GetVersionFromDownloadSource = xmlDoc.getElementsByTagName("version").Item(0).Text
   VersionRemarks = xmlDoc.getElementsByTagName("remarks").Item(0).Text

End Function

Private Function GetVersionNumbers(ByVal VersionString As String) As Long()

   Dim VersionNumberStrings() As String
   Dim VersionNumbers() As Long

   VersionNumberStrings = Split(VersionString, ".")

   ReDim VersionNumbers(UBound(VersionNumberStrings))
   
   Dim i As Long
   For i = LBound(VersionNumberStrings) To UBound(VersionNumberStrings)
      VersionNumbers(i) = Val(VersionNumberStrings(i))
   Next

   GetVersionNumbers = VersionNumbers

End Function

Private Function CompareVersions(ByVal BaseVersion As String, ByVal NewVersion As String) As Long

   Dim BaseVersionNumbers() As Long
   Dim NewVersionNumbers() As Long
   
   BaseVersionNumbers = GetVersionNumbers(BaseVersion)
   NewVersionNumbers = GetVersionNumbers(NewVersion)
   
   Dim i As Long
   
   Dim MaxIndex As Long
   MaxIndex = UBound(NewVersionNumbers)
   If UBound(BaseVersionNumbers) < MaxIndex Then
      MaxIndex = UBound(BaseVersionNumbers)
   End If
   
   For i = LBound(NewVersionNumbers) To MaxIndex
      If BaseVersionNumbers(i) <> NewVersionNumbers(i) Then
         CompareVersions = -1 + 2 * Abs(BaseVersionNumbers(i) > NewVersionNumbers(i))
         Exit For
      End If
   Next

End Function
