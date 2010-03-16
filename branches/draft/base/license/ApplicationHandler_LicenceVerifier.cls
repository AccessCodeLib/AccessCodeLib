VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_LicenceVerifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_WinLogin (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung für ApplicationHandler-Klasse
' </summary>
' <remarks>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/licence/ApplicationHandler_LicenceVerifier.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>crypt/md5/modMD5.bas</use>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const DEBUGMODE = 0

Private Const m_EXTENSION_KEY As String = "LicenceVerifier"
Private Const m_EXTENSION_PROPERTY_LicenceName As String = "LicenceName"
Private Const m_EXTENSION_PROPERTY_LicenceKey As String = "LicenceKey"

Private Const m_VERIFYINFO_KEY_invalid As String = "Der eingebene Lizenschlüssel ist ungültig."

'---------------------------------------------------------------------------------------
' Ereignisse
'---------------------------------------------------------------------------------------

Public Event LicenceInfoLookup(ByRef sLicenceName As String, ByRef sLicenceKey As String, _
                               ByVal sLicenceNameDefaultValue As String, ByVal sLicenceKeyDefaultValue As String)

'---------------------------------------------------------------------------------------
' Hilfsvariablen
'---------------------------------------------------------------------------------------

   Private m_ApplicationKey As String
   Private m_LicenceKey_KeyLen As Long
   Private m_LicenceKey_Prefix As String
   Private m_LicenceKey_Suffix As String
   Private m_LicenceKey_Loops As Long
  
   Private m_LicenceName As String
   Private m_LicenceKey As String
   
   Private m_LicenceNameInputboxDefault As String
   Private m_LicenceKeyInputboxDefault As String
   
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
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, exists As Boolean)
On Error Resume Next
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
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
On Error Resume Next
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_AppLicenceVerifier
'---------------------------------------------------------------------------------------
'
Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
   
On Error GoTo HandleErr

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName
      Case m_EXTENSION_PROPERTY_LicenceName
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
         ResumeMessage = True
     
      Case Else 'Property wurde nicht erkannt
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
       
   End Select

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExtWinLogin.getExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub m_ApplicationHandler_BeforeStartApplication(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)

   Dim bolVerifyFailed As Boolean
   Dim strVerifyMessage As String
   
   bolVerifyFailed = Not VerifyLicence(strVerifyMessage)
   If bolVerifyFailed Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Cancel
#If DEBUGMODE Then
      MsgBox "Application.Quit"
#Else
      Application.Quit
#End If
   End If

End Sub

Public Sub Init(ByVal sApplicationKey As String, _
                ByVal lLicenceKey_KeyLen As Long, _
       Optional ByVal sLicenceKey_Prefix As String = vbNullString, _
       Optional ByVal sLicenceKey_Suffix As String = vbNullString, _
       Optional ByVal lLicenceKey_Loops As Long = 0)

   m_ApplicationKey = sApplicationKey
   m_LicenceKey_KeyLen = lLicenceKey_KeyLen
   m_LicenceKey_Prefix = sLicenceKey_Prefix
   m_LicenceKey_Suffix = sLicenceKey_Suffix
   m_LicenceKey_Loops = lLicenceKey_Loops
   
End Sub

Public Property Get LicenceName() As String
   
   If StrPtr(m_LicenceName) = 0 Then
      readLicenceInfo
   End If
   LicenceName = m_LicenceName
   
End Property

Public Property Get LicenceKey() As String
   
   If StrPtr(m_LicenceKey) = 0 Then
      readLicenceInfo
   End If
   LicenceKey = m_LicenceKey
   
End Property

Private Sub readLicenceInfo()

On Error GoTo HandleErr

   RaiseEvent LicenceInfoLookup(m_LicenceName, m_LicenceKey, m_LicenceNameInputboxDefault, m_LicenceKeyInputboxDefault)
   
   'Name
   If StrPtr(m_LicenceName) = 0 Then
      m_LicenceName = getLicenceInfoFromAppDb(m_EXTENSION_PROPERTY_LicenceName)
'     If StrPtr(m_LicenceName) = 0 Then
'/**
' * @todo Lizendaten mit Formular abfragen
'**/
'     End If
           
      If StrPtr(m_LicenceName) = 0 Then 'Notfalllösung über Inputbox
         m_LicenceName = InputBox("Bitte Namen eingeben:", "Lizenzdaten eingeben", m_LicenceNameInputboxDefault)
         If Not StrPtr(m_LicenceName) = 0 Then 'speichern
            m_LicenceNameInputboxDefault = m_LicenceName
            m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenceName, m_LicenceName, dbText
         End If
      End If
   End If
   
   'Key
   If StrPtr(m_LicenceKey) = 0 Then
      m_LicenceKey = getLicenceInfoFromAppDb(m_EXTENSION_PROPERTY_LicenceKey)
      
      If StrPtr(m_LicenceKey) = 0 Then 'Notfalllösung über Inputbox
         m_LicenceKey = InputBox("Bitte Lizenzschlüssel eingeben:", "Lizenzdaten eingeben", m_LicenceKeyInputboxDefault)
         If Not StrPtr(m_LicenceKey) = 0 Then 'speichern
            m_LicenceKeyInputboxDefault = m_LicenceKey
            m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenceKey, m_LicenceKey, dbText
         End If
      End If
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "readLicenceInfo", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function getLicenceInfoFromAppDb(propName As String) As String
   
   Dim props As DAO.Properties
   Dim prop As DAO.Property

   Set props = m_ApplicationHandler.AppDb.Properties
   For Each prop In props
      If prop.Name = propName Then
         getLicenceInfoFromAppDb = prop.Value
         Exit For
      End If
   Next
   
End Function

Friend Function GenerateKey(ByVal sLicenceName As String, _
                            ByVal sApplicationKey As String, _
                            ByVal lLicenceKey_KeyLen As Long, _
                   Optional ByVal sLicenceKey_Prefix As String = vbNullString, _
                   Optional ByVal sLicenceKey_Suffix As String = vbNullString, _
                   Optional ByVal lLicenceKey_Loops As Long = 0) As String

   Dim strAppKeyCode As String
   Dim strLicenceNameCode As String
   Dim strLicenceKeyCode As String
   Dim i As Long
   
   strLicenceNameCode = md5(sLicenceKey_Prefix & UCase$(sLicenceName) & sLicenceKey_Suffix)
   strAppKeyCode = md5(sApplicationKey & sApplicationKey)
   
   For i = 1 To lLicenceKey_Loops
      strLicenceNameCode = md5(strLicenceNameCode)
      strAppKeyCode = md5(strAppKeyCode)
   Next

   strLicenceKeyCode = Left$(md5(strAppKeyCode & strLicenceNameCode & strAppKeyCode), lLicenceKey_KeyLen)
   Do While Len(strLicenceKeyCode) < lLicenceKey_KeyLen
      strLicenceKeyCode = Left$(strLicenceKeyCode & md5(strLicenceKeyCode), lLicenceKey_KeyLen)
   Loop
   
   GenerateKey = strLicenceKeyCode

End Function

Friend Function VerifyLicence(Optional ByRef sMsg As String) As Boolean
'Beispiel für LicenceKey: E71C2-7C3EE-34C7E-07C04

  
   Dim checkKey(1) As String

   checkKey(0) = Trim$(Replace$(Me.LicenceKey, "-", vbNullString))
   If Len(checkKey(0)) <> m_LicenceKey_KeyLen Then
      sMsg = m_VERIFYINFO_KEY_invalid
      VerifyLicence = False
      Exit Function
   End If
   
   checkKey(1) = GenerateKey(Me.LicenceName, m_ApplicationKey, m_LicenceKey_KeyLen, _
                             m_LicenceKey_Prefix, m_LicenceKey_Suffix, m_LicenceKey_Loops)

   If checkKey(0) = checkKey(1) Then
      VerifyLicence = True
   Else
      sMsg = m_VERIFYINFO_KEY_invalid
      VerifyLicence = False
   End If
   

End Function
