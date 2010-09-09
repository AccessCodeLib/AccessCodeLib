VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_LicenseVerifier"
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
'  <file>base/license/ApplicationHandler_LicenseVerifier.cls</file>
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

Private Const m_EXTENSION_KEY As String = "LicenseVerifier"
Private Const m_EXTENSION_PROPERTY_LicenseName As String = "LicenseName"
Private Const m_EXTENSION_PROPERTY_LicenseKey As String = "LicenseKey"

Private Const m_VERIFYINFO_KEY_invalid As String = "Der eingebene Lizenschlüssel ist ungültig."

Private m_UseInputBox As Boolean

'---------------------------------------------------------------------------------------
' Ereignisse
'---------------------------------------------------------------------------------------

Public Event LicenseInfoLookup(ByRef sLicenseName As String, ByRef sLicenseKey As String, _
                               ByVal sLicenseNameDefaultValue As String, ByVal sLicenseKeyDefaultValue As String)

Public Event LicenseChecked(ByVal Verfified As Boolean)

'---------------------------------------------------------------------------------------
' Hilfsvariablen
'---------------------------------------------------------------------------------------

   Private m_ApplicationKey As String
   Private m_LicenseKey_KeyLen As Long
   Private m_LicenseKey_Prefix As String
   Private m_LicenseKey_Suffix As String
   Private m_LicenseKey_Loops As Long
  
   Private m_LicenseName As String
   Private m_LicenseKey As String
   
   Private m_LicenseNameInputboxDefault As String
   Private m_LicenseKeyInputboxDefault As String
   
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
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
                     
   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ApplicationHandler_AppLicenseVerifier
'---------------------------------------------------------------------------------------
'
Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName
      Case m_EXTENSION_PROPERTY_LicenseName
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
         ResumeMessage = True
     
      Case Else 'Property wurde nicht erkannt
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
       
   End Select

End Sub

Private Sub m_ApplicationHandler_BeforeStartApplication(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)

   Dim bolVerifyFailed As Boolean
   Dim strVerifyMessage As String
   
   bolVerifyFailed = Not VerifyLicense(strVerifyMessage)
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
                ByVal lLicenseKey_KeyLen As Long, _
       Optional ByVal sLicenseKey_Prefix As String = vbNullString, _
       Optional ByVal sLicenseKey_Suffix As String = vbNullString, _
       Optional ByVal lLicenseKey_Loops As Long = 0, _
       Optional ByVal bUseInputBoxForLicenseData As Boolean = False)

   m_ApplicationKey = sApplicationKey
   m_LicenseKey_KeyLen = lLicenseKey_KeyLen
   m_LicenseKey_Prefix = sLicenseKey_Prefix
   m_LicenseKey_Suffix = sLicenseKey_Suffix
   m_LicenseKey_Loops = lLicenseKey_Loops
   m_UseInputBox = bUseInputBoxForLicenseData

End Sub

Public Property Get LicenseName() As String
   If StrPtr(m_LicenseName) = 0 Then
      readLicenseInfo
   End If
   LicenseName = m_LicenseName
End Property

Public Property Let LicenseName(newValue As String)
   m_LicenseName = newValue
End Property

Public Property Get LicenseKey() As String
   If StrPtr(m_LicenseKey) = 0 Then
      readLicenseInfo
   End If
   LicenseKey = m_LicenseKey
End Property

Public Property Let LicenseKey(newValue As String)
   m_LicenseKey = newValue
End Property

Private Sub readLicenseInfo()

   RaiseEvent LicenseInfoLookup(m_LicenseName, m_LicenseKey, m_LicenseNameInputboxDefault, m_LicenseKeyInputboxDefault)
   
   'Name
   If StrPtr(m_LicenseName) = 0 Then
      m_LicenseName = getLicenseInfoFromAppDb(m_EXTENSION_PROPERTY_LicenseName)
'     If StrPtr(m_LicenseName) = 0 Then
'/**
' * @todo Lizendaten mit Formular abfragen
'**/
'     End If
           
      If StrPtr(m_LicenseName) = 0 And m_UseInputBox Then 'Notfalllösung über Inputbox
         m_LicenseName = InputBox("Bitte Namen eingeben:", "Lizenzdaten eingeben", m_LicenseNameInputboxDefault)
         If Not StrPtr(m_LicenseName) = 0 Then 'speichern
            m_LicenseNameInputboxDefault = m_LicenseName
            m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseName, m_LicenseName, dbText
         End If
      End If
   End If
   
   'Key
   If StrPtr(m_LicenseKey) = 0 Then
      m_LicenseKey = getLicenseInfoFromAppDb(m_EXTENSION_PROPERTY_LicenseKey)
      
      If StrPtr(m_LicenseKey) = 0 And m_UseInputBox Then 'Notfalllösung über Inputbox
         m_LicenseKey = InputBox("Bitte Lizenzschlüssel eingeben:", "Lizenzdaten eingeben", m_LicenseKeyInputboxDefault)
         If Not StrPtr(m_LicenseKey) = 0 Then 'speichern
            m_LicenseKeyInputboxDefault = m_LicenseKey
            m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseKey, m_LicenseKey, dbText
         End If
      End If
   End If

End Sub

Private Function getLicenseInfoFromAppDb(propName As String) As String
   
   Dim props As DAO.Properties
   Dim prop As DAO.Property

   Set props = m_ApplicationHandler.AppDb.Properties
   For Each prop In props
      If prop.Name = propName Then
         getLicenseInfoFromAppDb = prop.Value
         Exit For
      End If
   Next
   
End Function

Friend Function GenerateKey(ByVal sLicenseName As String, _
                            ByVal sApplicationKey As String, _
                            ByVal lLicenseKey_KeyLen As Long, _
                   Optional ByVal sLicenseKey_Prefix As String = vbNullString, _
                   Optional ByVal sLicenseKey_Suffix As String = vbNullString, _
                   Optional ByVal lLicenseKey_Loops As Long = 0) As String

   Dim strAppKeyCode As String
   Dim strLicenseNameCode As String
   Dim strLicenseKeyCode As String
   Dim i As Long
   
   strLicenseNameCode = md5(sLicenseKey_Prefix & UCase$(sLicenseName) & sLicenseKey_Suffix)
   strAppKeyCode = md5(sApplicationKey & sApplicationKey)
   
   For i = 1 To lLicenseKey_Loops
      strLicenseNameCode = md5(strLicenseNameCode)
      strAppKeyCode = md5(strAppKeyCode)
   Next

   strLicenseKeyCode = Left$(md5(strAppKeyCode & strLicenseNameCode & strAppKeyCode), lLicenseKey_KeyLen)
   Do While Len(strLicenseKeyCode) < lLicenseKey_KeyLen
      strLicenseKeyCode = Left$(strLicenseKeyCode & md5(strLicenseKeyCode), lLicenseKey_KeyLen)
   Loop
   
   GenerateKey = strLicenseKeyCode

End Function

Friend Function VerifyLicense(Optional ByRef sMsg As String) As Boolean
'Beispiel für LicenseKey: E71C2-7C3EE-34C7E-07C04

  
   Dim checkKey(1) As String
   Dim bolVerified As Boolean

   checkKey(0) = Trim$(Replace$(Me.LicenseKey, "-", vbNullString))
   If Len(checkKey(0)) <> m_LicenseKey_KeyLen Then
      sMsg = m_VERIFYINFO_KEY_invalid
      VerifyLicense = False
   Else
      checkKey(1) = GenerateKey(Me.LicenseName, m_ApplicationKey, m_LicenseKey_KeyLen, _
                                m_LicenseKey_Prefix, m_LicenseKey_Suffix, m_LicenseKey_Loops)
   
      If checkKey(0) = checkKey(1) Then
         bolVerified = True
      Else
         sMsg = m_VERIFYINFO_KEY_invalid
         bolVerified = False
      End If
   
   End If
   
   RaiseEvent LicenseChecked(bolVerified)

   VerifyLicense = bolVerified

End Function

Public Sub RemoveLicenseKey()

   m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseName, Null, dbText
   m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseKey, Null, dbText
   
   m_LicenseName = vbNullString
   m_LicenseKey = vbNullString

End Sub

Friend Sub Save()

   m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseName, m_LicenseName, dbText
   m_ApplicationHandler.SetApplicationProperty m_EXTENSION_PROPERTY_LicenseKey, m_LicenseKey, dbText

End Sub
