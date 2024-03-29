VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_AppLogin_Win"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_WinLogin
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erweiterung f�r ApplicationHandler-Klasse
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:
' <code>AddApplicationHandlerExtension New ApplicationHandler_AppLogin_Win</code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/login/ApplicationHandler_AppLogin_Win.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'  <use>base/login/frmLogin.frm</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const DEBUGMODE = 0

Private Const m_EXTENSION_KEY As String = "AppLogin"

Private Const m_EXTENSION_PROPERTY_UserName As String = "UserName"

'---------------------------------------------------------------------------------------
' API-Definitionen f�r Windows-Login
'---------------------------------------------------------------------------------------
'
Private Enum LogonType
       'This logon type is intended for users who will be interactively using the computer, such as a user being logged on
       'by a terminal server, remote shell, or similar process.
       'This logon type has the additional expense of caching logon information for disconnected operations;
       'therefore, it is inappropriate for some client/server applications,
       'such as a mail server.
       LOGON32_LOGON_INTERACTIVE = 2
   
       'This logon type is intended for high performance servers to authenticate plaintext passwords.
       'The LogonUser function does not cache credentials for this logon type.
       LOGON32_LOGON_NETWORK = 3
   
       'This logon type is intended for batch servers, where processes may be executing on behalf of a user without
       'their direct intervention. This type is also for higher performance servers that process many plaintext
       'authentication attempts at a time, such as mail or Web servers.
       'The LogonUser function does not cache credentials for this logon type.
       LOGON32_LOGON_BATCH = 4
   
       'Indicates a service-type logon. The account provided must have the service privilege enabled.
       LOGON32_LOGON_SERVICE = 5
   
       'This logon type is for GINA DLLs that log on users who will be interactively using the computer.
       'This logon type can generate a unique audit record that shows when the workstation was unlocked.
       LOGON32_LOGON_UNLOCK = 7
   
       'This logon type preserves the name and password in the authentication package, which allows the server to make
       'connections to other network servers while impersonating the client. A server can accept plaintext credentials
       'from a client, call LogonUser, verify that the user can access the system across the network, and still
       'communicate with other servers.
       'NOTE: Windows NT:  This value is not supported.
       LOGON32_LOGON_NETWORK_CLEARTEXT = 8
   
       'This logon type allows the caller to clone its current token and specify new credentials for outbound connections.
       'The new logon session has the same local identifier but uses different credentials for other network connections.
       'NOTE: This logon type is supported only by the LOGON32_PROVIDER_WINNT50 logon provider.
       'NOTE: Windows NT:  This value is not supported.
       LOGON32_LOGON_NEW_CREDENTIALS = 9
   End Enum
   
   Private Enum LogonProvider
       'Use the standard logon provider for the system.
       'The default security provider is negotiate, unless you pass NULL for the domain name and the user name
       'is not in UPN format. In this case, the default provider is NTLM.
       'NOTE: Windows 2000/NT:   The default security provider is NTLM.
       LOGON32_PROVIDER_DEFAULT = 0
   End Enum

   Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" ( _
               ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, _
               ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
               
   Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean

'Private Declare Function RevertToSelf Lib "advapi32.dll" () As Long

'Private Const LOGON32_PROVIDER_DEFAULT          As Long = 0&
'Private Const LOGON32_PROVIDER_WINNT35          As Long = 1&
'Private Const LOGON32_LOGON_INTERACTIVE         As Long = 2&
'Private Const LOGON32_LOGON_NETWORK             As Long = 3&
'Private Const LOGON32_LOGON_NETWORK_CLEARTEXT   As Long = 8&
'Private Const LOGON32_LOGON_BATCH               As Long = 4&
'Private Const LOGON32_LOGON_SERVICE             As Long = 5&



'---------------------------------------------------------------------------------------
' Hilfsvariablen
'---------------------------------------------------------------------------------------

   Private m_UserDomain As String
   Private m_UserName As String
   Private m_LoginNameWithDomain As Boolean
   
   Private WithEvents m_LoginForm As Form_frmLogin
Attribute m_LoginForm.VB_VarHelpID = -1
   Private m_LoginCancelled As Boolean
   
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
' Erg�nzungen f�r Ereiterung: ApplicationHandler_WinLogin
'---------------------------------------------------------------------------------------
'
Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)
   Select Case PropertyName
      Case m_EXTENSION_PROPERTY_UserName
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Completed
         ResumeMessage = m_UserName
     
      Case Else 'Property wurde nicht erkannt
         ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Error
       
   End Select
End Sub

Public Property Set LoginForm(ByRef formRef As Form_frmLogin)
   Set m_LoginForm = formRef
End Property

' Login-Pr�fung vor Anwenduntsstart einschieben
Private Sub m_ApplicationHandler_BeforeStartApplication(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)

   Dim bolLoginFailed As Boolean

   bolLoginFailed = Not Login()
   If bolLoginFailed Then
      ResumeMode = ApplicationHandlerResumeModes.AppResumeMode_Cancel
#If DEBUGMODE Then
      MsgBox "Application.Quit"
#Else
      Application.Quit
#End If
   End If
 
End Sub

'---------------------------------------------------------------------------------------
' Function: Login (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Aufruf des Login-Vorgangs
' </summary>
' <returns>True = Login war erfolgreich</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Login() As Boolean
   
   Dim strDefaultUser As String
   Dim openArgString As String
   
   If Not (m_ApplicationHandler Is Nothing) Then
      openArgString = m_ApplicationHandler.ApplicationName & " - "
   End If
   
   If Len(m_UserName) > 0 Then
      strDefaultUser = m_UserName
   Else
      strDefaultUser = Environ("username")
   End If
   
   openArgString = openArgString & "Login@" & strDefaultUser

   DoCmd.OpenForm "frmLogin", acNormal, , , acFormEdit, acDialog, openArgString
   '=> Im LoginFormular erfolgt Aufruf der Eigenschaft "LoginForm",
   '   damit wird Ereignisbehandlung aktiviert => m_LoginCancelled wird eingestellt
   Login = Not m_LoginCancelled

End Function

Private Function CheckUser(ByVal Domain As String, ByVal userName As String, ByVal userPassword As String) As Boolean

   Dim tokenHandle As Long
   Dim logonAllowed As Boolean

   'mehrere LogonTypes ausprobieren
   logonAllowed = (LogonUser(userName, Domain, userPassword, LogonType.LOGON32_LOGON_NETWORK, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
   If logonAllowed = False Then
       logonAllowed = (LogonUser(userName, Domain, userPassword, LogonType.LOGON32_LOGON_INTERACTIVE, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
       If logonAllowed = False Then
           logonAllowed = (LogonUser(userName, Domain, userPassword, LogonType.LOGON32_LOGON_NETWORK_CLEARTEXT, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
       End If
   End If

   If logonAllowed Then
       If tokenHandle <> 0 Then
           CloseHandle tokenHandle
       Else
           logonAllowed = False
       End If
   End If
   
   If Not logonAllowed Then
      MsgBox "Anmeldung ist fehlgeschlagen", vbInformation
   End If
   
   CheckUser = logonAllowed

End Function

'Ereignisbehandlung vom Login-Formular
Private Sub m_LoginForm_cancelled()
   m_UserName = vbNullString
   m_LoginCancelled = True
End Sub

Private Sub m_LoginForm_Logon(ByVal LoginName As String, ByVal Password As String, Cancel As Boolean)
   
   Dim strDomain As String
   Dim pos As Long
   Dim bolLoginOk As Boolean
   
   m_LoginCancelled = False
   
   If LoginNameWithDomain Then 'Domain aus Username - Format: domain\username
      pos = InStr(1, LoginName, "\", vbTextCompare)
      If pos > 0 Then
         strDomain = Left$(LoginName, pos - 1)
         LoginName = Mid$(LoginName, pos + 1)
      End If
   Else
      strDomain = Domain
   End If
   
   bolLoginOk = CheckUser(strDomain, LoginName, Password)
   If bolLoginOk Then
      m_UserName = LoginName
   Else
      Cancel = True
   End If

End Sub

'---------------------------------------------------------------------------------------
' Property: Domain (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Domain
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Domain(ByVal DomainName As String)
   m_UserDomain = DomainName
End Property

'---------------------------------------------------------------------------------------
' Property: Domain (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Domain
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Domain() As String
   Domain = m_UserDomain
End Property

'---------------------------------------------------------------------------------------
' Property: LoginNameWithDomain (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Loginkennung enth�lt Domain
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let LoginNameWithDomain(ByVal WithDomain As Boolean)
   m_LoginNameWithDomain = WithDomain
End Property

'---------------------------------------------------------------------------------------
' Property: LoginNameWithDomain (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Loginkennung enth�lt Domain
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get LoginNameWithDomain() As Boolean
   LoginNameWithDomain = m_LoginNameWithDomain
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentUser (Josef P�tzl, 2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Aktuell angemeldeter User
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentUser()
   CurrentUser = m_UserName
End Property
