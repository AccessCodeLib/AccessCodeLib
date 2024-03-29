VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WinLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: WinLogin
'---------------------------------------------------------------------------------------
'/**
' @brief Login-Pr�fung �ber Windows-Accounts
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>user/login/WinLogin.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' API-Definitionen f�r Windows-Login
'---------------------------------------------------------------------------------------
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

#If VBA7 Then
   Private Declare PtrSafe Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" ( _
            ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, _
            ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
            
   Private Declare PtrSafe Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
#Else
   Private Declare Function LogonUser Lib "advapi32.dll" Alias "LogonUserA" ( _
               ByVal lpszUsername As String, ByVal lpszDomain As String, ByVal lpszPassword As String, _
               ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
               
   Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Boolean
#End If

Private m_UserDomain As String
Private m_UserName As String
Private m_LoginNameWithDomain As Boolean

Public Event LoginSuccessful(ByVal UserName As String)
   
Public Function Login(ByVal LoginName As String, ByVal Password As String) As Boolean
   
   Dim DomainName As String
   Dim Pos As Long
   Dim LoginOk As Boolean
   
   If LoginNameWithDomain Then 'Domain aus Username - Format: domain\username
      Pos = InStr(1, LoginName, "\", vbTextCompare)
      If Pos > 0 Then
         DomainName = Left$(LoginName, Pos - 1)
         LoginName = Mid$(LoginName, Pos + 1)
      End If
   Else
      DomainName = Domain
   End If
   
   LoginOk = CheckUser(DomainName, LoginName, Password)
   If LoginOk Then
      m_UserName = LoginName
      RaiseEvent LoginSuccessful(m_UserName)
   End If

   Login = LoginOk

End Function

Private Function CheckUser(ByVal DomainName As String, ByVal UserName As String, ByVal UserPassword As String) As Boolean

   Dim tokenHandle As Long
   Dim logonAllowed As Boolean

   'mehrere LogonTypes ausprobieren
   logonAllowed = (LogonUser(UserName, DomainName, UserPassword, LogonType.LOGON32_LOGON_NETWORK, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
   If logonAllowed = False Then
       logonAllowed = (LogonUser(UserName, DomainName, UserPassword, LogonType.LOGON32_LOGON_INTERACTIVE, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
       If logonAllowed = False Then
           logonAllowed = (LogonUser(UserName, DomainName, UserPassword, LogonType.LOGON32_LOGON_NETWORK_CLEARTEXT, LogonProvider.LOGON32_PROVIDER_DEFAULT, tokenHandle) > 0)
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

Public Property Let Domain(ByVal DomainName As String)
   m_UserDomain = DomainName
End Property

Public Property Get Domain() As String
   Domain = m_UserDomain
End Property

Public Property Let LoginNameWithDomain(ByVal WithDomain As Boolean)
   m_LoginNameWithDomain = WithDomain
End Property

Public Property Get LoginNameWithDomain() As Boolean
   LoginNameWithDomain = m_LoginNameWithDomain
End Property

Public Property Get CurrentUser()
   CurrentUser = m_UserName
End Property
