VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_JetUserHook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: ApplicationHandler_JetUserHook
'---------------------------------------------------------------------------------------
'/**
' @autor    Josef P�tzl
' <summary>
' Erweiterung f�r ApplicationHandler-Klasse
' </summary>
' <remarks>
' Aktivierung im _config_Application-Modul:
' <code>
'   Dim obJetUserHook As ApplicationHandler_JetUserHook
'   Set obJetUserHook = New ApplicationHandler_JetUserHook
'   AddApplicationHandlerExtension obJetUserHook
'   obJetUserHook.IdentifyBackendTableDefName = "Name einer verkn�pften Tabelle"
'   Set obJetUserHook = Nothing
' </code>
' </remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/jet/ApplicationHandler_JetUserHook.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const m_EXTENSION_KEY As String = "JetUserHook"

'---------------------------------------------------------------------------------------
' Hilfsvariablen
'---------------------------------------------------------------------------------------

   Private m_UserWorkspace As DAO.Workspace
   Private m_UserDatabase As DAO.Database
   
   Private m_IdentifyBackendTableDefName As String
   
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


' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
On Error Resume Next
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Erg�nzungen f�r Ereiterung: ApplicationHandler_JetUserHook
'---------------------------------------------------------------------------------------
'

' Benutzerspezifische Verbindung nach Anwendungsstart �ffnen
Private Sub m_ApplicationHandler_AfterStartApplication(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   
   openUserConnection

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

Private Sub openUserConnection()
   
   Dim strUserName As String
   Dim strTdfName As String
   Dim strDbName As String
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   strUserName = m_ApplicationHandler.GetExtensionProperty("AppLogin", "UserName", vbNullString)
   If Len(strUserName) = 0 Then
      Exit Sub 'oder einen Fehler ausl�sen?
   End If
   
   strTdfName = Me.IdentifyBackendTableDefName
   If Len(strTdfName) = 0 Then
      Exit Sub
   End If
   
   Set db = CurrentDb
   Set tdf = db.TableDefs(strTdfName)
   If Len(tdf.Connect) = 0 Then
      Exit Sub
   End If
   
   strDbName = Trim(Mid(tdf.Connect, InStr(1, tdf.Connect, "DATABASE=") + 9))
   If Len(strDbName) = 0 Then
      Exit Sub
   End If
   
   
   checkUserInSystemDb strUserName
   
   Set m_UserWorkspace = DBEngine.CreateWorkspace(m_EXTENSION_KEY, strUserName, vbNullString)
   Set m_UserDatabase = m_UserWorkspace.OpenDatabase(strDbName)
   ' ... diese Instanzen k�nnte man auch f�r Datenoperationen nutzen! => Property zum Auslesen anf�gen
   
End Sub


Private Sub checkUserInSystemDb(ByVal sUserName As String)

   Dim ws As DAO.Workspace
   Dim wsUsers As DAO.Users
   Dim usr As DAO.User
   Dim bolUserExists As Boolean
   
   Set ws = DBEngine.Workspaces(0)
   Set wsUsers = ws.Users
   
   For Each usr In wsUsers
      If usr.Name = sUserName Then
         bolUserExists = True
         Exit For
      End If
   Next
   
   If Not bolUserExists Then
      wsUsers.Append ws.CreateUser(sUserName, sUserName)
   End If

End Sub


Private Sub Class_Terminate()
On Error Resume Next

   If Not m_UserDatabase Is Nothing Then
      m_UserDatabase.Close
      Set m_UserDatabase = Nothing
   End If

   If Not m_UserWorkspace Is Nothing Then
      m_UserWorkspace.Close
      Set m_UserWorkspace = Nothing
   End If
End Sub


Public Property Get IdentifyBackendTableDefName() As String
   IdentifyBackendTableDefName = m_IdentifyBackendTableDefName
End Property
Public Property Let IdentifyBackendTableDefName(ByVal sTdfName As String)
   m_IdentifyBackendTableDefName = sTdfName
End Property