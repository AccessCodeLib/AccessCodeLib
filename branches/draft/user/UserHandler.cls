VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: UserHandler
'---------------------------------------------------------------------------------------
'/**
' @brief Benutzerverwaltung/Benutzersteuerung
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>user/UserHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>user/UserRoles.cls</use>
'  <test>_test/user/UserHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_UserName As String

Public Event UserBeforeChange(ByRef NewUserName As String, ByRef Cancel As Boolean)
Public Event UserChanged(ByVal NewUserName As String)

'Event zum Vermeiden von Zirkelbezug:
Public Event GetUserRoles(ByRef UserRolesRef As UserRoles)

Public Property Get UserName() As String
Attribute UserName.VB_UserMemId = 0
'Attribute Item.VB_UserMemId = 0 ' =>  Standardelement
   If StrPtr(m_UserName) = 0 Then
      SetUserName DefaultWindowsUser
   End If
   UserName = m_UserName
End Property

Public Property Let UserName(ByVal NewUserName As String)
   SetUserName NewUserName
End Property

Private Sub SetUserName(ByVal NewUserName As String)

   Dim Cancel As Boolean

   RaiseEvent UserBeforeChange(NewUserName, Cancel)
   If Cancel Then
      Exit Sub
   End If

   m_UserName = NewUserName
   RaiseEvent UserChanged(m_UserName)

End Sub

Public Property Get DefaultWindowsUser() As String
   DefaultWindowsUser = Environ("username") ' => Bei Bedarf auf API-Variante umstellen
End Property

Public Property Get Roles() As UserRoles

   Dim TempUserRoles As UserRoles

   RaiseEvent GetUserRoles(TempUserRoles)
   If TempUserRoles Is Nothing Then '... Notbehelf falls auf Ereignis nicht reagiert wurde
      Set TempUserRoles = New UserRoles
      Set TempUserRoles.User = Me
   End If

   Set Roles = TempUserRoles

End Property

Public Function IsRoleMember(ByVal Role As String) As Boolean
   IsRoleMember = Roles.Exists(Role)
End Function
