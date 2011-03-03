VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: UserRoles
'---------------------------------------------------------------------------------------
'/**
' @brief Diese Klasse ist für die Rechteverwaltung/Rollenzuordung zuständig
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>user/UserRoles.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>user/UserHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_User As UserHandler
Attribute m_User.VB_VarHelpID = -1
Private m_Roles() As String

Public Property Set User(ByVal NewUser As UserHandler)
   Set m_User = NewUser
   ReadUserRoles
End Property

Private Sub m_User_GetUserRoles(ByRef UserRolesRef As UserRoles)
   Set UserRolesRef = Me
End Sub

Private Sub m_User_UserChanged(ByVal NewUserName As String)
   ReadUserRoles
End Sub

Private Sub ReadUserRoles()

   Dim UserName As String

   UserName = m_User.UserName
   If Len(UserName) = 0 Then
      Erase m_Roles
      Exit Sub
   End If

   '...

End Sub

Public Function Exists(ByVal Role As String) As Boolean

   Dim tempRole As Variant
   Dim RoleExists As Boolean

   If (0 / 1) + (Not Not m_Roles) = 0 Then 'Array nicht initialisiert
      Exists = False
      Exit Function
   End If

   For Each tempRole In m_Roles
      If StrComp(tempRole, Role, vbTextCompare) = 0 Then
         RoleExists = True
         Exit For
      End If
   Next

   Exists = RoleExists

End Function
