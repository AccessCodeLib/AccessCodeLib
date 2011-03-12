VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserRolesDummy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestRelated
' UserRoles dummy (stub/mock)
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/user/UserRolesDummy.cls</file>
'  <use>user/UserRoles.cls</use>
'  <use>user/UserHandler.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Implements UserRoles

Private WithEvents m_User As UserHandler
Attribute m_User.VB_VarHelpID = -1
Private m_Roles() As String

Public Property Set User(ByVal NewUser As UserHandler)
   Set m_User = NewUser
End Property

Private Sub m_User_GetUserRoles(ByRef UserRolesRef As UserRoles)
   Set UserRolesRef = Me
End Sub

Public Function Exists(ByVal Role As String) As Boolean

   Dim tempRole As Variant
   Dim RoleExists As Boolean

   If (0 / 1) + (Not Not m_Roles) = 0 Then
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

Public Sub Add(ParamArray NewRoles() As Variant)
   
   Dim ArrayStartPos As Long

   If (0 / 1) + (Not Not m_Roles) = 0 Then
      ArrayStartPos = 0
      ReDim m_Roles(UBound(NewRoles))
   Else
      ArrayStartPos = UBound(m_Roles) + 1
      ReDim Preserve m_Roles(ArrayStartPos + UBound(NewRoles))
   End If

   Dim i As Long
   For i = 0 To UBound(NewRoles)
      m_Roles(ArrayStartPos + i) = NewRoles(i)
   Next

End Sub

Private Function UserRoles_Exists(ByVal Role As String) As Boolean
   UserRoles_Exists = Exists(Role)
End Function

Private Property Set UserRoles_User(ByVal NewUser As UserHandler)
   Set User = NewUser
End Property
