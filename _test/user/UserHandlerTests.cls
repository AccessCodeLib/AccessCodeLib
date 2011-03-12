VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UserHandlerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' AccUnit:TestClass
'--------------------------------------------------------------------
'/**
' @brief Tests für Benutzerverwaltung/Benutzersteuerung
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/user/UserHandlerTests.cls</file>
'  <use>_test/user/UserRolesDummy.cls</use>
'  <use>user/UserHandler.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'VBA-Inspector:Ignore

Private Const UserName As String = "TestUser"
Private WithEvents m_UserHandler As UserHandler
Attribute m_UserHandler.VB_VarHelpID = -1
Private m_UserRolesDummy As UserRolesDummy
Private m_NewUserName As String

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

Public Sub Setup()
   InitUserHandler
   m_NewUserName = vbNullString
End Sub

Public Sub Teardown()
   DisposeUserHandler
   m_NewUserName = vbNullString
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub UserName_UseDefaultMember_ReturnsString()
   
   Dim Actual As String

   Actual = m_UserHandler
   Assert.IsNotEmpty Actual

End Sub

Public Sub UserName_ChangeUserNameWithProperty_UseAssignedName()

   Const Expected As String = "abc"
   Dim Actual As String

   m_UserHandler.UserName = "abc"
   
   Actual = m_UserHandler.UserName
   Assert.That Actual, Iz.EqualTo(Expected)

   Actual = m_NewUserName
   Assert.That Actual, Iz.EqualTo(Expected), "event handler 'UserChanged' failed"

End Sub

Public Sub UserName_ReChangeUserNameInEvent_UseNameAssignedByEventProc()

   Const Expected As String = "xyz"
   m_NewUserName = "xyz"

   Dim Actual As String
   m_UserHandler.UserName = "abc"
   
   Actual = m_UserHandler.UserName
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Private Sub m_UserHandler_UserBeforeChange(ByRef NewUserName As String, Cancel As Boolean)
   If StrPtr(m_NewUserName) <> 0 Then
      NewUserName = m_NewUserName
   End If
End Sub

Private Sub m_UserHandler_UserChanged(ByVal NewUserName As String)
   m_NewUserName = NewUserName
End Sub



Public Sub Roles_UseDummy_ReturnsSameObject()

   Dim Expected As UserRoles
   Dim Actual As UserRoles

   Set Expected = m_UserRolesDummy
   Set Actual = m_UserHandler.Roles
   
   Assert.That Actual, Iz.SameAs(Expected)

End Sub

Private Sub m_UserHandler_GetUserRoles(UserRolesRef As UserRoles)
   Set UserRolesRef = m_UserRolesDummy
End Sub

'AccUnit:Row("Role1", True)
'AccUnit:Row("MissingRole", False)
Public Sub IsRoleMember_UseDummy_ReturnsTrueOrFalse(ByVal Role As String, ByVal Expected As Boolean)

   Dim Actual As Variant
   Actual = m_UserHandler.IsRoleMember(Role)
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub


'--------------------------------------------------------------------
' Test support
'--------------------------------------------------------------------

Private Sub InitUserHandler()
   Set m_UserHandler = New UserHandler
   m_UserHandler.UserName = UserName
   InitUserRolesDummy
End Sub

Private Sub DisposeUserHandler()
   Set m_UserRolesDummy = Nothing
   Set m_UserHandler = Nothing
End Sub

Private Sub InitUserRolesDummy()

   Set m_UserRolesDummy = New UserRolesDummy
   m_UserRolesDummy.Add "Role1", "Role2", "Role3"
   m_UserRolesDummy.Add "Role4", "Role5"

   Set m_UserRolesDummy.User = m_UserHandler

End Sub
