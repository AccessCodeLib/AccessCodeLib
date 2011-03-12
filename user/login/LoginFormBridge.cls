VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoginFormBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LoginFormBridge
'---------------------------------------------------------------------------------------
'/**
' @brief Login-Prüfung mit Formular
'
' Loginvariante muss von außen eingestellt werden ... Übergabe von LoginChecker
'
'**/
'<codelib>
'  <file>user/login/LoginFormBridge.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>user/login/LoginForm.frm</use>
'</codelib>
'---------------------------------------------------------------------------------------'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const LoginFormName As String = "LoginForm"

Private m_LoginChecker As Object 'Late binding, damit unterschiedliche Loginvarianten verwendet werden können
                                 'Schöner wäre ein Interface, damit machte ich aber schlechte Erfahrungen unter VBA (Abstürze & Co. während der Entwicklung)
Private WithEvents m_LoginForm As Form_LoginForm
Attribute m_LoginForm.VB_VarHelpID = -1
Private m_LoginCancelled As Boolean
Private m_Caption As String

Public Property Set LoginChecker(ByVal LoginCheckerRef As Object)
   Set m_LoginChecker = LoginCheckerRef
End Property

Public Property Let Caption(ByVal FormCaption As String)
   m_Caption = FormCaption
End Property

Public Property Set LoginForm(ByVal FormRef As Form_LoginForm)
   Set m_LoginForm = FormRef
End Property

Public Function Login(ByVal UserName As String) As Boolean
   
   Dim DefaultUser As String
   Dim OpenArgString As String
   Dim Caption As String

   If Len(UserName) > 0 Then
      DefaultUser = UserName
   ElseIf Len(m_LoginChecker.CurrentUser) > 0 Then
      DefaultUser = m_LoginChecker.CurrentUser
   Else
      DefaultUser = Environ("username")
   End If
   
   OpenArgString = CStr(ObjPtr(Me)) & "@"
   If Len(m_Caption) > 0 Then
      Caption = m_Caption
   Else
      Caption = "Login"
   End If
   OpenArgString = OpenArgString & Caption & "@" & DefaultUser

   DoCmd.OpenForm LoginFormName, acNormal, , , acFormEdit, acDialog, OpenArgString
   '=> Im LoginFormular erfolgt Aufruf der Eigenschaft "LoginForm",
   '   damit wird Ereignisbehandlung aktiviert => m_LoginCancelled wird eingestellt
   Login = Not m_LoginCancelled

End Function

'Ereignisbehandlung vom Login-Formular
Private Sub m_LoginForm_cancelled()
   m_LoginCancelled = True
End Sub

Private Sub m_LoginForm_Logon(ByVal LoginName As String, ByVal Password As String, Cancel As Boolean)

   Dim LoginOk As Boolean
   
   m_LoginCancelled = False
   
   LoginOk = m_LoginChecker.Login(LoginName, Password)
   If Not LoginOk Then
      Cancel = True
   End If

End Sub
