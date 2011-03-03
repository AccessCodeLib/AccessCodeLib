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
' @brief Login-Pr�fung mit Formular
'
' Loginvariante muss von au�en eingestellt werden ... �bergabe von LoginChecker
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

Const LoginFormName As String = "LoginForm"

Private m_LoginChecker As Object 'Late binding, damit unterschiedliche Loginvarianten verwendet werden k�nnen
                                 'Sch�ner w�re ein Interface, damit machte ich aber schlechte Erfahrungen unter VBA (Abst�rze & Co. w�hrend der Entwicklung)
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

Public Property Set LoginForm(ByRef FormRef As Form_LoginForm)
   Set m_LoginForm = FormRef
End Property

Public Function Login(ByVal UserName As String) As Boolean
   
   Dim strDefaultUser As String
   Dim openArgString As String

   If Len(UserName) > 0 Then
      strDefaultUser = UserName
   ElseIf Len(m_LoginChecker.CurrentUser) > 0 Then
      strDefaultUser = m_LoginChecker.CurrentUser
   Else
      strDefaultUser = Environ("username")
   End If
   
   openArgString = CStr(ObjPtr(Me)) & "@"
   If Len(m_Caption) > 0 Then
      openArgString = openArgString & m_Caption & " - "
   End If
   openArgString = openArgString & "Login@" & strDefaultUser

   DoCmd.OpenForm LoginFormName, acNormal, , , acFormEdit, acDialog, openArgString
   '=> Im LoginFormular erfolgt Aufruf der Eigenschaft "LoginForm",
   '   damit wird Ereignisbehandlung aktiviert => m_LoginCancelled wird eingestellt
   Login = Not m_LoginCancelled

End Function

'Ereignisbehandlung vom Login-Formular
Private Sub m_LoginForm_cancelled()
   m_LoginCancelled = True
End Sub

Private Sub m_LoginForm_Logon(ByVal LoginName As String, ByVal Password As String, Cancel As Boolean)

   Dim bolLoginOk As Boolean
   
   m_LoginCancelled = False
   
   bolLoginOk = m_LoginChecker.Login(LoginName, Password)
   If Not bolLoginOk Then
      Cancel = True
   End If

End Sub