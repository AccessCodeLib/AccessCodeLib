VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorTrappingObserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: vba.ErrorTrappingObserver
'---------------------------------------------------------------------------------------
'
' Auxiliary class for error trapping manipulation
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>vba/ErrorTrappingObserver.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Private Const ErrorTrappingOptionName As String = "Error Trapping"

Public Enum ErrorTrappingOptions
   BreakOnAllErrors = 0
   BreakInClassModule = 1
   BreakOnUnhandledErrors = 2
End Enum

Private m_InitialErrorTrapping As ErrorTrappingOptions
Private m_KeepNewValueOnTerminate As Boolean

Private Sub Class_Initialize()
   m_InitialErrorTrapping = Application.GetOption(ErrorTrappingOptionName)
End Sub

Private Sub Class_Terminate()
   If Not m_KeepNewValueOnTerminate Then
      RestoreErrorTrapping
   End If
End Sub

Public Property Get KeepNewValueOnTerminate() As Boolean
   KeepNewValueOnTerminate = m_KeepNewValueOnTerminate
End Property

Public Property Let KeepNewValueOnTerminate(ByVal NewValue As Boolean)
   m_KeepNewValueOnTerminate = NewValue
End Property

Public Property Get ErrorTrapping() As ErrorTrappingOptions
   ErrorTrapping = Application.GetOption(ErrorTrappingOptionName)
End Property

Public Property Let ErrorTrapping(ByVal NewValue As ErrorTrappingOptions)
    If ErrorTrapping <> NewValue Then
        Application.SetOption ErrorTrappingOptionName, NewValue
    End If
End Property

Private Sub RestoreErrorTrapping()
    If m_InitialErrorTrapping <> ErrorTrapping Then
        Application.SetOption ErrorTrappingOptionName, m_InitialErrorTrapping
    End If
End Sub
