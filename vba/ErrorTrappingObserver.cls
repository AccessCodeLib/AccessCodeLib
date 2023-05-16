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
'<codelib>
'  <file>vba/ErrorTrappingObserver.cls</file>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

Public Enum ErrorTrapping
    BreakOnAllErrors = 0
    BreakInClassModule = 1
    BreakOnUnhandledErrors = 2
End Enum

Private m_InitialErrorTrapping As ErrorTrapping

Private Sub Class_Initialize()
   m_InitialErrorTrapping = Application.GetOption("Error Trapping")
End Sub

Private Sub Class_Terminate()
   RestoreErrorTrapping
End Sub

Public Property Get ErrorTrapping() As ErrorTrapping
   ErrorTrapping = Application.GetOption("Error Trapping")
End Property

Public Property Let ErrorTrapping(ByVal NewValue As ErrorTrapping)
    If ErrorTrapping <> NewValue Then
        Application.SetOption "Error Trapping", NewValue
    End If
End Property

Private Sub RestoreErrorTrapping()
    If m_InitialErrorTrapping <> ErrorTrapping Then
        Application.SetOption "Error Trapping", m_InitialErrorTrapping
    End If
End Sub
