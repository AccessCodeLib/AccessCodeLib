VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: utils.AccessStatusBar
'---------------------------------------------------------------------------------------
'
' Helper class for using the Access status bar
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>utils/AccessStatusBar.cls</file>
'  <test>_test/utils/AccessStatusBarTests.cls</test>
'  <license>_codelib/license.bas</license>
'  <description>Sets and clears the text of the Access status bar</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Text As String
Private m_Result As Variant

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'
' Description text
'
' Returns:
'     String
'
' Remarks:
'     Is displayed in the status bar
'
'---------------------------------------------------------------------------------------
Public Property Get Text() As String
    Text = m_Text
End Property

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'
' Description text
'
' Parameters:
'     Value - Description text to be displayed
'
' Remarks:
'     Is displayed in the status bar
'
'---------------------------------------------------------------------------------------
Public Property Let Text(ByVal Value As String)
    m_Text = Value
    m_Result = Access.Application.SysCmd(acSysCmdSetStatus, m_Text)
End Property

'---------------------------------------------------------------------------------------
' Function: Clear
'---------------------------------------------------------------------------------------
'
' Reset the status bar
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function Clear() As Boolean
    m_Result = Access.Application.SysCmd(acSysCmdClearStatus)
    Clear = True
End Function
