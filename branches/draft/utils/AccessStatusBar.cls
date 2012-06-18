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
' Class Module: AccessStatusBar
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Hilfsklasse zur Verwendung der Access Status-Zeile
' \ingroup utils
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>utils\AccessStatusBar.cls</file>
'  <test>_test\utils\AccessStatusBarTests.cls</test>
'  <license>_codelib\license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Text As String
Private m_Result As Variant

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beschreibungstext
' </summary>
' <returns>String</returns>
' <remarks>Wird auf der Statuszeile angezeigt</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Text() As String
    Text = m_Text
End Property

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beschreibungstext
' </summary>
' <param name="Value">Beschreibungstext der angezeigt werden soll</param>
' <remarks>Wird auf der Statuszeile angezeigt</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Text(ByVal Value As String)
    m_Text = Value
    m_Result = Access.Application.SysCmd(acSysCmdSetStatus, m_Text)
End Property

'---------------------------------------------------------------------------------------
' Function: Clear
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Zurücksetzen der Statuszeile
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Function Clear() As Boolean
    m_Result = Access.Application.SysCmd(acSysCmdClearStatus)
    Clear = True
End Function
