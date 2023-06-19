VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: utils.AccessProgressBar
'---------------------------------------------------------------------------------------
'
' Helper class for using the integrated Access ProgressBar ("Meter")
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>utils/AccessProgressBar.cls</file>
'  <test>_test/utils/AccessProgressBarTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Result As Variant
Private m_IsInitialized As Boolean
Private m_ProgressBarValueMin As Long
Private m_ProgressBarValueMax As Long
Private m_ProgressBarStep As Long
Private m_ProgressBarCurrStep As Long
Private m_ProgressBarText As String

'---------------------------------------------------------------------------------------
' Function: Init
'---------------------------------------------------------------------------------------
'
' Initializes the progress bar for a new operation
'
' Parameters:
'     Text        - Text to be displayed next to the progress bar
'     StepCount   - Total number of steps, default: 100
'     Step        - Step size, default: 1
'
' Remarks:
'     Afterwards the PerformStep() method can be called
'
'---------------------------------------------------------------------------------------
Public Function Init(ByVal Text As String, Optional ByVal StepCount As Long = 100, Optional ByVal Step As Integer = 1) As Boolean
    AccessProgressBar.Clear
    AccessProgressBar.Text = Text
    AccessProgressBar.ValueMin = 1
    AccessProgressBar.ValueMax = StepCount
    AccessProgressBar.Step = Step
    m_ProgressBarCurrStep = 0
    Init = True
End Function

'---------------------------------------------------------------------------------------
' Function: Start
'---------------------------------------------------------------------------------------
'
' Starts an operation
'
' Returns:
'     Boolean
'
' Remarks:
'     Optional, executed by the PerformStep() method if required.
'
'---------------------------------------------------------------------------------------
Public Function Start() As Boolean
    If Not m_IsInitialized Then
        m_IsInitialized = True
        m_ProgressBarCurrStep = 0
        m_Result = Access.Application.SysCmd(acSysCmdInitMeter, m_ProgressBarText, m_ProgressBarValueMax)
        Start = True
    Else
        Start = False
    End If
End Function

'---------------------------------------------------------------------------------------
' Function: PerformStep
'---------------------------------------------------------------------------------------
'
' Executes an iteration step and "moves" the progressbar one step further.
'
' Parameters:
'     ClearOnLastStep - Should the Clear() method be called automatically at the last step?
'
' Returns:
'     Long - Current step
'
' Remarks:
'     Automatically calls the Start() method if the progressbar has not yet been initialized.
'     The maximum value specified in the ValueMax property cannot be exceeded.
'     Unless otherwise specified, the Clear() method is automatically called at the last step.
'
'---------------------------------------------------------------------------------------
Public Function PerformStep(Optional ByVal ClearOnLastStep = True) As Long
    
    If Not m_IsInitialized Then AccessProgressBar.Start
    
    If m_ProgressBarCurrStep < m_ProgressBarValueMax Then _
       m_ProgressBarCurrStep = m_ProgressBarCurrStep + m_ProgressBarStep

    m_Result = Access.Application.SysCmd(acSysCmdUpdateMeter, m_ProgressBarCurrStep)
    PerformStep = m_ProgressBarCurrStep
    
    If ClearOnLastStep And m_ProgressBarCurrStep = m_ProgressBarValueMax Then AccessProgressBar.Clear
    
End Function

'---------------------------------------------------------------------------------------
' Function: Clear
'---------------------------------------------------------------------------------------
'
' Resetting the ProgressBar and, if necessary, canceling running processes
'
' Returns:
'     Boolean
'
' Remarks:
'     All properties are reset to the default settings
'
'---------------------------------------------------------------------------------------
Public Function Clear() As Boolean
    m_Result = Access.Application.SysCmd(acSysCmdClearStatus)
    m_IsInitialized = False
    m_ProgressBarValueMin = 1
    m_ProgressBarValueMax = 100
    m_ProgressBarStep = 1
    m_ProgressBarCurrStep = 1
    m_ProgressBarText = vbNullString
    Clear = Not m_IsInitialized
End Function

'---------------------------------------------------------------------------------------
' Property: IsInitialized
'---------------------------------------------------------------------------------------
'
' Indicates whether the Access Progressbar is initialized.
'
' Returns:
'     Boolean
'
' Remarks:
'     If True, the progress bar is currently being used.
'     The currently running operation can be terminated by calling the Clear() method.
'
'---------------------------------------------------------------------------------------
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

'---------------------------------------------------------------------------------------
' Property: Step
'---------------------------------------------------------------------------------------
'
' Step size
'
'---------------------------------------------------------------------------------------
Public Property Get Step() As Long
    Step = m_ProgressBarStep
End Property

Public Property Let Step(ByVal Value As Long)
    m_ProgressBarStep = Value
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMax
'---------------------------------------------------------------------------------------
'
' Maximum value of the progress bar or final value
'
'---------------------------------------------------------------------------------------
Public Property Get ValueMax() As Long
    ValueMax = m_ProgressBarValueMax
End Property

Public Property Let ValueMax(ByVal Value As Long)
    m_ProgressBarValueMax = Value
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMin
'---------------------------------------------------------------------------------------
'
' Minimum value of the progress bar or start value
'
'---------------------------------------------------------------------------------------
Public Property Get ValueMin() As Long
    ValueMin = m_ProgressBarValueMin
End Property

Public Property Let ValueMin(ByVal Value As Long)
    m_ProgressBarValueMin = Value
End Property

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'
' Description text
'
'---------------------------------------------------------------------------------------
Public Property Get Text() As String
    Text = m_ProgressBarText
End Property

Public Property Let Text(ByVal Value As String)
    m_ProgressBarText = Value
End Property
