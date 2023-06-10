VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QpcStopwatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: misc.QpcStopwatch
'---------------------------------------------------------------------------------------
'
' QueryPerformanceCounter stopwatch for time measurement
'
' Author:
'     Josef Poetzl
'
' Remarks:
'  Uses the API functions QueryPerformanceCounter and QueryPerformanceFrequency from Kernel32 to measure time.
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>misc/QpcStopwatch.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'temp. measurement result
Private p_TempTime As Currency
Private p_TestFrequ As Currency

'Cache for runtime
Private p_Timer As Currency

'Time at start
Private p_Timer_Start As Currency

'Win API
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (QPC As Currency) As Boolean
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (QPF As Currency) As Boolean
                        
Private Sub Class_Initialize()
    p_Timer_Start = 0
    p_Timer = 0
    QueryPerformanceFrequency p_TestFrequ
End Sub

'---------------------------------------------------------------------------------------
' Sub: StartT
'---------------------------------------------------------------------------------------
'
' Start/restart stopwatch, start with 0
'
'---------------------------------------------------------------------------------------
Public Sub StartT()
    p_Timer = 0
    QueryPerformanceCounter p_Timer_Start
End Sub

'---------------------------------------------------------------------------------------
' Sub: StopT
'---------------------------------------------------------------------------------------
'
' Stop timer
'
'---------------------------------------------------------------------------------------
Public Sub StopT()
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning Then
        p_Timer = p_Timer + p_TempTime - p_Timer_Start
        p_Timer_Start = 0
    End If
End Sub

'---------------------------------------------------------------------------------------
' Sub: Restart
'---------------------------------------------------------------------------------------
'
' Keep stopwatch running
'
'---------------------------------------------------------------------------------------
Public Sub Restart()
    'Uhr muss stehen!
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning = False Then
        p_Timer_Start = p_TempTime
    End If
End Sub

'---------------------------------------------------------------------------------------
' Sub: Reset
'---------------------------------------------------------------------------------------
'
' Everything to zero
'
'---------------------------------------------------------------------------------------
Public Sub Reset()
    p_Timer_Start = 0
    p_Timer = 0
End Sub

'---------------------------------------------------------------------------------------
' Property: IsRunning
'---------------------------------------------------------------------------------------
'
' Indicates whether the stopwatch is still running
'
'---------------------------------------------------------------------------------------
Public Property Get IsRunning() As Boolean
    IsRunning = (p_Timer_Start > 0)
End Property

'---------------------------------------------------------------------------------------
' Property: TimeMSec
'---------------------------------------------------------------------------------------
'
' Time in milliseconds
'
'---------------------------------------------------------------------------------------
Public Property Get TimeMSec() As Currency
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning Then
        TimeMSec = (p_Timer + p_TempTime - p_Timer_Start) * 1000 / p_TestFrequ  ' [ms]
   Else
        TimeMSec = p_Timer * 1000 / p_TestFrequ
    End If
End Property

'---------------------------------------------------------------------------------------
' Function: TimeString
'---------------------------------------------------------------------------------------
'
' Time as string with formatting
'
' Parameters:
'     FormatString   - Formatting of time incl. seconds but without fractions of seconds
'     SecDigits      - Number of digits for the fractions of a second (after the decimal point)
'
' Returns:
'     String
'
' Example:
'     TimeString("hh:nn:ss", 5) return formatted string like "00:05:34,24532"
'
'---------------------------------------------------------------------------------------
Public Function TimeString(Optional FormatString As String = "hh:nn:ss", _
                           Optional SecDigits As Long = 3) As String

    Dim TimeSec As Currency
    Dim FormattedTime As String

    TimeSec = VBA.Round(Me.TimeMSec / 1000, SecDigits)
   
    FormattedTime = Format(TimeSerial(0, 0, Fix(TimeSec)), FormatString)
    If SecDigits > 0 Then
        FormattedTime = FormattedTime & Format(TimeSec - Fix(TimeSec), "." & String(SecDigits, "0"))
    End If
    
    TimeString = FormattedTime
    
End Function
