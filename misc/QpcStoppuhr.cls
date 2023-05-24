VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "QpcStoppuhr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: QpcStoppuhr
'---------------------------------------------------------------------------------------
'/**
' <summary>
' QueryPerformanceCounter-Stoppuhr
' zur Zeitmessung
' </summary>
' <remarks>
' Nutzt die API-Funktionen QueryPerformanceCounter und QueryPerformanceFrequency
' aus Kernel32 zum Messen der Zeit.
' </remarks>
' \ingroup misc
'**/
'<codelib>
'  <file>misc/QpcStoppuhr.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

' Messung mit API
' QueryPerformanceCounter


'temp. Messergebnis
Private p_TempTime As Currency
Private p_TestFrequ As Currency

'Zwischenspeicher für Laufzeit
Private p_Timer As Currency

'Zeit bei Start
Private p_Timer_Start As Currency

'Win-API
Private Declare PtrSafe Function QueryPerformanceCounter Lib "kernel32" (QPC As Currency) As Boolean
Private Declare PtrSafe Function QueryPerformanceFrequency Lib "kernel32" (QPF As Currency) As Boolean
                        
'Klasse initialisieren
Private Sub Class_Initialize()
    p_Timer_Start = 0
    p_Timer = 0
    QueryPerformanceFrequency p_TestFrequ
End Sub

'---------------------------------------------------------------------------------------
' Sub: StartT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Stoppuhr starten/neustarten, beginnen mit 0
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub StartT()
    p_Timer = 0
    QueryPerformanceCounter p_Timer_Start
End Sub

'---------------------------------------------------------------------------------------
' Sub: StopT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Stoppuhr anhalten
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub StopT()
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning Then
        p_Timer = p_Timer + p_TempTime - p_Timer_Start
        p_Timer_Start = 0
    End If
End Sub


'---------------------------------------------------------------------------------------
' Sub: RestartT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Stoppuhr weiterlaufen lassen
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RestartT()
    'Uhr muss stehen!
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning = False Then
        p_Timer_Start = p_TempTime
    End If
End Sub

'---------------------------------------------------------------------------------------
' Sub: ResetT
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alles auf null
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub ResetT()
    p_Timer_Start = 0
    p_Timer = 0
End Sub

'---------------------------------------------------------------------------------------
' Property: IsRunning
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an, ob die Stoppuhr noch läuft
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get IsRunning() As Boolean
    IsRunning = (p_Timer_Start > 0)
End Property

'---------------------------------------------------------------------------------------
' Property: Time
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zeit in Millisekunden
' </summary>
' <returns>Currency</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Time() As Currency
    QueryPerformanceCounter p_TempTime
    If Me.IsRunning Then
        Time = (p_Timer + p_TempTime - p_Timer_Start) * 1000 / p_TestFrequ  ' [ms]
   Else
        Time = p_Timer * 1000 / p_TestFrequ
    End If
End Property

'---------------------------------------------------------------------------------------
' Property: TimeString (2009-12-09)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zeit als String mit Formatierung
' </summary>
' <param name="sSecFormatString">Formatierung der Zeit inkl. Sekunden aber ohne Sekundenbruchteile</param>
' <param name="iMSecDigits">Anzahl der Stellen für die Millisekunden (nach dem Komma)</param>
' <returns>String</returns>
' <remarks>
' Beispiel: TimeString("hh:nn:ss", 5) ergibt "00:05:34,24532"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TimeString(Optional sSecFormatString As String = "hh:nn:ss", _
                                Optional iMSecDigits As Long = 3) As String
    
'sSecFormatString:      Format für Uhrzeit ohne Sekundenbruchteile
'iMSecDigits:           Anzahl Stellen nach dem Komma. z.B.: 2 = 1/100-Sek.-Anzeige

    Dim curTimeSec As Currency
    Dim strTimeString As String

    curTimeSec = VBA.Round(Me.Time / 1000, iMSecDigits) 'mathematisches Runden ist ok
   
    strTimeString = Format(TimeSerial(0, 0, Fix(curTimeSec)), sSecFormatString)
    If iMSecDigits > 0 Then
        strTimeString = strTimeString & Format(curTimeSec - Fix(curTimeSec), "." & String(iMSecDigits, "0"))
    End If
    
    TimeString = strTimeString
    
End Property
