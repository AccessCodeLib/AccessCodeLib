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
' Class Module: AccessProgressBar
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Hilfsklasse zur Verwendung der integrierten Access ProgressBar ("Meter")
' \ingroup utils
'**/
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
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Initialisiert die Progressbar für einen neuen Vorgang
' </summary>
' <param name="Text">Text der neben der Progressbar angezeigt werden soll</param>
' <param name="StepCount">Gesamtanzahl der Schritte, Default 100</param>
' <param name="Step">Schrittweite, Default: 1</param>
' <returns></returns>
' <remarks>
' Im Anschluss kann die PerformStep() Methode aufgerufen werden
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
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
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Startet einen Vorgang
' </summary>
' <returns>Boolean</returns>
' <remarks>
' Optional, wird bei Bedarf durch die Methode PerformStep() ausgeführt.
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
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
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Führt einen Iterationsschritt aus und "bewegt" die Progressbar einen Schritt weiter.
' </summary>
' <param name="ClearOnLastStep">Soll beim letzten Schritt automatisch die Clear() Methode aufgerufen werden?</param>
' <returns>Aktuellen Schritt als Long</returns>
' <remarks>
' Ruft automatisch die Start() Methode auf, sofern die Progressbar noch nicht initialisiert worden ist.
' Der in der Eigenschaft ValueMax angegebene Maximalwert kann nicht überschritten werden.
' Sofern nicht anders angegeben, wird beim letzten Schritt automatisch die Clear() Methode aufgerufen.
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
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
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Zurücksetzen der ProgressBar und ggf. Abbrechen laufender Vorgänge
' </summary>
' <returns>Boolean</returns>
' <remarks>
' Alle Eigenschaften werden auf die Standard-Vorgaben zurückgesetzt
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
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
'/**
' <summary>
' Gibt an ob die Access Progressbar initialisiert ist
' </summary>
' <returns>Boolean</returns>
' <remarks>
' Bei True wird die Progressbar gerade verwendet. Der aktuell laufende Vorgang kann
' durch Aufruf der Clear() Methode beendet werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get IsInitialized() As Boolean
    IsInitialized = m_IsInitialized
End Property

'---------------------------------------------------------------------------------------
' Property: Step
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Schrittweite
' </summary>
' <returns>Long</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Step() As Long
    Step = m_ProgressBarStep
End Property

'---------------------------------------------------------------------------------------
' Property: Step
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Schrittweite
' </summary>
' <param name="Value">Schrittweite die verwendet werden soll</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Step(ByVal Value As Long)
    m_ProgressBarStep = Value
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMax
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Maximalwert der Progressbar bzw. Endwert
' </summary>
' <returns>Long</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ValueMax() As Long
    ValueMax = m_ProgressBarValueMax
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMax
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Maximalwert der Progressbar bzw. Endwert
' </summary>
' <param name="Value">Maximalwert der Progressbar bzw. Endwert der verwendet werden soll</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ValueMax(ByVal Value As Long)
    m_ProgressBarValueMax = Value
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMin
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Minimalwet der Progressbar bzw. Startwert
' </summary>
' <returns>Long</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ValueMin() As Long
    ValueMin = m_ProgressBarValueMin
End Property

'---------------------------------------------------------------------------------------
' Property: ValueMin
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Minimalwet der Progressbar bzw. Startwert
' </summary>
' <param name="Value">Minimalwet der Progressbar bzw. Startwert der verwendet werden soll</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ValueMin(ByVal Value As Long)
    m_ProgressBarValueMin = Value
End Property

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Beschreibungstext
' </summary>
' <returns>String</returns>
' <remarks>
' Wird links neben der Progressbar angezeigt
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Property Get Text() As String
    Text = m_ProgressBarText
End Property

'---------------------------------------------------------------------------------------
' Property: Text
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Beschreibungstext
' </summary>
' <param name="Value">Beschreibungstext der angezeigt werden soll</param>
' <returns>String</returns>
' <remarks>Wird links neben der Progressbar angezeigt</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let Text(ByVal Value As String)
    m_ProgressBarText = Value
End Property
