VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ApplicationHandler_TracingTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: ApplicationHandler_TracingTest (Josef Pötzl, 2010-07-03)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ApplicationHandler_Tracing-Test
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/base/log/ApplicationHandler_TracingTest.cls</file>
'  <use>base/log/ApplicationHandler_Tracing.cls</use>
'  <use>test/simplyvbunit/VBATestSuite.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Aktivierung für TestSuite.RunAll
'<VBATestSuite><testclass>true</testclass></VBATestSuite>


'---------------------------------------------------------------------------------------
' Grundkonfiguration für AppHandler.WriteLog-Test
'

Private AppHandler As ApplicationHandler
Private Tracing As ApplicationHandler_Tracing

'Ersatz für _config_Application.InitConfig
Private Sub initApplicationHandlerTracing()
   Set AppHandler = New ApplicationHandler
   Set Tracing = New ApplicationHandler_Tracing
   Set Tracing.ApplicationHandlerRef = AppHandler
End Sub


'Ersatz für modApplication.TraceLog
Private Sub writeTraceLog(ByRef Msg As String, ParamArray Args() As Variant)
   AppHandler.WriteLog Msg, ApplicationHandlerLogType.AppLogType_Tracing, Args, False
End Sub


'---------------------------------------------------------------------------------------
' TESTS
'

Public Sub WriteTraceLog2MemoryAndDebug()

   Dim testTraces As ApplicationHandler_Tracing
   Set testTraces = New ApplicationHandler_Tracing
   
   testTraces.OutputMode = TracingOutputMode_Memory + TracingOutputMode_Debug
   
   Dim LogMsg As String
   Dim LogParam(1 To 3) As Variant
   
   Debug.Print "--- Trace log (Start) ---"
   
   LogMsg = "TEXT1"
   LogParam(1) = "abc"
   LogParam(2) = 123
   LogParam(3) = #12/24/2010#
   testTraces.Log LogMsg, LogParam(1), LogParam(2), LogParam(3)
   
   LogMsg = "TEXT2"
   LogParam(1) = "Nur ein Parameter"
   testTraces.Log LogMsg, LogParam(1)

   Debug.Print "--- Trace log (Ende) ---"
   
   '-----------------------------------------
   ' Prüfungen
   '
   'Anzahl=1
   Assert.That testTraces.TraceLogCollection.Count, Iz.EqualTo(2)
   
   Dim traceLogArray As Variant
   
   ' Msg 1
   traceLogArray = testTraces.TraceLogCollection.Item(1)
   Assert.That traceLogArray(0), Iz.EqualTo("TEXT1")
   ' Param1
   Assert.That traceLogArray(1)(0), Iz.EqualTo("abc")
   ' Param2
   Assert.That traceLogArray(1)(1), Iz.EqualTo(123)
   ' Param3
   Assert.That traceLogArray(1)(2), Iz.EqualTo(#12/24/2010#)
   
   ' Msg 2
   traceLogArray = testTraces.TraceLogCollection.Item(2)
   Assert.That traceLogArray(0), Iz.EqualTo("TEXT2")
   ' Param1
   Assert.That traceLogArray(1)(0), Iz.EqualTo("Nur ein Parameter")


End Sub

Public Sub WriteTraceLog2MemoryWithAppHandler()

   initApplicationHandlerTracing
   
   Tracing.OutputMode = TracingOutputMode_Memory
   
   writeTraceLog "TEXT1", 123, "abc", #12/24/2010#
   writeTraceLog "TEXT2", "Nur ein Parameter"
   
   '-----------------------------------------
   ' Prüfungen
   '
   'Anzahl=2
   Assert.That Tracing.TraceLogCollection.Count, Iz.EqualTo(2)
   
   Dim traceLogArray As Variant
   
   '1. Msg
   traceLogArray = Tracing.TraceLogCollection.Item(1)
   Assert.That traceLogArray(0), Iz.EqualTo("TEXT1")
   ' Param1
   Assert.That traceLogArray(1)(0), Iz.EqualTo(123)
   ' Param2
   Assert.That traceLogArray(1)(1), Iz.EqualTo("abc")
   ' Param3
   Assert.That traceLogArray(1)(2), Iz.EqualTo(#12/24/2010#)
   
   '2. Msg
   traceLogArray = Tracing.TraceLogCollection.Item(2)
   Assert.That traceLogArray(0), Iz.EqualTo("TEXT2")
   ' Param1
   Assert.That traceLogArray(1)(0), Iz.EqualTo("Nur ein Parameter")

End Sub
