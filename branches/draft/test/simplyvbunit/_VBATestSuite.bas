Attribute VB_Name = "_VBATestSuite"
'---------------------------------------------------------------------------------------
' Module: _VBATestSuite (Josef P�tzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsmodul f�r den Aufruf einer anwendungsweiten Instanz von VBATestSuite
' </summary>
' <remarks>
' Ben�tigt <a href="http://sourceforge.net/projects/simplyvbunit/">SimplyVBUnit 3.0</a>
'
'
'Test-Klasse(n) anh�ngen:                                                             \n
'  TestSuite.Add new DataTestClass,  new OtherTestClass                               \n
'
'Tests ausf�hren:                                                                     \n
'  TestSuite.Run                                                                      \n
'
'
'Oder alles in einer Zeile:                                                           \n
'  TestSuite.Reset(True).Add(new DataTestClass, new OtherTestClass).Run               \n
'
'Alle Testklassen der aktiven Anwendunge (VbProject) verwenden:                       \n
'  TestSuite.Reset(True).AddAllFromVbProject().Run                                    \n
'Kurzfassung: TestSuite.RunAll
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>test/simplyvbunit/_VBATestSuite.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>test/simplyvbunit/VBATestSuite.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_TestSuite As VBATestSuite

'---------------------------------------------------------------------------------------
' Property: TestSuite (Josef P�tzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt eine anwendungsweite Instanz von VBATestSuite
' </summary>
' <returns>VBATestSuite</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TestSuite() As VBATestSuite
   If m_TestSuite Is Nothing Then
      Set m_TestSuite = New VBATestSuite
   End If
   Set TestSuite = m_TestSuite
End Property
