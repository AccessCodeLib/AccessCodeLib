Attribute VB_Name = "_VBATestSuite"
'---------------------------------------------------------------------------------------
' Module: _VBATestSuite (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsmodul für den Aufruf einer anwendungsweiten Instanz von VBATestSuite
' </summary>
' <remarks>
' Benötigt <a href="http://sourceforge.net/projects/simplyvbunit/">SimplyVBUnit 3.0</a>
'
'
'Test-Klasse(n) anhängen:                                                             \n
'  TestSuite.Add new DataTestClass,  new OtherTestClass                               \n
'
'Tests ausführen:                                                                     \n
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
' Property: TestSuite (Josef Pötzl, 2010-06-20)
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
