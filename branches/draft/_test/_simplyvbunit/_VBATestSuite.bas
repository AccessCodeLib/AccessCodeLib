Attribute VB_Name = "_VBATestSuite"
'---------------------------------------------------------------------------------------
' Module: _VBATestSuite (Josef Pötzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfssmodul für den Aufruf einer anwendungsweiten Instanz von VBATestSuite
' </summary>
' <remarks>
' Benötigt <a href="http://sourceforge.net/projects/simplyvbunit/">SimplyVBUnit 3.0</a>
'
'
'Test-Klasse(n) anhängen:                                                             \n
'  MyTestSuite.Add new DataTestClass,  new OtherTestClass                             \n
'
'Tests ausführen:                                                                     \n
'  MyTestSuite.Run                                                                    \n
'
'
'Oder alles in einer Zeile:                                                           \n
'  MyTestSuite.Reset(True).Add(new DataTestClass, new OtherTestClass).Run             \n
'
'Alle Testklassen der aktiven Anwendunge (VbProject) verwenden:                       \n
'  MyTestSuite.Reset(True).AddAllFromVbProject().Run                                  \n
'Kurzfassung: MyTestSuite.RunAll
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/_simplyvbunit/_VBATestSuite.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>_test/_simplyvbunit/VBATestSuite.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Test-Klasse(n) anhängen:
'  MyTestSuite.Add new DataTestClass,  new OtherTestClass
'
'Tests ausführen:
'  MyTestSuite.Run
'
'
' oder alles in einer Zeile:
' MyTestSuite.Reset(True).Add(new DataTestClass, new OtherTestClass).Run
'
' alle Testklassen der aktiven Anwendunge (VbProject) verwenden:
' MyTestSuite.Reset(True).AddAllFromVbProject().Run
' Kurzfassung: MyTestSuite.RunAll
'

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
