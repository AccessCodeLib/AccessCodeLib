Attribute VB_Name = "_VBATestSuite"
'---------------------------------------------------------------------------------------
' Module: _VBATestSuite (Josef P�tzl, 2010-06-20)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfssmodul f�r den Aufruf einer anwendungsweiten Instanz von VBATestSuite
' </summary>
' <remarks>
' Ben�tigt <a href="http://sourceforge.net/projects/simplyvbunit/">SimplyVBUnit 3.0</a>
'
'
'Test-Klasse(n) anh�ngen:                                                             \n
'  MyTestSuite.Add new DataTestClass,  new OtherTestClass                             \n
'
'Tests ausf�hren:                                                                     \n
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

'Test-Klasse(n) anh�ngen:
'  MyTestSuite.Add new DataTestClass,  new OtherTestClass
'
'Tests ausf�hren:
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
