Attribute VB_Name = "DateTime2_Example"
'---------------------------------------------------------------------------------------
' Class Module: DateTime2_Example
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Beispiel zur Verwendung der DateTime2 Klasse
' \ingroup datetime
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>datetime/DateTime2_Example.bas</file>
'  <use>datetime/DateTime2.cls</use>
'  <test>_test\datetime\DateTime2Tests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub DateTime2_Example()

    'Verwendung als statische Klasse
    Debug.Print DateTime2.NewValue("01.01.2012").AddMonth(1).LastDayOfMonth.ToDate '29.02.2012

    'Verwendung als Objektvariable
    Dim Jetzt As DateTime2
    Set Jetzt = DateTime2.Now.TrimTime
    
    Dim Morgen As DateTime2
    Set Morgen = Jetzt.AddDay(1)
    
    Debug.Print "Weil heute " & Jetzt.GetWeekDayNameShort & ", der " & _
                Jetzt & " ist, ist Morgen " & Morgen.GetWeekDayName & ", der " & Morgen & "."
    
    Set Jetzt = Nothing
    Set Morgen = Nothing
    
End Sub
