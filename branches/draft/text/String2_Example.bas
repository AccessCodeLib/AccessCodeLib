Attribute VB_Name = "String2_Example"
'---------------------------------------------------------------------------------------
' Class Module: DateTime2_Example
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Beispiel zur Verwendung der String2 Klasse
' \ingroup datetime
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text\String2_Example.bas</file>
'  <use>text\String2.cls</use>
'  <test>_test\text\String2Tests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub String2_Example()

    'Verwendung als statische Klasse
    Debug.Print String2.NewValue("   Hello").Append(" ").Append("World  ").Trim().Lenght '11

    'Verwendung als Objektvariable
    Dim Message As String2
    Set Message = String2.NewValue("Hello World")
    
    Debug.Print "Der String " & Message & " besteht aus " & Message.Lenght & " Zeichen."
    Message = Message.Append("!")
    
    If Message.EndsWith("!") Then
        
        Message = Message.Substring(0, 11)
        Debug.Print Message.Lenght '11
        
    End If
    
    Set Message = Nothing
    
End Sub
