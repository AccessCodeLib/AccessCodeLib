Attribute VB_Name = "String2_Example"
'---------------------------------------------------------------------------------------
' Package: _examples.text.String2_Example
'---------------------------------------------------------------------------------------
'
' Example for using the String2 class
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_examples/text/String2_Example.bas</file>
'  <use>text/String2.cls</use>
'  <test>_test/text/String2Tests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub String2_Example()

    'Verwendung als statische Klasse
    Debug.Print String2.NewValue("   Hello").Append(" ").Append("World  ").Trim().Length '11

    'Verwendung als Objektvariable
    Dim Message As String2
    Set Message = String2.NewValue("Hello World")
    
    Debug.Print "Der String " & Message & " besteht aus " & Message.Length & " Zeichen."
    Message = Message.Append("!")
    
    If Message.EndsWith("!") Then
        
        Message = Message.Substring(0, 11)
        Debug.Print Message.Length '11
        
    End If
    
    Set Message = Nothing
    
End Sub
