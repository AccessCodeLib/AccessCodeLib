Attribute VB_Name = "StringTools"
Attribute VB_Description = "SQL-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Modul: StringTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Pötzl
' <summary>
' Text-Hilfsfunktionen
' </summary>
' <remarks></remarks>
'
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/StringTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/text/StringToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Function: IsNullOrEmpty
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an, ob der übergebene Wert Null oder eine leere Zeichenfolge ist.
' </summary>
' <param name="vValue">Übergabewert</param>
' <param name="IgnoreSpaces">Leerzeichen am Anfang u. Ende ignorieren</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function IsNullOrEmpty(ByVal vValue As Variant, Optional ByVal IgnoreSpaces As Boolean = False) As Boolean
   If IgnoreSpaces Then
      vValue = Trim(vValue)
   End If
   If Len(vValue) > 0 Then
      IsNullOrEmpty = False
   Else
      IsNullOrEmpty = True
   End If
End Function
