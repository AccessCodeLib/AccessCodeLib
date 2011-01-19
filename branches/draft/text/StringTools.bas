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
' <param name="ValueToTest">Zu prüfender Wert</param>
' <param name="IgnoreSpaces">Leerzeichen am Anfang u. Ende ignorieren</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function IsNullOrEmpty(ByVal ValueToTest As Variant, Optional ByVal IgnoreSpaces As Boolean = False) As Boolean
   
   Dim TempValue As String
   
   If IsNull(ValueToTest) Then
      IsNullOrEmpty = True
      Exit Function
   End If
   
   TempValue = CStr(ValueToTest)
   
   If IgnoreSpaces Then
      TempValue = Trim$(TempValue)
   End If
   
   IsNullOrEmpty = (Len(TempValue) = 0)
   
End Function

'---------------------------------------------------------------------------------------
' Function: FormatText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Fügt in den Platzhalter des Formattextes die übergebenen Parameter ein
' </summary>
' <param name="FormatString">Textformat mit Platzhalter ... Beispiel: "XYZ{0}, {1}"</param>
' <param name="Args">übergabeparameter in passender Reihenfolge</param>
' <returns>String</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FormatText(ByVal FormatString As String, ParamArray Args() As Variant) As String

   Dim Arg As Variant
   Dim temp As String
   Dim i As Long
   
   temp = FormatString
   For Each Arg In Args
      temp = Replace(temp, "{" & i & "}", CStr(Arg))
      i = i + 1
   Next
   
   FormatText = temp

End Function

'---------------------------------------------------------------------------------------
' Function: Format
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ersetzt die VBA-Formatfunktion
' Erweiterung: [h] bzw. [hh] für Stundenanzeige über 24
' </summary>
' <param name="Expression"></param>
' <param name="FormatString">Ein gültiger benannter oder benutzerdefinierter Formatausdruck inkl. Erweiterung für Stundenanzeige über 24 (Standard-Formatanweisungen siehe VBA.Format)</param>
' <param name="FirstDayOfWeek">Wird an VBA.Format weitergereicht</param>
' <param name="FirstWeekOfYear">Wird an VBA.Format weitergereicht</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Format(ByVal Expression As Variant, Optional ByVal FormatString As Variant, _
              Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
              Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String

   Dim Hours As Long
   
   If IsDate(Expression) Then
      If InStr(1, FormatString, "[h", vbTextCompare) > 0 Then
         Hours = Fix(Round(CDate(Expression) * 24, 1))
         If Hours < 24 Then
            FormatString = Replace(FormatString, "[hh]", "hh")
            FormatString = Replace(FormatString, "[h]", "h")
         Else
            FormatString = Replace(FormatString, "[hh]", CStr(Hours))
            FormatString = Replace(FormatString, "[h]", CStr(Hours))
         End If
      End If
   End If

   Format = VBA.Format$(Expression, FormatString, FirstDayOfWeek, FirstWeekOfYear)

End Function
