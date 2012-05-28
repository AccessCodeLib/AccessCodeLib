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
' Enum: TrimOption
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Verfügbare Trim-Optionen für die Round-Funktion
' </summary>
' <list type="table">
'   <item><term>TrimBoth (1)</term><description>Führende und nachgestellte Leerzeichen entfernen</description></item>
'   <item><term>TrimStart (2)</term><description>test</description>Führende Leerzeichen aus einer Zeichenfolgenvariablen entfernen</item>
'   <item><term>TrimEnd (3)</term><description>test</description>Nachgestellte Leerzeichen aus einer Zeichenfolgenvariablen entfernen</item>
' </list>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Enum TrimOption
    TrimBoth
    TrimStart
    TrimEnd
End Enum

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
         If Abs(Hours) < 24 Then
            FormatString = Replace(FormatString, "[hh]", "hh", , , vbTextCompare)
            FormatString = Replace(FormatString, "[h]", "h", , , vbTextCompare)
         Else
            FormatString = Replace(FormatString, "[hh]", "[h]", , , vbTextCompare)
            FormatString = Replace(FormatString, "[h]", Replace(CStr(Hours), "0", "\0"), , , vbTextCompare)
         End If
      End If
   End If

   Format = VBA.Format$(Expression, FormatString, FirstDayOfWeek, FirstWeekOfYear)

End Function

'---------------------------------------------------------------------------------------
' Function: PadLeft
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Linksbündiges Auffüllen eines Strings
' </summary>
' <param name="value">String der augefüllt werden soll</param>
' <param name="totalWidth">Gesamtlänge der resultierenen Zeichenfolge</param>
' <param name="padChar">Zeichen mit dem aufgefüllt werden soll</param>
' <returns>String</returns>
' <remarks>
' Wenn die Länge von value größer oder gleich totalWidth ist, wird das Resultat auf totalWidth Zeichen begrenzt
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function PadLeft(ByVal Value As String, ByVal totalWidth As Integer, Optional ByVal padChar As String = " ") As String
    PadLeft = VBA.Right$(VBA.String$(totalWidth, padChar) & Value, totalWidth)
End Function

'---------------------------------------------------------------------------------------
' Function: PadRight
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Rechtsbündiges Auffüllen eines Strings
' </summary>
' <param name="value">String der augefüllt werden soll</param>
' <param name="totalWidth">Gesamtlänge der resultierenen Zeichenfolge</param>
' <param name="padChar">Zeichen mit dem aufgefüllt werden soll</param>
' <returns>String</returns>
' <remarks>
' Wenn die Länge von value größer oder gleich totalWidth ist, wird das Resultat auf totalWidth Zeichen begrenzt
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function PadRight(ByVal Value As String, ByVal totalWidth As Integer, Optional ByVal padChar As String = " ") As String
    PadRight = VBA.Left$(Value & VBA.String$(totalWidth, padChar), totalWidth)
End Function

'---------------------------------------------------------------------------------------
' Function: Contains
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an ob searchValue in der Zeichenfolge checkValue vorkommt.
' </summary>
' <param name="checkValue">Zeichenfolge die durchsucht werden soll</param>
' <param name="searchValue">Zeichenfolge nach der gesucht werden soll</param>
' <returns>Boolean</returns>
' <remarks>
' Ergibt True, wenn searchValue in checkValue enthalten ist oder searchValue den Wert vbNullString hat
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Contains(ByVal checkValue As String, ByVal searchValue As String) As Boolean
    Contains = VBA.InStr(1, checkValue, searchValue, vbTextCompare) > 0
End Function

'---------------------------------------------------------------------------------------
' Function: EndsWith
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an ob die Zeichenfolge checkValue mit searchValue endet.
' </summary>
' <param name="checkValue">Zeichenfolge die durchsucht werden soll</param>
' <param name="searchValue">Zeichenfolge nach der gesucht werden soll</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function EndsWith(ByVal checkValue As String, ByVal searchValue As String) As Boolean
    EndsWith = VBA.Right$(checkValue, VBA.Len(searchValue)) = searchValue
End Function

'---------------------------------------------------------------------------------------
' Function: StartsWith
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt an ob die Zeichenfolge checkValue mit searchValue beginnt.
' </summary>
' <param name="checkValue">Zeichenfolge die durchsucht werden soll</param>
' <param name="searchvalue">Zeichenfolge nach der gesucht werden soll</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function StartsWith(ByVal checkValue As String, ByVal searchValue As String) As Boolean
    StartsWith = VBA.Left$(checkValue, VBA.Len(searchValue)) = searchValue
End Function

'---------------------------------------------------------------------------------------
' Function: Lenght
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die Anzahl von Zeichen in Value zurück
' </summary>
' <returns>Anzahl Zeichen von Value als Long</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Lenght(ByVal Value As String) As Long
    Lenght = VBA.Len(Value)
End Function

'---------------------------------------------------------------------------------------
' Function: Concat
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Fügt der Zeichenfolge ValueA die Zeihenfolge ValueB an.
' </summary>
' <param name="ValueA">Zeichenfolge</param>
' <param name="ValueB">Zeichenfolge</param>
' <returns>ValueB angefügt an ValueA als String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Concat(ByVal ValueA As String, ByVal ValueB As String) As String
    Concat = ValueA & ValueB
End Function

'---------------------------------------------------------------------------------------
' Function: Trim
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Entfernt führende und/oder nachfolgende Leerzeichen einer Zeichenfolge.
' Ersetzt die Funktion VBA.Trim().
' </summary>
' <param name="Value">Zeichenfolge</param>
' <param name="TrimType">Trim-Optionen</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Trim(ByVal Value As String, Optional ByVal TrimType As TrimOption = TrimOption.TrimBoth) As String
        
    Select Case TrimType
        Case TrimOption.TrimBoth
            Trim = VBA.Trim$(Value)
            Exit Function
        Case TrimOption.TrimStart
            Trim = VBA.LTrim$(Value)
            Exit Function
        Case TrimOption.TrimEnd
            Trim = VBA.RTrim(Value)
            Exit Function
        Case Else
            Trim = Value
            Exit Function
    End Select
    
End Function

'---------------------------------------------------------------------------------------
' Function: Substring
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt einen Teil der Zeichenfolge Value zurück, die an der Postiion startIndex beginnt
' und die Länge length hat.
' </summary>
' <param name="Value">Zeichenfolge</param>
' <param name="startIndex">Startposition in der Zeichenfolge</param>
' <param name="length">Anzahl Zeichen die Zurückgegeben werden sollen</param>
' <returns>String</returns>
' <remarks>
' startIndex ist Nullterminiert, analog zu String.Substring() in .NET
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Substring(ByVal Value As String, ByVal startIndex As Long, Optional ByVal length As Long = 0) As String
    If length = 0 Then length = StringTools.Lenght(Value) - startIndex
    Substring = VBA.Mid$(Value, startIndex + 1, length)
End Function
