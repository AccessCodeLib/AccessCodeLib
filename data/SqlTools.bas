Attribute VB_Name = "SqlTools"
'---------------------------------------------------------------------------------------
' Modul: SqlTools
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' SQL-Hilfsfunktionen
' </summary>
' <remarks></remarks>
'
' \warning Nicht vergessen: SQL_DEFAULT_TEXTDELIMITER und SQL_DEFAULT_DATEFORMAT
'          für das DBMS anpassen oder die Parameter entsprechend einstellen.
'
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/SqlTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/data/SqlToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Enum SqlToolsErrorNumbers
   ERRNR_NOCONFIG = vbObjectError + 1
End Enum

Public Const SQL_DEFAULT_TEXTDELIMITER As String = "'"
Public Const SQL_DEFAULT_DATEFORMAT As String = "" ' => SqlDateFormat wird verwendet.
                                                   '    Zum Deaktivieren Wert eintragen (z. B. "\#yyyy\-mm\-dd\#")
Public SqlDateFormat As String

Private Const ResultTextIfNull As String = "NULL"

'---------------------------------------------------------------------------------------
' Function: TextToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text für SQL-Anweisung aufbereiten.
' </summary>
' <param name="Value">Übergabewert</param>
' <param name="Delimiter">Begrenzungszeichen für Text-Werte. (In den meisten DBMS wird ' als Begrenzungszeichen verwendet.)</param>
' <param name="WithoutLeftRightDelim">Nur Begrenzungszeichnen innerhalb des Werte verdoppeln, Eingrenzung jedoch nicht setzen.</param>
' <param name="ValueIfNull">Ersatzstring bei NULL (Standard = "NULL")</param>
' <returns>String</returns>
' <remarks>
' Beispiel: strSQL = "select ... from tabelle where Feld = " & TextToSqlText("ab'cd")
'           => strSQL = "select ... from tabelle where Feld = 'ab''cd'"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function TextToSqlText(ByVal Value As Variant, _
                     Optional ByVal Delimiter As String = SQL_DEFAULT_TEXTDELIMITER, _
                     Optional ByVal WithoutLeftRightDelim As Boolean = False) As String
   
   Dim Result As String
   
   If IsNull(Value) Then
      TextToSqlText = ResultTextIfNull
      Exit Function
   End If
   
   Result = Replace$(Value, Delimiter, Delimiter & Delimiter)
   If Not WithoutLeftRightDelim Then
      Result = Delimiter & Result & Delimiter
   End If
   
   TextToSqlText = Result

End Function

'---------------------------------------------------------------------------------------
' Function: DateToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datumswert in String für SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="vValue">Übergabewert</param>
' <param name="sFormatString">Datumsformat (von DBMS abhängig!)</param>
' <param name="ValueIfNull">Ersatzstring bei NULL (Standard = "NULL")</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function DateToSqlText(ByVal Value As Variant, _
                     Optional ByVal FormatString As String = SQL_DEFAULT_DATEFORMAT) As String

   If IsNull(Value) Then
      DateToSqlText = ResultTextIfNull
      Exit Function
   End If

   If Len(FormatString) = 0 Then
      FormatString = SqlDateFormat
      If Len(FormatString) = 0 Then
         Err.Raise SqlToolsErrorNumbers.ERRNR_NOCONFIG, "DateToSqlText", "date format is not defined"
      End If
   End If
   
   DateToSqlText = Format$(Value, FormatString)

End Function

'---------------------------------------------------------------------------------------
' Function: NumberToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zahl für SQL-Text aufbereiten
' </summary>
' <param name="Value">Übergabewert</param>
' <param name="ValueIfNull">Ersatzstring bei NULL (Standard = "NULL")</param>
' <returns>String</returns>
' <remarks>
' Durch Str-Funktion wird . statt , verwendet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function NumberToSqlText(ByVal Value As Variant) As String

   Dim Result As String

   If IsNull(Value) Then
      NumberToSqlText = ResultTextIfNull
      Exit Function
   End If
   
   Result = Trim$(Str$(Value))
   If Left(Result, 1) = "." Then
      Result = "0" & Result
   End If
   
   NumberToSqlText = Result
   
End Function
