Attribute VB_Name = "modSQL_Tools"
Attribute VB_Description = "SQL-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Modul: modSQL
'---------------------------------------------------------------------------------------
'/**
' \author       Josef P�tzl
' <summary>
' SQL-Hilfsfunktionen
' </summary>
' <remarks></remarks>
'
' \warning Nicht vergessen: SQL_DEFAULT_TEXTDELIMITER und SQL_DEFAULT_DATEFORMAT
'          f�r das DBMS anpassen oder die Parameter entsprechend einstellen.
'
' \ingroup      data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/modSQL_Tools.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/defGlobal.bas</use>
'  <test>_test/data/modSQL_Tools_FormatTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Enum SqlToolsErrorNumbers
   ERRNR_NOCONFIG = vbObjectError + 1
End Enum

Public Const SQL_DEFAULT_TEXTDELIMITER As String = "'"
Public Const SQL_DEFAULT_DATEFORMAT As String = "" ' => SQL_DATEFORMAT wird verwendet.
                                                   '    Zum Deaktivieren Wert eintragen (z. B. "\#yyyy\-mm\-dd\#")
Public SQL_DATEFORMAT As String

'---------------------------------------------------------------------------------------
' Function: GetSQLString_Text (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text f�r SQL-Anweisung aufbereiten.
' </summary>
' <param name="vValue">�bergabewert</param>
' <param name="sDelimiter">Begrenzungszeichen f�r Text-Werte. (In den meinsten DBMS wird ' als Begrenzungszeichen verwendet.)</param>
' <param name="bWithoutLeftRightDelim">Nur Begrenzungszeichnen innerhalb des Werte verdoppeln, Eingrenzung jedoch nicht setzen.</param>
' <returns>String</returns>
' <remarks>
' Beispiel: strSQL = "select ... from tabelle where Feld = " & GetSQLString_Text("ab'cd")
'           => strSQL = "select ... from tabelle where Feld = 'ab''cd'"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetSQLString_Text(ByVal vValue As Variant, Optional ByVal sDelimiter As String = SQL_DEFAULT_TEXTDELIMITER, _
                         Optional ByVal bWithoutLeftRightDelim As Boolean = False) As String
   If bWithoutLeftRightDelim Then
      GetSQLString_Text = Replace$(Nz(vValue, vbNullString), sDelimiter, sDelimiter & sDelimiter)
   ElseIf IsNull(vValue) Then
      GetSQLString_Text = "NULL"
   Else
      GetSQLString_Text = sDelimiter & Replace$(vValue, sDelimiter, sDelimiter & sDelimiter) & sDelimiter
   End If
End Function

'---------------------------------------------------------------------------------------
' Function: GetSQLString_Date (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datumswert in String f�r SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="vValue">�bergabewert</param>
' <param name="sFormatString">Datumsformat (von DBMS abh�ngig!)</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function GetSQLString_Date(ByVal vValue As Variant, Optional ByVal sFormatString As String = SQL_DEFAULT_DATEFORMAT) As String
   If IsNull(vValue) Then
      GetSQLString_Date = "NULL"
   Else
      If Len(sFormatString) = 0 Then
         sFormatString = SQL_DATEFORMAT
         If Len(sFormatString) = 0 Then
            Err.Raise SqlToolsErrorNumbers.ERRNR_NOCONFIG, "GetSQLString_Date", "Kein Datumsformat verf�gbar"
         End If
      End If
      GetSQLString_Date = Format$(vValue, sFormatString)
   End If
End Function

'---------------------------------------------------------------------------------------
' Function: GetSQLString_Number (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zahl f�r SQL-Text aufbereiten
' </summary>
' <param name="vValue">�bergabewert</param>
' <returns>String</returns>
' <remarks>
' Durch Str-Funktion wird . statt , verwendet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetSQLString_Number(ByVal vValue As Variant) As String
   If IsNull(vValue) Then
      GetSQLString_Number = "NULL"
   Else
      GetSQLString_Number = Trim$(Str$(vValue))
   End If
End Function
