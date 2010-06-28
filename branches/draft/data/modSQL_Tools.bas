Attribute VB_Name = "modSQL_Tools"
Attribute VB_Description = "SQL-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Modul: modSQL
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Pötzl
' <summary>
' SQL-Hilfsfunktionen
' </summary>
' <remarks></remarks>
'
' \warning Nicht vergessen: SQL_DEFAULT_TEXTDELIMITER und SQL_DEFAULT_DATEFORMAT
'          für das DBMS anpassen oder die Parameter entsprechend einstellen.
'
' \ingroup      data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/modSQL_Tools.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/defGlobal.bas</use>
'  <use>base/modErrorHandler.bas</use>
'  <test>_test/data/modSQL_Tools_FormatTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit


'---------------------------------------------------------------------------------------
' Function: GetSQLString_Text (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text für SQL-Anweisung aufbereiten.
' </summary>
' <param name="vValue">Übergabewert</param>
' <param name="sDelimiter">Begrenzungszeichen für Text-Werte. (In den meinsten DBMS wird ' als Begrenzungszeichen verwendet.)</param>
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
                         
On Error GoTo HandleErr

   If bWithoutLeftRightDelim Then
      GetSQLString_Text = Replace$(Nz(vValue, vbNullString), sDelimiter, sDelimiter & sDelimiter)
   ElseIf IsNull(vValue) Then
      GetSQLString_Text = "NULL"
   Else
      GetSQLString_Text = sDelimiter & Replace$(vValue, sDelimiter, sDelimiter & sDelimiter) & sDelimiter
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetSQLString_Text", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: GetSQLString_Date (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datumswert in String für SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="vValue">Übergabewert</param>
' <param name="sFormatString">Datumsformat (von DBMS abhängig!)</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function GetSQLString_Date(ByVal vValue As Variant, Optional sFormatString As String = SQL_DEFAULT_DATEFORMAT) As String

On Error GoTo HandleErr

   If IsNull(vValue) Then
      GetSQLString_Date = "NULL"
   Else
      GetSQLString_Date = Format$(vValue, sFormatString)
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetSQLString_Date", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: GetSQLString_Number (2009-07-25)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zahl für SQL-Text aufbereiten
' </summary>
' <param name="vValue">Übergabewert</param>
' <returns>String</returns>
' <remarks>
' Durch Str-Funktion wird . statt , verwendet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetSQLString_Number(ByVal vValue As Variant) As String

On Error GoTo HandleErr

   If IsNull(vValue) Then
      GetSQLString_Number = "NULL"
   Else
      GetSQLString_Number = Trim$(Str$(vValue))
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetSQLString_Number", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function
