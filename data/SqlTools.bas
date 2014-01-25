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
Public Const SQL_DEFAULT_BOOLTRUESTRING As String = "" ' => SqlBooleanTrueString wird verwendet.
                                                   '    Zum Deaktivieren Wert eintragen (z. B. "1")

Public Const SQL_DEFAULT_WILDCARD As String = "*"

Public SqlDateFormat As String
Public SqlBooleanTrueString As String
Private m_SqlWildCardString As String

Private Const ResultTextIfNull As String = "NULL"

Public Enum SqlRelationalOperators
   SQL_Not = 1
   SQL_Equal = 2
   SQL_LessThan = 4
   SQL_GreaterThan = 8
   SQL_Like = 256
   SQL_Between = 512
   SQL_In = 1024
   SQL_Add_WildCardSuffix = 2048
   SQL_Add_WildCardPrefix = 4096
End Enum

Public Enum SqlFieldDataType
   SQL_Boolean = 1
   SQL_Numeric = 2
   SQL_Text = 3
   SQL_Date = 4
End Enum


Public Property Get SqlWildCardString() As String
   If Len(m_SqlWildCardString) > 0 Then
      SqlWildCardString = m_SqlWildCardString
   Else
      SqlWildCardString = SQL_DEFAULT_WILDCARD
   End If
End Property

Public Property Let SqlWildCardString(ByVal NewValue As String)
   m_SqlWildCardString = NewValue
End Property


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

'---------------------------------------------------------------------------------------
' Function: BooleanToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Boolean für SQL-Text aufbereiten
' </summary>
' <param name="Value">Übergabewert</param>
' <returns>String</returns>
' <remarks>
' Durch Str-Funktion wird . statt , verwendet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function BooleanToSqlText(ByVal Value As Variant, _
                        Optional ByVal TrueString As String = SQL_DEFAULT_BOOLTRUESTRING) As String

   Dim Result As String

   If IsNull(Value) Then
      BooleanToSqlText = ResultTextIfNull
      Exit Function
   End If

   If Value = True Then
      If Len(TrueString) = 0 Then
         TrueString = SqlBooleanTrueString
         If Len(TrueString) = 0 Then
            Err.Raise SqlToolsErrorNumbers.ERRNR_NOCONFIG, "BooleanToSqlText", "boolean string for true is not defined"
         End If
      End If
      BooleanToSqlText = TrueString
   Else
      BooleanToSqlText = "0"
   End If
   
End Function


Public Function BuildCriteria(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                               ByVal RelationalOperator As SqlRelationalOperators, _
                               ByVal FilterValue As Variant, _
                      Optional ByVal FilterValue2 As Variant = Null, _
                      Optional ByVal IgnoreValue As Variant = Null) As String

   Dim FilterValueString As String
   Dim FilterValue2String As String
   Dim OperatorString As String
   Dim Criteria As String
   Dim FilterString As String
   
   If NullFilterOrEmptyFilter(FieldName, Nz(FilterValue, FilterValue2), IgnoreValue, FilterString) Then
      BuildCriteria = FilterString
      Exit Function
   End If
   
   If (RelationalOperator And SQL_In) = SQL_In Then
      If IsArray(FilterValue) Then
         Criteria = GetValueArrayString(FilterValue, FieldDataType, ",", IgnoreValue)
      ElseIf VarType(FilterValue) = vbString Then ' Value ist bereits die Auflistung als String
         Criteria = FilterValue
      Else
         Criteria = GetFilterValueString(FilterValue, FieldDataType)
      End If
      If Len(Criteria) > 0 Then
         BuildCriteria = FieldName & " In (" & Criteria & ")"
      End If
      Exit Function
   End If

   Dim itm As Variant
   Dim ItmCriteria As String
   If IsArray(FilterValue) Then 'Kritereien über Or verbinden
      With New StringCollection
         For Each itm In FilterValue
            ItmCriteria = BuildCriteria(FieldName, FieldDataType, RelationalOperator, itm, , IgnoreValue)
            If Len(ItmCriteria) > 0 Then
               .Add ItmCriteria
            End If
         Next
         BuildCriteria = .ToString(" Or ")
      End With
      Exit Function
   End If

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      If (RelationalOperator And SQL_Add_WildCardSuffix) = SQL_Add_WildCardSuffix Then
         FilterValue = FilterValue & SqlWildCardString
      End If
      If (RelationalOperator And SQL_Add_WildCardPrefix) = SQL_Add_WildCardPrefix Then
         FilterValue = SqlWildCardString & FilterValue
      End If
   End If

   FilterValueString = GetFilterValueString(FilterValue, FieldDataType)
   FilterValue2String = GetFilterValueString(FilterValue2, FieldDataType)
      

   If (RelationalOperator And SQL_Between) = SQL_Between Then
      If IsNull(FilterValue2) Or IsMissing(FilterValue2) Then
         RelationalOperator = SQL_GreaterThan + SQL_Equal
      ElseIf IsNull(FilterValue) Then
         RelationalOperator = SQL_LessThan + SQL_Equal
         FilterValueString = FilterValue2String
      Else
         BuildCriteria = FieldName & " Between " & FilterValueString & " And " & FilterValue2String
         Exit Function
      End If
   End If

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      BuildCriteria = FieldName & " like " & FilterValueString
      Exit Function
   End If
   

   If (RelationalOperator And SQL_LessThan) = SQL_LessThan Then
      OperatorString = OperatorString & "<"
   End If
   
   If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
      OperatorString = OperatorString & ">"
   End If

   If (RelationalOperator And SQL_Equal) = SQL_Equal Then
      OperatorString = OperatorString & "="
   End If

   Criteria = FieldName & " " & OperatorString & " " & FilterValueString
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      Criteria = "Not " & Criteria
   End If

   BuildCriteria = Criteria

End Function

Private Function NullFilterOrEmptyFilter(ByVal FieldName As String, ByVal Value As Variant, ByVal IgnoreValue As Variant, _
                                         ByRef NullFilterString As String) As Boolean
   
   If IsNull(Value) Then
      If Not IsNull(IgnoreValue) Then
         NullFilterString = FieldName & " Is Null"
      End If
      NullFilterOrEmptyFilter = True
   ElseIf IsArray(Value) Then
      Dim a() As Variant
      a = Value
      If (0 / 1) + (Not Not a) = 0 Then ' leerer Array
         NullFilterOrEmptyFilter = True
      End If
   ElseIf Value = IgnoreValue Then
      NullFilterOrEmptyFilter = True
   End If

End Function

Private Function GetValueArrayString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType, _
                                     ByVal Delimiter As String, ByVal IgnoreValue As Variant) As String
   
   Dim i As Long

   With New StringCollection
      For i = LBound(Value) To UBound(Value)
         If Value(i) = IgnoreValue Then
         ElseIf IsNull(Value(i)) And IsNull(IgnoreValue) Then
         Else
         .Add GetFilterValueString(Value(i), FieldDataType)
         End If
      Next
      GetValueArrayString = .ToString(Delimiter)
   End With

End Function

Private Function GetFilterValueString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType) As String

   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Numeric
         GetFilterValueString = SqlTools.NumberToSqlText(Value)
      Case SqlFieldDataType.SQL_Text
         GetFilterValueString = SqlTools.TextToSqlText(Value)
      Case SqlFieldDataType.SQL_Date
         GetFilterValueString = SqlTools.DateToSqlText(Value)
      Case SqlFieldDataType.SQL_Boolean
         GetFilterValueString = SqlTools.BooleanToSqlText(Value)
      Case Else
         Err.Raise vbObjectError, "FilterStringBuilder.GetFilterValueString", "SqlFieldDataType '" & FieldDataType & "' wird nicht unterstützt."

   End Select
End Function
