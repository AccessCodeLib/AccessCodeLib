VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Attribute VB_PredeclaredId = True
'---------------------------------------------------------------------------------------
' Klasse: SqlTools
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
'  <file>data/SqlTools.cls</file>
'  <license>_codelib/license.bas</license>
'  <replace>data/SqlTools.bas</replace>
'  <test>_test/data/SqlToolsTests.cls</test>
'  <test>_test/data/SqlToolsBuildCriteriaTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Enum SqlToolsErrorNumbers
   ERRNR_NOCONFIG = vbObjectError + 1
End Enum

Private Const SQL_DEFAULT_TEXTDELIMITER As String = "'"
Private Const SQL_DEFAULT_DATEFORMAT As String = "" ' => SqlDateFormat wird verwendet.
                                                   '    Zum Deaktivieren Wert eintragen (z. B. "\#yyyy\-mm\-dd\#")
Private Const SQL_DEFAULT_BOOLTRUESTRING As String = "" ' => SqlBooleanTrueString wird verwendet.
                                                   '    Zum Deaktivieren Wert eintragen (z. B. "1")

Private Const SQL_DEFAULT_WILDCARD As String = "*" ' * = Standardwert,
                                                   ' benötigte Abweichungen über SqlWildCardString einstellen

Private Const SqlAndConcatString As String = " And "

Private m_SqlDateFormat As String
Private m_SqlBooleanTrueString As String
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

' Konfiguration für den SQL-Dialekt
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

Public Property Get SqlDateFormat() As String
   If Len(m_SqlDateFormat) > 0 Then
      SqlDateFormat = m_SqlDateFormat
   Else
      SqlDateFormat = SQL_DEFAULT_DATEFORMAT
   End If
End Property

Public Property Let SqlDateFormat(ByVal NewValue As String)
   m_SqlDateFormat = NewValue
End Property


Public Property Get SqlBooleanTrueString() As String
   If Len(m_SqlBooleanTrueString) > 0 Then
      SqlBooleanTrueString = m_SqlBooleanTrueString
   Else
      SqlBooleanTrueString = SQL_DEFAULT_BOOLTRUESTRING
   End If
End Property

Public Property Let SqlBooleanTrueString(ByVal NewValue As String)
   m_SqlBooleanTrueString = NewValue
End Property

'---------------------------------------------------------------------------------------
' Function: Clone
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Neue Instanz mit Grundeinstellungen der aktuellen Instanz erstellen.
' </summary>
' <returns>String</returns>
' <remarks>
' Beispiel: strSQL = "select ... from tabelle where Feld = " & TextToSqlText("ab'cd")
'           => strSQL = "select ... from tabelle where Feld = 'ab''cd'"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Clone(Optional ByVal NewSqlDateFormat As String = SQL_DEFAULT_DATEFORMAT, _
                      Optional ByVal NewSqlBooleanTrueString As String = SQL_DEFAULT_BOOLTRUESTRING, _
                      Optional ByVal NewSqlWildCardString As String = SQL_DEFAULT_WILDCARD) As SqlTools


   If Len(NewSqlDateFormat) = 0 Then NewSqlDateFormat = Me.SqlDateFormat
   If Len(NewSqlBooleanTrueString) = 0 Then NewSqlBooleanTrueString = Me.SqlBooleanTrueString
   If Len(NewSqlWildCardString) = 0 Then NewSqlWildCardString = Me.SqlWildCardString

   Set Clone = NewInstance(NewSqlDateFormat, NewSqlBooleanTrueString, NewSqlWildCardString)

End Function

Public Function NewInstance(ByVal NewSqlDateFormat As String, _
                            ByVal NewSqlBooleanTrueString As String, _
                            ByVal NewSqlWildCardString As String) As SqlTools
   
   Dim NewInst As SqlTools

   Set NewInst = New SqlTools
   With NewInst
      .SqlDateFormat = NewSqlDateFormat
      .SqlBooleanTrueString = NewSqlBooleanTrueString
      .SqlWildCardString = NewSqlWildCardString
   End With

   Set NewInstance = NewInst

End Function


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

'---------------------------------------------------------------------------------------
' Function: BuildCriteria
'---------------------------------------------------------------------------------------
'/**
' <summary>
' SQL-Kriterium erstellen
' </summary>
' <param name="FieldName">Feldname in der Datenquelle, die gefiltert werden soll</param>
' <param name="RelationalOperator">Vergleichsoperator (=, <=, usw.)</param>
' <param name="FilterValue">Filterwert (kann einzelner Wert oder auch Array mit Werten sein)</param>
' <param name="FilterValue2">Optionale 2. Filterwert (für Between)</param>
' <param name="IgnoreValue">Jener Wert, für den keine Filterbedingung erzeugt werden soll.</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function BuildCriteria(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                               ByVal RelationalOperator As SqlRelationalOperators, _
                               ByVal FilterValue As Variant, _
                      Optional ByVal FilterValue2 As Variant = Null, _
                      Optional ByVal IgnoreValue As Variant = Null) As String

   Dim FilterValueString As String
   Dim FilterValue2String As String
   Dim OperatorString As String
   Dim Criteria As String
   
   If NullFilterOrEmptyFilter(FieldName, Nz(FilterValue, FilterValue2), IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If TryBuildArrayCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, FilterValue2, IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If TryBuildBetweenCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, FilterValue2, IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      If SqlWildCardString <> "*" Then
         If InStr(1, FilterValue, "*") > 0 Then
            FilterValue = Replace(FilterValue, "[*]", "@@@|||STAR|||@@@")
            FilterValue = Replace(FilterValue, "*", SqlWildCardString)
            FilterValue = Replace(FilterValue, "@@@|||STAR|||@@@", "*")
         End If
      End If
   End If

   If (RelationalOperator And SQL_Add_WildCardSuffix) = SQL_Add_WildCardSuffix Then
      If (RelationalOperator And SQL_Like) = SQL_Like Then
         FilterValue = FilterValue & SqlWildCardString
      ElseIf (FieldDataType And SQL_Date) = SQL_Date Then
         If (RelationalOperator And SQL_LessThan) = 0 Then ' kein < daher:  >, >= oder nur =
            If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
               ' nichts ändern => >= DataValue  / SQL_Add_WildCardSuffix ist nicht logisch
            Else ' ganzen Tag berücksichtigen FieldName >= DateValue and FieldName < DateAdd("d", 1, FilterValue))
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue) & _
                               SqlAndConcatString & _
                               BuildCriteria(FieldName, FieldDataType, SQL_LessThan, DateAdd("d", 1, FilterValue))
               Exit Function
            End If
         Else
            If (RelationalOperator And SQL_Equal) = SQL_Equal Then
               RelationalOperator = RelationalOperator - SQL_Equal
            End If
            FilterValue = DateAdd("d", 1, FilterValue)
         End If
      End If
   End If

   If (RelationalOperator And SQL_Add_WildCardPrefix) = SQL_Add_WildCardPrefix Then
      If (RelationalOperator And SQL_Like) = SQL_Like Then
         FilterValue = SqlWildCardString & FilterValue
      End If
   End If

   FilterValueString = GetFilterValueString(FilterValue, FieldDataType)

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      BuildCriteria = FieldName & " like " & FilterValueString
      Exit Function
   End If
   
   OperatorString = GetRelationalOperatorString(RelationalOperator)

   Criteria = FieldName & " " & OperatorString & " " & FilterValueString
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      Criteria = "Not " & Criteria
   End If

   BuildCriteria = Criteria

End Function

Private Function GetRelationalOperatorString(ByVal RelationalOperator As SqlRelationalOperators) As String

   Dim OperatorString As String

   If (RelationalOperator And SQL_LessThan) = SQL_LessThan Then
      OperatorString = OperatorString & "<"
   End If
   
   If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
      OperatorString = OperatorString & ">"
   End If

   If (RelationalOperator And SQL_Equal) = SQL_Equal Then
      OperatorString = OperatorString & "="
   End If

   GetRelationalOperatorString = OperatorString

End Function

Private Function TryBuildArrayCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                       ByRef RelationalOperator As SqlRelationalOperators, _
                                       ByRef FilterValue As Variant, _
                                       ByRef FilterValue2 As Variant, _
                                       ByRef IgnoreValue As Variant, _
                                       ByRef Criteria As String) As Boolean
   Dim itm As Variant
   Dim ItmCriteria As String
   If Not IsArray(FilterValue) Then
      Exit Function
   End If

   'Kriterien über Or verbinden
   For Each itm In FilterValue
      ItmCriteria = BuildCriteria(FieldName, FieldDataType, RelationalOperator, itm, , IgnoreValue)
      If Len(ItmCriteria) > 0 Then
         Criteria = Criteria & " Or (" & ItmCriteria & ")"
      End If
   Next
   If Len(Criteria) > 0 Then
      Criteria = Mid(Criteria, 5) ' 1. Or wegschneiden
   End If

   TryBuildArrayCriteria = True

End Function

Private Function TryBuildInCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                    ByRef RelationalOperator As SqlRelationalOperators, _
                                    ByRef FilterValue As Variant, _
                                    ByRef FilterValue2 As Variant, _
                                    ByRef IgnoreValue As Variant, _
                                    ByRef Criteria As String) As Boolean

   If (RelationalOperator And SQL_In) = 0 Then
      Exit Function
   End If

   If IsArray(FilterValue) Then
      Criteria = GetValueArrayString(FilterValue, FieldDataType, ",", IgnoreValue)
   ElseIf VarType(FilterValue) = vbString Then ' Value ist bereits die Auflistung als String
      Criteria = FilterValue
   Else
      Criteria = GetFilterValueString(FilterValue, FieldDataType)
   End If

   If Len(Criteria) > 0 Then
      Criteria = FieldName & " In (" & Criteria & ")"
   End If

   TryBuildInCriteria = True

End Function

Private Function TryBuildBetweenCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                         ByRef RelationalOperator As SqlRelationalOperators, _
                                         ByRef FilterValue As Variant, _
                                         ByRef FilterValue2 As Variant, _
                                         ByRef IgnoreValue As Variant, _
                                         ByRef Criteria As String) As Boolean
   
   If (RelationalOperator And SQL_Between) = SQL_Between Then
      If IsNull(FilterValue2) Or IsMissing(FilterValue2) Or (FilterValue2 = IgnoreValue) Then
         RelationalOperator = SQL_GreaterThan + SQL_Equal
      ElseIf IsNull(FilterValue) Or (FilterValue = IgnoreValue) Then
         RelationalOperator = SQL_LessThan + SQL_Equal
         FilterValue = FilterValue2
         FilterValue2 = IgnoreValue
      ElseIf (FieldDataType And SQL_Date) = SQL_Date And (RelationalOperator And SQL_Add_WildCardSuffix) Then
         Criteria = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue) & _
                    SqlAndConcatString & _
                    BuildCriteria(FieldName, FieldDataType, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, FilterValue2)
         TryBuildBetweenCriteria = True
         Exit Function
      Else
         Criteria = FieldName & " Between " & GetFilterValueString(FilterValue, FieldDataType) & SqlAndConcatString & GetFilterValueString(FilterValue2, FieldDataType)
         TryBuildBetweenCriteria = True
         Exit Function
      End If
   End If

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
   Dim s As String

   For i = LBound(Value) To UBound(Value)
      If Value(i) = IgnoreValue Then
      ElseIf IsNull(Value(i)) And IsNull(IgnoreValue) Then
      Else
         s = s & Delimiter & GetFilterValueString(Value(i), FieldDataType)
      End If
   Next
   If Len(s) > 0 And Len(Delimiter) > 0 Then
      s = Mid(s, Len(Delimiter) + 1)
   End If
   GetValueArrayString = s

End Function

Private Function GetFilterValueString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType) As String

   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Numeric
         GetFilterValueString = NumberToSqlText(Value)
      Case SqlFieldDataType.SQL_Text
         GetFilterValueString = TextToSqlText(Value)
      Case SqlFieldDataType.SQL_Date
         GetFilterValueString = DateToSqlText(Value)
      Case SqlFieldDataType.SQL_Boolean
         GetFilterValueString = BooleanToSqlText(Value)
      Case Else
         Err.Raise vbObjectError, "FilterStringBuilder.GetFilterValueString", "SqlFieldDataType '" & FieldDataType & "' wird nicht unterstützt."

   End Select
End Function
