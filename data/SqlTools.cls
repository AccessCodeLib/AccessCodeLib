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
' <remarks>Attribute VB_PredeclaredId = True einstellen, damit SqlTools ohne Instanzierung erstellt werden kann</remarks>
'
' \warning Nicht vergessen: Parameter f�r Datumsformat, Boolean und WildCard f�r das DBMS eintellen
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
Private Const SQL_DEFAULT_DATEFORMAT As String = "" ' => SqlDateFormat-Eigenschaft verwenden.
                                                    '    Zum Deaktivieren Wert eintragen (z. B. "\#yyyy\-mm\-dd\#"),
                                                    '    dann wird dieser Wert als Standardeintrag verwendet.
Private Const SQL_DEFAULT_BOOLTRUESTRING As String = "" ' => SqlBooleanTrueString wird verwendet.
                                                        '    Zum Deaktivieren Wert eintragen (z. B. "True oder 1")

Private Const SQL_DEFAULT_WILDCARD As String = "*" ' * = Standardwert,
                                                   ' ben�tigte Abweichungen �ber SqlWildCardString einstellen

Private Const SqlAndConcatString As String = " And "
Private Const SqlOrConcatString As String = " Or "

Private m_SqlDateFormat As String
Private m_SqlBooleanTrueString As String
Private m_SqlWildCardString As String

Private Const ResultTextIfNull As String = "NULL"

Public Enum SqlRelationalOperators
   [_IgnoreAll] = &H80000000
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

Public Enum SqlLogicalOperator
   [_SQL_Default] = 0
   SQL_And = 1
   SQL_Or = 2
End Enum

' Konfiguration f�r den SQL-Dialekt
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
' Function: DateToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Werte in String f�r SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="Value">�bergabewert</param>
' <param name="FieldDataType">Datentyp des zu konvertierenden Wertes</param>
' <returns>String</returns>
'**/
'---------------------------------------------------------------------------------------
Public Function ConvertToSqlText(ByVal Value As Variant, _
                                 ByVal FieldDataType As SqlFieldDataType) As String
                     
   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Text
         ConvertToSqlText = TextToSqlText(Value)
      Case SqlFieldDataType.SQL_Numeric
         ConvertToSqlText = NumberToSqlText(Value)
      Case SqlFieldDataType.SQL_Date
         ConvertToSqlText = DateToSqlText(Value)
      Case SqlFieldDataType.SQL_Boolean
         ConvertToSqlText = BooleanToSqlText(Value)
      Case Else
         Err.Raise vbObjectError, "SqlTools.ConvertToSqlText", "FieldDataType '" & FieldDataType & "' not supported"
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Function: TextToSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Text f�r SQL-Anweisung aufbereiten.
' </summary>
' <param name="Value">�bergabewert</param>
' <param name="Delimiter">Begrenzungszeichen f�r Text-Werte. (In den meisten DBMS wird ' als Begrenzungszeichen verwendet.)</param>
' <param name="WithoutLeftRightDelim">Nur Begrenzungszeichnen innerhalb des Werte verdoppeln, Eingrenzung jedoch nicht setzen.</param>
' <returns>String</returns>
' <remarks>
' Beispiel: strSQL = "select ... from tabelle where Feld = " & TextToSqlText("ab'cd")
'        => strSQL = "select ... from tabelle where Feld = 'ab''cd'"
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
' Datumswert in String f�r SQL-Anweisung umwandeln, die per VBA zusammengesetzt wird.
' </summary>
' <param name="Value">�bergabewert</param>
' <param name="FormatString">Datumsformat (von DBMS abh�ngig!)</param>
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
' Zahl f�r SQL-Text aufbereiten
' </summary>
' <param name="Value">�bergabewert</param>
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
' Boolean f�r SQL-Text aufbereiten
' </summary>
' <param name="Value">�bergabewert</param>
' <returns>String</returns>
' <remarks>
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
' <param name="FilterValue2">Optionale 2. Filterwert (f�r Between)</param>
' <param name="IgnoreValue">Jener Wert, f�r den keine Filterbedingung erzeugt werden soll. (Array-�bergabe von Werten m�glich)</param>
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

   If (RelationalOperator And [_IgnoreAll]) = [_IgnoreAll] Then
      Exit Function
   End If

   If NullFilterOrEmptyFilter(FieldName, FieldDataType, Nz(FilterValue, FilterValue2), IgnoreValue, Criteria) Then
      BuildCriteria = Criteria
      Exit Function
   End If

   If TryBuildInCriteria(FieldName, FieldDataType, RelationalOperator, FilterValue, FilterValue2, IgnoreValue, Criteria) Then
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
               ' nichts �ndern => >= DataValue  / SQL_Add_WildCardSuffix ist nicht logisch
            Else ' ganzen Tag ber�cksichtigen FieldName >= DateValue and FieldName < DateAdd("d", 1, FilterValue))
               BuildCriteria = BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan + SQL_Equal, FilterValue) & _
                               SqlAndConcatString & _
                               BuildCriteria(FieldName, FieldDataType, SQL_LessThan, DateAdd("d", 1, CDate(CLng(FilterValue))))
               Exit Function
            End If
         Else
            If (RelationalOperator And SQL_Equal) = SQL_Equal Then
               RelationalOperator = RelationalOperator - SQL_Equal
            End If
            FilterValue = DateAdd("d", 1, CDate(CLng(FilterValue)))
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
      OperatorString = " Like "
      If (RelationalOperator And SQL_Not) = SQL_Not Then
         OperatorString = " Not" & OperatorString
      End If
      BuildCriteria = FieldName & OperatorString & FilterValueString
      Exit Function
   End If
   
   OperatorString = GetRelationalOperatorString(RelationalOperator)

   Criteria = FieldName & " " & OperatorString & " " & FilterValueString
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      Criteria = "Not " & Criteria
   End If

   BuildCriteria = Criteria

End Function

Private Function GetRelationalOperatorString(ByRef RelationalOperator As SqlRelationalOperators) As String

   Dim OperatorString As String
   Dim op As SqlRelationalOperators
   
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      
      op = RelationalOperator Xor SQL_Not
      
      If op = SqlRelationalOperators.SQL_Equal Then ' => "=" zu "<>" .. null ber�cksichtigen?
         RelationalOperator = SQL_LessThan + SQL_GreaterThan
      ElseIf op = SQL_GreaterThan + SQL_LessThan Then ' => "<>" zu "=" .. null ber�cksichtigen?
         RelationalOperator = SQL_Equal
      Else
         RelationalOperator = RelationalOperator Xor SQL_Not
         If (op And SQL_Equal) = SQL_Equal Then
            RelationalOperator = RelationalOperator Xor SQL_Equal
         Else
            RelationalOperator = RelationalOperator Or SQL_Equal
         End If
         If (op And SQL_LessThan) = SQL_LessThan Then
            RelationalOperator = RelationalOperator Xor SQL_LessThan
            RelationalOperator = RelationalOperator Or SQL_GreaterThan
         End If
         If (op And SQL_GreaterThan) = SQL_GreaterThan Then
            RelationalOperator = RelationalOperator Xor SQL_GreaterThan
            RelationalOperator = RelationalOperator Or SQL_LessThan
         End If
      End If
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

   'Kriterien �ber Or verbinden
   For Each itm In FilterValue
      ItmCriteria = BuildCriteria(FieldName, FieldDataType, RelationalOperator, itm, , IgnoreValue)
      If Len(ItmCriteria) > 0 Then
         Criteria = Criteria & SqlOrConcatString & ItmCriteria
      End If
   Next
   If Len(Criteria) > 0 Then
      Criteria = Mid(Criteria, Len(SqlOrConcatString) + 1) ' 1. Or wegschneiden
   End If

   TryBuildArrayCriteria = True

End Function

Private Function TryBuildInCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                    ByRef RelationalOperator As SqlRelationalOperators, _
                                    ByRef FilterValue As Variant, _
                                    ByRef FilterValue2 As Variant, _
                                    ByRef IgnoreValue As Variant, _
                                    ByRef Criteria As String) As Boolean

   Dim OperatorString As String

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

   OperatorString = " In "
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      OperatorString = " Not" & OperatorString
   End If

   If Len(Criteria) > 0 Then
      Criteria = FieldName & OperatorString & "(" & Criteria & ")"
   End If

   TryBuildInCriteria = True

End Function

Private Function TryBuildBetweenCriteria(ByRef FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                         ByRef RelationalOperator As SqlRelationalOperators, _
                                         ByRef FilterValue As Variant, _
                                         ByRef FilterValue2 As Variant, _
                                         ByRef IgnoreValue As Variant, _
                                         ByRef Criteria As String) As Boolean
   
   If (RelationalOperator And SQL_Between) = False Then
      TryBuildBetweenCriteria = False
      Exit Function
   End If
   
   If (RelationalOperator And SQL_Not) = SQL_Not Then 'Bedingung umdrehen
      Criteria = BuildCriteria(FieldName, FieldDataType, SQL_LessThan, FilterValue, , IgnoreValue) & _
                 SqlAndConcatString & _
                 BuildCriteria(FieldName, FieldDataType, SQL_GreaterThan, FilterValue2, , IgnoreValue)
                 ' & SqlOrConcatString & FieldName & " Is Null)"
      TryBuildBetweenCriteria = True
      Exit Function
   End If
   
   If IsNull(FilterValue2) Or IsMissing(FilterValue2) Or ValuesAreEqual(FieldDataType, FilterValue2, IgnoreValue) Then
      RelationalOperator = SQL_GreaterThan + SQL_Equal
   ElseIf IsNull(FilterValue) Or ValuesAreEqual(FieldDataType, FilterValue, IgnoreValue) Then
      RelationalOperator = SQL_LessThan + SQL_Equal
      FilterValue = FilterValue2
      FilterValue2 = GetCheckedIgnoreValue(IgnoreValue)
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

End Function

Private Function GetCheckedIgnoreValue(ByVal IgnoreValue As Variant) As Variant
   If IsArray(IgnoreValue) Then
      GetCheckedIgnoreValue = IgnoreValue(LBound(IgnoreValue))
   Else
      GetCheckedIgnoreValue = IgnoreValue
   End If
End Function

Private Function NullFilterOrEmptyFilter(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                                         ByVal Value As Variant, ByVal IgnoreValue As Variant, _
                                         ByRef NullFilterString As String) As Boolean
   
   If IsObject(IgnoreValue) Then
   If IgnoreValue Is Nothing Then
      If IsNull(Value) Then
         NullFilterString = FieldName & " Is Null"
         NullFilterOrEmptyFilter = True
      Else
         NullFilterOrEmptyFilter = False
      End If
      Exit Function
   End If
   End If
   
   If IsNull(Value) Then
      If Not ValuesAreEqual(FieldDataType, Value, IgnoreValue) Then
         NullFilterString = FieldName & " Is Null"
      End If
      NullFilterOrEmptyFilter = True
   ElseIf IsArray(Value) Then
      Dim a() As Variant
      a = Value
      If (0 / 1) + (Not Not a) = 0 Then ' leerer Array
         NullFilterOrEmptyFilter = True
      End If
   Else
      NullFilterOrEmptyFilter = ValuesAreEqual(FieldDataType, Value, IgnoreValue)
   End If

End Function

Private Function ValuesAreEqual(ByVal FieldDataType As SqlFieldDataType, ByVal Value As Variant, ByVal Value2 As Variant) As Boolean
   
   If IsArray(Value2) Then
      ValuesAreEqual = ArrayContains(FieldDataType, Value2, Value)
   ElseIf IsNull(Value) Then
      ValuesAreEqual = IsNull(Value2)
   ElseIf IsNull(Value2) Then
      ValuesAreEqual = False
   Else
      Select Case FieldDataType
         Case SqlFieldDataType.SQL_Text
            ValuesAreEqual = (VBA.StrComp(Value, Value2, vbTextCompare) = 0)
         Case SqlFieldDataType.SQL_Numeric
            ValuesAreEqual = (CDbl(Value) = CDbl(Value2))
         Case SqlFieldDataType.SQL_Date
            ValuesAreEqual = (CDate(Value) = CDate(Value2))
         Case SqlFieldDataType.SQL_Boolean
            ValuesAreEqual = (CBool(Value) = CBool(Value2))
         Case Else
            ValuesAreEqual = (Value = Value2)
      End Select
   End If

End Function

Private Function ArrayContains(ByVal FieldDataType As SqlFieldDataType, ByVal ArrayToCheck As Variant, ByVal SearchValue As Variant) As Boolean
   
   Dim i As Long

   If IsNull(SearchValue) Then
      ArrayContains = ArrayContainsNull(ArrayToCheck)
      Exit Function
   End If

   For i = LBound(ArrayToCheck) To UBound(ArrayToCheck)
      If ValuesAreEqual(FieldDataType, ArrayToCheck(i), SearchValue) Then
         ArrayContains = True
         Exit Function
      End If
   Next

   ArrayContains = False

End Function

Private Function ArrayContainsNull(ByVal ArrayToCheck As Variant) As Boolean
   
   Dim i As Long

   For i = LBound(ArrayToCheck) To UBound(ArrayToCheck)
      If IsNull(ArrayToCheck(i)) Then
         ArrayContainsNull = True
         Exit Function
      End If
   Next

   ArrayContainsNull = False

End Function

Private Function GetValueArrayString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType, _
                                     ByVal Delimiter As String, ByVal IgnoreValue As Variant) As String
   
   Dim i As Long
   Dim s As String

   For i = LBound(Value) To UBound(Value)
      If IsArray(IgnoreValue) Then
         If ArrayContains(FieldDataType, IgnoreValue, Value(i)) Then
         Else
            s = s & Delimiter & GetFilterValueString(Value(i), FieldDataType)
         End If
      Else
         If Value(i) = IgnoreValue Then
         ElseIf IsNull(Value(i)) And IsNull(IgnoreValue) Then
         Else
            s = s & Delimiter & GetFilterValueString(Value(i), FieldDataType)
         End If
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
         Err.Raise vbObjectError, "FilterStringBuilder.GetFilterValueString", "SqlFieldDataType '" & FieldDataType & "' wird nicht unterst�tzt."
   End Select

End Function