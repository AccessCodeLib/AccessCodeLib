Attribute VB_Name = "FilterStringBuilderNet_Examples"
'---------------------------------------------------------------------------------------
' Beispiel-Modul: FilterStringBuilder_Examples
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Beispiel zur Verwendung der FilterStringBuilder-Klasse
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/FilterStringBuilderNet_Examples.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

' --------------------------
' Beispiele
' --------------------------

Private Sub EqualFilter()

Debug.Print "EqualFilter:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_Equal, "ab'c"
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_Equal, 133.45
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_Equal, Date
      .Add "BoolField", FieldDataType_Boolean, RelationalOperators_Equal, True

      Debug.Print .ToString

   End With

End Sub

Private Sub EqualFilterTSql()

Debug.Print "EqualFilter (T-SQL):"

   With NewFilterStringBuilderTSql

      .Add "TextField", FieldDataType_Text, RelationalOperators_Equal, "ab'c"
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_Equal, 133.45
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_Equal, Date
      .Add "BoolField", FieldDataType_Boolean, RelationalOperators_Equal, True

      Debug.Print .ToString

   End With

End Sub

Private Sub EqualOrGreaterFilter()

Debug.Print "EqualOrGreaterFilter:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_Equal + RelationalOperators_GreaterThan, "abc"
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_Equal + RelationalOperators_GreaterThan, 133.45
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_Equal + RelationalOperators_GreaterThan, Date

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilter()

Debug.Print "BetweenFilter:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_Between, Array("abc", "xyz")
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_Between, Array(133.45, 456)
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_Between, Array(DateSerial(Year(Date), 1, 1), Date)

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilterWithNullValues()

Debug.Print "BetweenFilterWithNullValues:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_Between, Array(Null, "xyz")
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_Between, Array(133.45, Null)
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_Between, Array(Null, Null)

      Debug.Print .ToString

   End With

End Sub

Private Sub LikeFilter()

Debug.Print "LikeFilter:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_Like, "a*"
      .Add "TextField", FieldDataType_Text, RelationalOperators_Like + RelationalOperators_AddWildcardSuffix, "a"

      Debug.Print .ToString; RelationalOperators_AddWildcardSuffix

   End With

End Sub

Private Sub SqlInFilter()

Debug.Print "SqlInFilter:"

   With NewFilterStringBuilder

      .Add "TextField", FieldDataType_Text, RelationalOperators_In, Array("a", "c", "e")
      .Add "NumericField", FieldDataType_Numeric, RelationalOperators_In, Array(1, 3.5, 4.25)
      .Add "DateField", FieldDataType_DateTime, RelationalOperators_In, Array(Date - 2, Date, Date + 12)

      Debug.Print .ToString

   End With

End Sub

Private Sub SqlDateTimeBetween()

Debug.Print "SqlDateTimeBetween:"

Dim StartDate As Date
Dim EndDate As Date

StartDate = #1/1/2014#
EndDate = #1/31/2014#

With NewFilterStringBuilder

   .Add "D1a", FieldDataType_DateTime, RelationalOperators_LessThan + RelationalOperators_Equal, EndDate
   .Add "D1b", FieldDataType_DateTime, RelationalOperators_LessThan + RelationalOperators_Equal + RelationalOperators_AddWildcardSuffix, EndDate
   .Add "D2a", FieldDataType_DateTime, RelationalOperators_Equal, EndDate
   .Add "D2b", FieldDataType_DateTime, RelationalOperators_Equal + RelationalOperators_AddWildcardSuffix, EndDate
   .Add "D3a", FieldDataType_DateTime, RelationalOperators_Between, Array(StartDate, EndDate)
   .Add "D3b", FieldDataType_DateTime, RelationalOperators_Between + RelationalOperators_AddWildcardSuffix, Array(StartDate, EndDate)
   Debug.Print .ToString(LogicalOperator_And)

End With

End Sub

Private Sub SqlDateTimeBetweenTSql()

Debug.Print "SqlDateTimeBetween:"

Dim StartDate As Date
Dim EndDate As Date

StartDate = #1/1/2014#
EndDate = #1/31/2014#

With NewFilterStringBuilderTSql

   .Add "D1a", FieldDataType_DateTime, RelationalOperators_LessThan + RelationalOperators_Equal, EndDate
   .Add "D1b", FieldDataType_DateTime, RelationalOperators_LessThan + RelationalOperators_Equal + RelationalOperators_AddWildcardSuffix, EndDate
   .Add "D2a", FieldDataType_DateTime, RelationalOperators_Equal, EndDate
   .Add "D2b", FieldDataType_DateTime, RelationalOperators_Equal + RelationalOperators_AddWildcardSuffix, EndDate
   .Add "D3a", FieldDataType_DateTime, RelationalOperators_Between, Array(StartDate, EndDate)
   .Add "D3b", FieldDataType_DateTime, RelationalOperators_Between + RelationalOperators_AddWildcardSuffix, Array(StartDate, EndDate)
   Debug.Print .ToString(LogicalOperator_And)

End With

End Sub



' --------------------------
' Hilfsprozeduren
' --------------------------
Private Function NewFilterStringBuilder() As ACLibSqlTools.ConditionStringBuilder
   Set NewFilterStringBuilder = ACLib.NetSqlTools.ConditionStringBuilder(ACLib.NetSqlTools.SqlConverters.DaoSqlConverter)
End Function

Private Function NewFilterStringBuilderTSql() As ACLibSqlTools.ConditionStringBuilder
   Set NewFilterStringBuilderTSql = ACLib.NetSqlTools.ConditionStringBuilder(ACLib.NetSqlTools.SqlConverters.TsqlSqlConverter)
End Function
