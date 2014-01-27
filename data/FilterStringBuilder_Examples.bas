Attribute VB_Name = "FilterStringBuilder_Examples"
'---------------------------------------------------------------------------------------
' Class Module: FilterStringBuilder_Examples
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Beispiel zur Verwendung der FilterStringBuilder-Klasse
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data\FilterStringBuilder_Examples.bas</file>
'  <use>data\FilterStringBuilder.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private Function NewFilterStringBuilder() As FilterStringBuilder
   With New FilterStringBuilder
      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")
      Set NewFilterStringBuilder = .Self
   End With
End Function


' Beispiele:
Private Sub EqualFilter()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, "abc", SQL_Equal
      .Add "NumericField", SQL_Numeric, 133.45, SQL_Equal
      .Add "DateField", SQL_Date, Date, SQL_Equal
      .Add "BoolField", SQL_Boolean, True, SQL_Equal

      Debug.Print .ToString

   End With

End Sub

Private Sub EqualOrGreaterFilter()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, "abc", SQL_Equal + SQL_GreaterThan
      .Add "NumericField", SQL_Numeric, 133.45, SQL_Equal + SQL_GreaterThan
      .Add "DateField", SQL_Date, Date, SQL_Equal + SQL_GreaterThan

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilter()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, "abc", SQL_Between, "xyz"
      .Add "NumericField", SQL_Numeric, 133.45, SQL_Between
      .Add "DateField", SQL_Date, DateSerial(Year(Date), 1, 1), SQL_Between, Date

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilterWithNullValues()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, Null, SQL_Between, "xyz"
      .Add "NumericField", SQL_Numeric, 133.45, SQL_Between, Null
      .Add "DateField", SQL_Date, Null, SQL_Between, Null

      Debug.Print .ToString

   End With

End Sub

Private Sub LikeFilter()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, "a*", SQL_Like

      Debug.Print .ToString

   End With

End Sub

Private Sub SqlInFilter()

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, Array("a", "c", "e"), SQL_In
      .Add "NumericField", SQL_Numeric, Array(1, 3.5, 4.25), SQL_In
      .Add "DateField", SQL_Date, Array(Date - 2, Date, Date + 12), SQL_In

      Debug.Print .ToString

   End With

End Sub

Private Sub SqlDateTimeBetween()

Dim StartDate As Date
Dim EndDate As Date

StartDate = #1/1/2014#
EndDate = #1/31/2014#

With NewFilterStringBuilder

   .Add "D1a", SQL_Date, EndDate, SQL_LessThan + SQL_Equal
   .Add "D1b", SQL_Date, EndDate, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix
   .Add "D2a", SQL_Date, EndDate, SQL_Equal
   .Add "D2b", SQL_Date, EndDate, SQL_Equal + SQL_Add_WildCardSuffix
   .Add "D3a", SQL_Date, StartDate, SQL_Between, EndDate
   .Add "D3b", SQL_Date, StartDate, SQL_Between + SQL_Add_WildCardSuffix, EndDate
   Debug.Print .ToString(" AND " & vbNewLine)

End With

End Sub
