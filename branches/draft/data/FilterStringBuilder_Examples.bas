Attribute VB_Name = "FilterStringBuilder_Examples"
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
'  <file>data/FilterStringBuilder_Examples.bas</file>
'  <use>data/FilterStringBuilder.cls</use>
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

      .Add "TextField", SQL_Text, SQL_Equal, "abc"
      .Add "NumericField", SQL_Numeric, SQL_Equal, 133.45
      .Add "DateField", SQL_Date, SQL_Equal, Date
      .Add "BoolField", SQL_Boolean, SQL_Equal, True

      Debug.Print .ToString

   End With

End Sub

Private Sub EqualOrGreaterFilter()

Debug.Print "EqualOrGreaterFilter:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_Equal + SQL_GreaterThan, "abc"
      .Add "NumericField", SQL_Numeric, SQL_Equal + SQL_GreaterThan, 133.45
      .Add "DateField", SQL_Date, SQL_Equal + SQL_GreaterThan, Date

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilter()

Debug.Print "BetweenFilter:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_Between, "abc", "xyz"
      .Add "NumericField", SQL_Numeric, SQL_Between, 133.45, 456
      .Add "DateField", SQL_Date, SQL_Between, DateSerial(Year(Date), 1, 1), Date

      Debug.Print .ToString

   End With

End Sub

Private Sub BetweenFilterWithNullValues()

Debug.Print "BetweenFilterWithNullValues:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_Between, Null, "xyz"
      .Add "NumericField", SQL_Numeric, SQL_Between, 133.45, Null
      .Add "DateField", SQL_Date, SQL_Between, Null, Null

      Debug.Print .ToString

   End With

End Sub

Private Sub LikeFilter()

Debug.Print "LikeFilter:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_Like, "a*"

      Debug.Print .ToString

   End With

End Sub

Private Sub SqlInFilter()

Debug.Print "SqlInFilter:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_In, Array("a", "c", "e")
      .Add "NumericField", SQL_Numeric, SQL_In, Array(1, 3.5, 4.25)
      .Add "DateField", SQL_Date, SQL_In, Array(Date - 2, Date, Date + 12)

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

   .Add "D1a", SQL_Date, SQL_LessThan + SQL_Equal, EndDate
   .Add "D1b", SQL_Date, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, EndDate
   .Add "D2a", SQL_Date, SQL_Equal, EndDate
   .Add "D2b", SQL_Date, SQL_Equal + SQL_Add_WildCardSuffix, EndDate
   .Add "D3a", SQL_Date, SQL_Between, StartDate, EndDate
   .Add "D3b", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, StartDate, EndDate
   Debug.Print .ToString(" AND " & vbNewLine)

End With

End Sub



' --------------------------
' Hilfsprozeduren
' --------------------------
Private Function NewFilterStringBuilder() As FilterStringBuilder
   With New FilterStringBuilder
      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")
      Set NewFilterStringBuilder = .Self
   End With
End Function
