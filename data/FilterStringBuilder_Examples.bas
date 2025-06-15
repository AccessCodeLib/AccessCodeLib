Attribute VB_Name = "FilterStringBuilder_Examples"
'---------------------------------------------------------------------------------------
' Example Module: FilterStringBuilder_Examples
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Example for using the FilterStringBuilder class
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
' Examples
' --------------------------

Private Sub EqualFilter()

Debug.Print "EqualFilter:"

   With NewFilterStringBuilder

      .Add "TextField", SQL_Text, SQL_Equal, "ab'c"
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
      .Add "TextField", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
      
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
   Debug.Print .ToString(SQL_And)

End With

End Sub

Private Sub ConditionGroups()

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .NewConditionGroup(SQL_Or)
         .Add "F2a", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
         .Add "F2b", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
         .Add "F2c", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
      End With
      With .NewConditionGroup(SQL_Or)
         .Add "F3a", SQL_Boolean, SQL_Equal, True
         .Add "F3b", SQL_Boolean, SQL_Equal, True
         .Add "F3c", SQL_Boolean, SQL_Equal, True
      End With

      Debug.Print .ToString(SQL_And)

   End With

End Sub

Private Sub SubSelectCriteria()

   With New FilterStringBuilder
   
      With .AddSubSelectCriteria("fiXyz", SQL_In + SQL_Not, "Select idXyz From Tabelle", True, SQL_Or)
         .Add "x", SQL_Numeric, SQL_Equal, 4
         .Add "y", SQL_Numeric, SQL_Equal, 5
      End With
      Debug.Print .ToString
      
   End With

End Sub

Private Sub SubSelectCriteriaMitConditionGroup()

   With New FilterStringBuilder
   
      .Add "a", SQL_Numeric, SQL_Equal, 123
   
      With .AddSubSelectCriteria("fiXyz", SQL_In + SQL_Not, "Select idXyz From Tabelle", False, SQL_And)
         .Add "b", SQL_Text, SQL_Equal, "xyz"
         With .NewConditionGroup(SQL_Or)
            .Add "x", SQL_Numeric, SQL_Equal, 4
            .Add "y", SQL_Numeric, SQL_Equal, 5
         End With
      End With
      Debug.Print .ToString
      
   End With

End Sub

Private Sub ExistsCriteria()

   With New FilterStringBuilder
   
      With .AddExistsCriteria("Select * From Tabelle")
         .AddCriteria "idXyz = T2.fiXyz"
         .Add "x", SQL_Numeric, SQL_Equal, 4
      End With
      Debug.Print .ToString
      
   End With

End Sub

Private Sub SplitValueToArray()

With New FilterStringBuilder

   .Add "X", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Array("a", "b", "c")
   .Add "Y", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Split("a;b;c", ";")
   .Add "Z", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix + SQL_SplitValueToArray, "a;b;c"
   
   Debug.Print .ToString
   
End With

End Sub

' --------------------------
' Helper procedures
' --------------------------
Private Function NewFilterStringBuilder() As FilterStringBuilder
   With New FilterStringBuilder
      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")
      Set NewFilterStringBuilder = .Self
   End With
End Function
