Attribute VB_Name = "SqlTools_BuildCriteria_Examples"
'---------------------------------------------------------------------------------------
' Beispiel-Modul: SqlTools_BuildCriteria_Examples
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Beispiel zur Verwendung der BuildCriteria-Methode der SqlTools-Klasse
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/SqlTools_BuildCriteria_Examples.bas</file>
'  <use>ata/SqlTools.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

'--------------------------
' Filtervarianten
'--------------------------
Private Sub Filtervarianten()

   Dim Sql As String

   With DaoSqlTools
      
'------------------------------------------------
' Equal & Co. ... =, >, <
'------------------------------------------------

      ' F = 'abc'
      Sql = .BuildCriteria("F", SQL_Text, SQL_Equal, "abc"): Debug.Print Sql

      ' F <= 123
      Sql = .BuildCriteria("F", SQL_Numeric, SQL_Equal + SQL_LessThan, 123): Debug.Print Sql

      ' F >= #2014-01-01#
      Sql = .BuildCriteria("F", SQL_Date, SQL_Equal + SQL_GreaterThan, #1/1/2014#): Debug.Print Sql
   
      ' Spezialfälle:
      ' <= + SQL_Add_WildCardSuffix
      ' ... zum Kennzeichnen des "ganzen Tages" obwohl nur Datum ohne Uhrzeit übergeben wird
      ' F <= #2013-12-31*#  =>  F < #2014-01-01
      Sql = .BuildCriteria("F", SQL_Date, SQL_Equal + SQL_LessThan + SQL_Add_WildCardSuffix, #12/31/2013#): Debug.Print Sql

'------------------------------------------------
' Like
'------------------------------------------------
      ' F Like 'abc*'
      Sql = .BuildCriteria("F", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "abc"): Debug.Print Sql

      ' F Like '*abc*'
      Sql = .BuildCriteria("F", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix + SQL_Add_WildCardPrefix, "abc"): Debug.Print Sql


'------------------------------------------------
' Between
'------------------------------------------------
      ' F Between 123 And 456
      Sql = .BuildCriteria("F", SQL_Numeric, SQL_Between, 123, 456): Debug.Print Sql

      ' F Between #2014-01-01# And #2014-01-31#
      Sql = .BuildCriteria("F", SQL_Date, SQL_Between, #1/1/2014#, #1/31/2014#): Debug.Print Sql

      ' Spezialfälle:
      ' SQL_Between + SQL_Add_WildCardSuffix zum Kennzeichnen des "ganzen Tages"
      ' F Between #2013-01-01# And #2013-12-31*# =>  F >= #2013-01-01# And F < #2014-01-01#
      Sql = .BuildCriteria("F", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, #1/1/2013#, #12/31/2013#): Debug.Print Sql


'------------------------------------------------
' Or ... mehrere Werte auf das gleiche Datenfeld
'------------------------------------------------
      ' F Like 'a*' Or F = 'c*' Or F = 'e*'
      Sql = .BuildCriteria("F", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Array("a", "c", "e")): Debug.Print Sql

      ' F = 123 Or F = 456 Or F = 789 ... Anm.: diesen Ausdruck könnte man eventuell auch als In(123,456,789) darstellen
      Sql = .BuildCriteria("F", SQL_Numeric, SQL_Equal, Array(123, 456, 789)): Debug.Print Sql

'------------------------------------------------
' In(...)
'------------------------------------------------
      ' F In ('a','c','e')
      Sql = .BuildCriteria("F", SQL_Text, SQL_In, Array("a", "c", "e")): Debug.Print Sql

      ' F In (123,456,789)
      Sql = .BuildCriteria("F", SQL_Numeric, SQL_In, Array(123, 456, 789)): Debug.Print Sql

   End With

End Sub


'--------------------------
' Weitere Beispiele
'--------------------------

Private Sub EqualCriteria()

   With DaoSqlTools

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Equal, "abc")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Equal, 133.45)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Equal, Date)
      Debug.Print .BuildCriteria("BoolField", SQL_Boolean, SQL_Equal, True)

   End With

End Sub

Private Sub EqualOrGreaterFilter()

   With DaoSqlTools

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Equal + SQL_GreaterThan, "abc")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Equal + SQL_GreaterThan, 133.45)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Equal + SQL_GreaterThan, Date)

   End With

End Sub

Private Sub BetweenFilter()

   With DaoSqlTools

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Between, "abc", "xyz")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Between, 133.45, 456)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Between, DateSerial(Year(Date), 1, 1), Date)

   End With

End Sub

Private Sub SqlDateTime()

   Dim StartDate As Date
   Dim EndDate As Date
   
   StartDate = #1/1/2014#
   EndDate = #1/31/2014#
   
   With DaoSqlTools
   
      Debug.Print .BuildCriteria("D1", SQL_Date, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, EndDate)
      Debug.Print .BuildCriteria("D2", SQL_Date, SQL_Equal + SQL_Add_WildCardSuffix, EndDate)
      Debug.Print .BuildCriteria("D3", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, StartDate, EndDate)
   
   End With

End Sub


' --------------------------
' Hilfsprozeduren
' --------------------------
Private Property Get DaoSqlTools() As SqlTools
   Static m_SqlTool As SqlTools
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")
   End If
   Set DaoSqlTools = m_SqlTool
End Property
