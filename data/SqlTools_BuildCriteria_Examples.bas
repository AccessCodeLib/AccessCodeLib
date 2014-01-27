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

' --------------------------
' Beispiele
' --------------------------


Private Sub EqualCriteria()

   With SqlTool

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Equal, "abc")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Equal, 133.45)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Equal, Date)
      Debug.Print .BuildCriteria("BoolField", SQL_Boolean, SQL_Equal, True)

   End With

End Sub

Private Sub EqualOrGreaterFilter()

   With SqlTool

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Equal + SQL_GreaterThan, "abc")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Equal + SQL_GreaterThan, 133.45)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Equal + SQL_GreaterThan, Date)

   End With

End Sub

Private Sub BetweenFilter()

   With SqlTool

      Debug.Print .BuildCriteria("TextField", SQL_Text, SQL_Between, "abc", "xyz")
      Debug.Print .BuildCriteria("NumericField", SQL_Numeric, SQL_Between, 133.45, 456)
      Debug.Print .BuildCriteria("DateField", SQL_Date, SQL_Between, DateSerial(Year(Date), 1, 1), Date)

   End With

End Sub

Private Sub SqlDateTimeBetween()

   Dim StartDate As Date
   Dim EndDate As Date
   
   StartDate = #1/1/2014#
   EndDate = #1/31/2014#
   
   With SqlTool
   
      Debug.Print .BuildCriteria("D1", SQL_Date, SQL_LessThan + SQL_Equal + SQL_Add_WildCardSuffix, EndDate)
      Debug.Print .BuildCriteria("D2", SQL_Date, SQL_Equal + SQL_Add_WildCardSuffix, EndDate)
      Debug.Print .BuildCriteria("D3", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, StartDate, EndDate)
   
   End With

End Sub


' --------------------------
' Hilfsprozeduren
' --------------------------
Private Property Get SqlTool() As SqlTools
   Static m_SqlTool As SqlTools
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")
   End If
   Set SqlTool = m_SqlTool
End Property
