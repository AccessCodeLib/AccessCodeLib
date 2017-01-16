VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterStringBuilderTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' FilterStringBuilderTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für FilterStringBuilder
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/FilterStringBuilderTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/FilterStringBuilder.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub Add_StringNumericDateBooleanFilter_CheckDoaSyntax()
   
   Const Expected As String = "(S Like 'a*') And (N >= 5.1) And (D < #2015-02-01#) And (B = True)"
   
   Dim Actual As Variant
   
   With New FilterStringBuilder
      
      Set .SqlTool = SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
      
      .Add "S", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
      .Add "N", SQL_Numeric, SQL_GreaterThan + SQL_Equal, 5.1
      .Add "D", SQL_Date, SQL_LessThan, #2/1/2015#
      .Add "B", SQL_Boolean, SQL_Equal, True
      
      Actual = .ToString(SQL_And, , True)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub Add_StringNumericDateBooleanFilterWithoutBrackets_CheckDoaSyntax()
   
   Const Expected As String = "S Like 'a*' And N >= 5.1 And D < #2015-02-01# And B = True"
   
   Dim Actual As Variant
   
   With New FilterStringBuilder
      
      Set .SqlTool = SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
      
      .Add "S", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
      .Add "N", SQL_Numeric, SQL_GreaterThan + SQL_Equal, 5.1
      .Add "D", SQL_Date, SQL_LessThan, #2/1/2015#
      .Add "B", SQL_Boolean, SQL_Equal, True
      
      Actual = .ToString(SQL_And, , False)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddCriteria_AddConString_CheckSqlText()
   
   Const Expected As String = "(a > b)"
   
   Dim Actual As Variant
   
   With New FilterStringBuilder
   
      .AddCriteria ("a > b")
      
      Actual = .ToString(SQL_And, , True)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ConfigSqlFormat_DAO_CheckFormats()
   
   Const TestDateFormat As String = "\#YYYY-mm-DD\#"
   Const TestBooleanString As String = "T"
   Const TestWildCardString As String = "?"
   
   With New FilterStringBuilder
   
      .ConfigSqlFormat TestDateFormat, TestBooleanString, TestWildCardString
      
      Assert.That .SqlTool.SqlDateFormat, Iz.EqualTo(TestDateFormat)
      Assert.That .SqlTool.SqlBooleanTrueString, Iz.EqualTo(TestBooleanString)
      Assert.That .SqlTool.SqlWildCardString, Iz.EqualTo(TestWildCardString)

   End With
   
End Sub

'AccUnit:Row(True, "((F1 = 1)) And ((F2a Like 'a*') Or (F2b Like 'a*') Or (F2c Like 'a*')) And ((F3a = True) Or (F3b = True) Or (F3c = True))")
'AccUnit:Row(False, "(F1 = 1) And (F2a Like 'a*' Or F2b Like 'a*' Or F2c Like 'a*') And (F3a = True Or F3b = True Or F3c = True)")
Public Sub ConditionGroups_AddOrGroups_CheckSqlText(ByVal WithCondBracket As Boolean, ByVal Expected As String)

   'AccUnit-Bugfix:
   Expected = Replace(Expected, "=True", "= True")
   
   Dim Actual As String

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

      Actual = .ToString(, , WithCondBracket)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(True, "((F1 = 1)) And (fiXyz In (Select idXyz From TabXyz Where (F2a Like 'a*') Or (F2b Like 'b*')))")
'AccUnit:Row(False, "(F1 = 1) And (fiXyz In (Select idXyz From TabXyz Where F2a Like 'a*' Or F2b Like 'b*'))")
Public Sub AddSubSelectCriteria_WithSubSelectCond_CheckSqlText(ByVal WithCondBracket As Boolean, ByVal Expected As String)

   'Const Expected As String = "((F1 = 1)) And (fiXyz In (Select idXyz From TabXyz Where (F2a Like 'a*') Or (F2b Like 'b*')))"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("fiXyz", SQL_In, "Select idXyz From TabXyz", , , SQL_Or)
         .Add "F2a", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
         .Add "F2b", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "b"
      End With

      Actual = .ToString(, , WithCondBracket)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(True, "((F1 = 1)) And (idXyz In (Select fiXyz From TabXyzItems Where (fiTyp In (1,2,3)) Group By fiXyz Having Count(*) = 3))")
'AccUnit:Row(False, "(F1 = 1) And (idXyz In (Select fiXyz From TabXyzItems Where fiTyp In (1,2,3) Group By fiXyz Having Count(*) = 3))")
Public Sub AddSubSelectCriteria_UseWhereReplacementText_CheckSqlText(ByVal WithCondBracket As Boolean, ByVal Expected As String)

   'Const Expected As String = "((F1 = 1)) And (idXyz In (Select fiXyz From TabXyzItems Where (fiTyp In (1,2,3)) Group By fiXyz Having Count(*) = 3))"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("idXyz", SQL_In, "Select fiXyz From TabXyzItems [WhereStatement] Group By fiXyz Having Count(*) = 3", , True)
         .Add "fiTyp", SQL_Numeric, SQL_In, Array(1, 2, 3)
      End With

      Actual = .ToString(, , WithCondBracket)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddSubSelectCriteria_IgnoreEmptySubFilter_CheckSqlText()

   Const Expected As String = "(F1 = 1)"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("fiXyz", SQL_In, "Select idXyz From TabXyz", True, , SQL_Or)
         .Add "F2a", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Null
         .Add "F2b", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Null
      End With

      Actual = .ToString()

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(True, "((F1 = 1)) And (idXyz In (Select fiXyz From TabXyzItems Where (fiTyp = 2) Group By fiXyz Having Count(*) = 3))")
'AccUnit:Row(False, "(F1 = 1) And (idXyz In (Select fiXyz From TabXyzItems Where fiTyp = 2 Group By fiXyz Having Count(*) = 3))")
Public Sub AddSubSelectCriteria_WithSubFilterAndUseWhereReplacementText_CheckSqlText(ByVal WithCondBracket As Boolean, ByVal Expected As String)

   'Const Expected As String = "((F1 = 1)) And (idXyz In (Select fiXyz From TabXyzItems Where (fiTyp = 2) Group By fiXyz Having Count(*) = 3))"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("idXyz", SQL_In, "Select fiXyz From TabXyzItems [WhereStatement] Group By fiXyz Having Count(*) = 3", False, True)
         .Add "fiTyp", SQL_Numeric, SQL_Equal, 2
      End With

      Actual = .ToString(, , WithCondBracket)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddSubSelectCriteria_IgnoreEmptySubFilterAndUseWhereReplacementText_CheckSqlText()

   Const Expected As String = "(F1 = 1)"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("idXyz", SQL_In, "Select fiXyz From TabXyzItems [WhereStatement] Group By fiXyz Having Count(*) = 3", True, True)
         .Add "fiTyp", SQL_Numeric, SQL_In, Null
      End With

      Actual = .ToString()

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddSubSelectCriteria_WithEmptySubFilterAndUseWhereReplacementText_CheckSqlText()

   Const Expected As String = "(F1 = 1) And (idXyz In (Select fiXyz From TabXyzItems Group By fiXyz Having Count(*) = 3))"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddSubSelectCriteria("idXyz", SQL_In, "Select fiXyz From TabXyzItems [WhereStatement] Group By fiXyz Having Count(*) = 3", False, True)
         .Add "fiTyp", SQL_Numeric, SQL_In, Null
      End With

      Actual = .ToString()

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(True, "((F1 = 1)) And (Exists (Select * From TabXyz Where (idXyz = T.fiXyz) And (F2a Like 'a*') And (F2b Like 'b*')))")
'AccUnit:Row(False, "(F1 = 1) And (Exists (Select * From TabXyz Where idXyz = T.fiXyz And F2a Like 'a*' And F2b Like 'b*'))")
Public Sub AddExistsCriteria_WithSubSelectCond_CheckSqlText(ByVal WithCondBracket As Boolean, ByVal Expected As String)

   'Const Expected As String = "((F1 = 1)) And (Exists (Select * From TabXyz Where (idXyz = T.fiXyz) And (F2a Like 'a*') And (F2b Like 'b*')))"
   
   'AccUnit-Bugfix:
   Expected = Replace(Expected, "=T.fiXyz", "= T.fiXyz")
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "F1", SQL_Numeric, SQL_Equal, 1
      
      With .AddExistsCriteria("Select * From TabXyz")
         .AddCriteria "idXyz = T.fiXyz"
         .Add "F2a", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "a"
         .Add "F2b", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, "b"
      End With

      Actual = .ToString(, , WithCondBracket)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddCriteria_WithCommaSeparator_CheckSqlText()

   Const Expected As String = "@F1 = 1, @F2 = 2, @F3 = 3"
   
   Dim Actual As String

   With New FilterStringBuilder

      Set .SqlTool = SqlTools.NewInstance("\#yyyy-mm-dd\#", "True", "*")

      .Add "@F1", SQL_Numeric, SQL_Equal, 1
      .Add "@F2", SQL_Numeric, SQL_Equal, 2
      .Add "@F3", SQL_Numeric, SQL_Equal, 3
   
      Actual = .ToString(SQL_CommaSeparator)

   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
