VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlToolsBuildCriteriaTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' SqlToolsBuildCriteriaTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit test class for SqlTools.BuildCriteria
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/SqlToolsBuildCriteriaTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_SqlTools As SqlTools

Public Sub Setup()
   Set m_SqlTools = New SqlTools
End Sub

Public Sub Teardown()
   Set m_SqlTools = Nothing
End Sub

Private Property Get DaoTools() As SqlTools
   Set DaoTools = m_SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
End Property

Private Property Get TSqlTools() As SqlTools
   Set TSqlTools = m_SqlTools.Clone("'yyyymmdd'", "1", "%")
End Property

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------


Public Sub BuildCriteria_IgnoreAll_ReturnEmptyString()

   Dim Actual As String

   Actual = SqlTools.BuildCriteria("F", SQL_Text, [_IgnoreAll], 1, 2)
   Assert.That "" & Actual, Iz.Empty

End Sub

'AccUnit:Row(2, Null, Null, Null, "", "").Name("Empty")
'AccUnit:Row(2, "a", Null, Null, "F = 'a'", "F = 'a'").Name("Equal")
'AccUnit:Row(2+8, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, "a", Null, Null, "F <= 'a'", "F <= 'a'").Name("Equal+LessThan")
'AccUnit:Row(256, "a", Null, Null, "F Like 'a'", "F Like 'a'").Name("Like")
'AccUnit:Row(256, "a*", Null, Null, "F Like 'a*'", "F Like 'a%'").Name("Like+*")
'AccUnit:Row(256, "a[*]", Null, Null, "F Like 'a[*]'", "F Like 'a*'").Name("Like+[*]")
'AccUnit:Row(256+2048, "a", Null, Null, "F Like 'a*'", "F Like 'a%'").Name("Like+WildCardSuffix")
'AccUnit:Row(256+4096, "a", Null, Null, "F Like '*a'", "F Like '%a'").Name("Like+WildCardPrefix")
'AccUnit:Row(256+2048+4096, "a", Null, Null, "F Like '*a*'", "F Like '%a%'").Name("Like+WildCardSuffix+Prefix")
'AccUnit:Row(512, "a", "d", Null, "F Between 'a' And 'd'", "F Between 'a' And 'd'").Name("Between")
'AccUnit:Row(512, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Between+V2=Null")
'AccUnit:Row(512, Null, "d", Null, "F <= 'd'", "F <= 'd'").Name("Between+V1=Null")
'AccUnit:Row(1 + 256, "d", Null, Null, "F Not Like 'd'", "F Not Like 'd'").Name("not like")
'AccUnit:Row(1 + 512, "a", "d", Null, "F < 'a' And F > 'd'", "F < 'a' And F > 'd'").Name("Not Between")
'AccUnit:Row(1+2, "a", Null, Null, "F <> 'a'", "F <> 'a'").Name("Not Equal")
'AccUnit:Row(1+2+4, "a", Null, Null, "F > 'a'", "F > 'a'").Name("Not <=")
'AccUnit:Row(1+4, "a", Null, Null, "F >= 'a'", "F >= 'a'").Name("Not <")
'AccUnit:Row(1+2+8, "a", Null, Null, "F < 'a'", "F < 'a'").Name("Not >=")
'AccUnit:Row(67586, "a", Null, Null, "F = 'a*'", "F = 'a%'").Name("Equal + SQL_Add_WildCardSuffix + SQL_UseLikeBehavior")
Public Sub BuildCriteria_TextFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDao As String, ByVal ExpectedTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedDao)

   Actual = TSqlTools.BuildCriteria("F", SQL_Text, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedTSql)

End Sub

'AccUnit:Row(256+2048, "a", Null, Null, "F Like 'a%'").Name("Like+WildCardSuffix")
Public Sub BuildCriteria_TextFilter_DefaultWildCardString( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = SqlTools.BuildCriteria("F", SQL_Text, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(2, Null, Null, Null, "").Name("Empty")
'AccUnit:Row(2, 123, Null, Null, "F = 123").Name("Equal")
'AccUnit:Row(1+2, 123, Null, Null, "F <> 123").Name("Not Equal")
'AccUnit:Row(2, 1.23, Null, Null, "F = 1.23").Name("Equal")
'AccUnit:Row(2+8, 123, Null, Null, "F >= 123").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, 123, Null, Null, "F <= 123").Name("Equal+LessThan")
'AccUnit:Row(512, 123, 456, Null, "F Between 123 And 456").Name("Between")
'AccUnit:Row(512, 12.3, "45.*", Null, "F >= 12.3 And F < 46").Name("Between*")
'AccUnit:Row(2+2048, 12, Null, Null, "F >= 12 And F < 13").Name("Equal+*")
'AccUnit:Row(512, 123, Null, Null, "F >= 123").Name("Between+V2=Null")
'AccUnit:Row(512, Null, 456, Null, "F <= 456").Name("Between+V1=Null")
'AccUnit:Row(2, ">123", Null, Null, "F > 123").Name("Rel op in value: >")
'AccUnit:Row(2, "{Null}", "{Empty}", Null, "F Is Null").Name("Rel op in value: Is Null")
'AccUnit:Row(2, "1-4", Null, Null, "F Between 1 And 4").Name("Rel op in value: a-b")
'AccUnit:Row(2, "1-4*", Null, Null, "F >= 1 And F < 5").Name("Rel op in value: a-b*")
'AccUnit:Row(2, "*", Null, Null, "F Is Not Null").Name("* => Not Null")
'AccUnit:Row(2+16384, "{@<>5}", Null, Null, "F<>5").Name("manually expression")
Public Sub BuildCriteria_NumericFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDaoOrTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Numeric, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedDaoOrTSql)

   Actual = TSqlTools.BuildCriteria("F", SQL_Numeric, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedDaoOrTSql)

End Sub

'AccUnit:Row(2, "{@<>5}", Null, Null, "F<>5").Name("manually expression not allowed")
Public Sub BuildCriteria_NumericFilter_ThrowError( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDaoOrTSql As String)
                           
   Dim Actual As String
   
   Assert.Throws 13
   Actual = DaoTools.BuildCriteria("F", SQL_Numeric, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
                           
End Sub

'AccUnit:Row(2, Null, Null, Null, "", "").Name("Empty")
'AccUnit:Row(2, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F = #2014-05-16#", "F = '20140516'").Name("Equal")
'AccUnit:Row(2+8, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F >= #2014-05-16#", "F >= '20140516'").Name("Equal+GreaterThan")
'AccUnit:Row(2+4, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F <= #2014-05-16#", "F <= '20140516'").Name("Equal+LessThan")
'AccUnit:Row(2+4+2048, System.Convert.ToDateTime("2014-05-16"), Null, Null, "F < #2014-05-17#", "F < '20140517'").Name("Equal+LessThan+WildCardSuffix")
'AccUnit:Row(512, System.Convert.ToDateTime("2014-05-16"), System.Convert.ToDateTime("2014-07-18"), Null, "F Between #2014-05-16# And #2014-07-18#", "F Between '20140516' And '20140718'").Name("Between+LessThan+WildCardSuffix")
'AccUnit:Row(512+2048, System.Convert.ToDateTime("2014-05-16"), System.Convert.ToDateTime("2014-07-18"), Null, "F >= #2014-05-16# And F < #2014-07-19#", "F >= '20140516' And F < '20140719'").Name("Between+LessThan+WildCardSuffix")
Public Sub BuildCriteria_DateFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDao As String, ByVal ExpectedTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedDao)

   Actual = TSqlTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedTSql)

End Sub

'AccUnit:Row(2, "01.12.2023-05.12.2023", Null, Null, "F Between #2023-12-01# And #2023-12-05#", "F Between '20231201' And '20231205'").Name("1.- 5.12.")
'AccUnit:Row(2, "<=2023-12-10", Null, Null, "F <= #2023-12-10#", "F <= '20231210'").Name("<=2023-12-10")
'AccUnit:Row(2, "<=2023-12-10", Null, Null, "F <= #2023-12-10#", "F <= '20231210'").Name("<=2023-12-10")
Public Sub BuildCriteria_DateFilter_SpecialFeatures( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue As Variant, ByVal FilterValue2 As Variant, ByVal IgnoreValue As Variant, _
                           ByVal ExpectedDao As String, ByVal ExpectedTSql As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedDao)

   Actual = TSqlTools.BuildCriteria("F", SQL_Date, RelationalOperator, FilterValue, FilterValue2, IgnoreValue)
   Assert.That "" & Actual, Iz.EqualTo(ExpectedTSql)

End Sub

'AccUnit:Row(2, Null, Null, Null, "").Name("Empty")
'AccUnit:Row(2, "a", "b", "c", "F = 'a' Or F = 'b' Or F = 'c'").Name("Equal")
'AccUnit:Row(2+8, "a", "b", "c", "F >= 'a' Or F >= 'b' Or F >= 'c'").Name("Equal+GreaterThan")
'AccUnit:Row(256, "a*", "b*", "c*", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'").Name("Like+*")
'AccUnit:Row(256+2048, "a", "b", "c", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'").Name("Like+WildCardSuffix")
'AccUnit:Row(256+2048+4096, "a", "b", "c", "F Like '*a*' Or F Like '*b*' Or F Like '*c*'").Name("Like+WildCardSuffix+Prefix")
'AccUnit:Row(1024, "a", "b", "c", "F In ('a','b','c')").Name("In")
'AccUnit:Row(1+1024, "a", "b", "c", "F Not In ('a','b','c')").Name("Not In")
Public Sub BuildCriteria_TextValueArrayFilter( _
                           ByVal RelationalOperator As SqlRelationalOperators, _
                           ByVal FilterValue1 As Variant, ByVal FilterValue2 As Variant, ByVal FilterValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, RelationalOperator, Array(FilterValue1, FilterValue2, FilterValue3), , Null)
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(Null, "a", "b", Null, "")
'AccUnit:Row("a", "a", "b", Null, "")
'AccUnit:Row("x", "a", "b", Null, "F = 'x'")
Public Sub BuildCriteria_IgnoreValueArray( _
                           ByVal FilterValue1 As Variant, _
                           ByVal IgnoreValue1 As Variant, ByVal IgnoreValue2 As Variant, ByVal IgnoreValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_Equal, FilterValue1, , Array(IgnoreValue1, IgnoreValue2, IgnoreValue3))
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row(Null, "a", "b", "a", "b", Null, "")
'AccUnit:Row(Null, "a", "b", "a", "b", "c", "F Is Null")
'AccUnit:Row(Null, "a", "b", "x", "y", "z", "F Is Null Or F In ('a','b')")
Public Sub BuildCriteria_InAndIgnoreValueArray( _
                           ByVal FilterValue1 As Variant, ByVal FilterValue2 As Variant, ByVal FilterValue3 As Variant, _
                           ByVal IgnoreValue1 As Variant, ByVal IgnoreValue2 As Variant, ByVal IgnoreValue3 As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_In, Array(FilterValue1, FilterValue2, FilterValue3), , Array(IgnoreValue1, IgnoreValue2, IgnoreValue3))
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("a ;b; c", "F In ('a','b','c')")
'AccUnit:Row("a", "F In ('a')")
'AccUnit:Row(Null, "")
'AccUnit:Row("", "F In ('')")
Public Sub BuildCriteria_InAndSplitValueToArray( _
                           ByVal FilterValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_In + SQL_SplitValueToArray, FilterValue, , Null)
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("*a and *b AND *c", "F Like '*a*' And F Like '*b*' And F Like '*c*'")
'AccUnit:Row("*a + *b + *c", "F Like '*a*' And F Like '*b*' And F Like '*c*'")
'AccUnit:Row("a or b or c", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'")
'AccUnit:Row("a ;b; c", "F Like 'a*' Or F Like 'b*' Or F Like 'c*'")
'AccUnit:Row("a", "F Like 'a*'")
'AccUnit:Row(Null, "")
'AccUnit:Row("", "F Like '*'")
Public Sub BuildCriteria_LikeAndSplitValueToArray( _
                           ByVal FilterValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix + SQL_SplitValueToArray, FilterValue, , Null)
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("1,2 ;2; 3", "F = 1.2 Or F = 2 Or F = 3")
'AccUnit:Row("<2 ;>3", "F < 2 Or F > 3")
'AccUnit:Row("<2 or >3", "F < 2 Or F > 3")
'AccUnit:Row(">=2 and <=3", "F >= 2 And F <= 3")
Public Sub BuildCriteria_EqualNumericAndSplitValueToArray( _
                           ByVal FilterValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = DaoTools.BuildCriteria("F", SQL_Numeric, SQL_Equal + SQL_SplitValueToArray, FilterValue)
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("1;2;3", "F = 1 Or F = 2 Or F = 3")
Public Sub BuildCriteria_EqualNumericStringArray_VbaSplit( _
                           ByVal FilterValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = SqlTools.BuildCriteria("F", SQL_Numeric, SQL_Equal, Split(FilterValue, ";"))
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("{@=5}", "Not F=5")
Public Sub BuildCriteria_NotDirectCriteria( _
                           ByVal FilterValue As Variant, _
                           ByVal Expected As String)

   Dim Actual As String

   Actual = SqlTools.BuildCriteria("F", SQL_Numeric, SQL_Not + SQL_AllowSqlDirect, FilterValue)
   
   Assert.That "" & Actual, Iz.EqualTo(Expected)

End Sub
