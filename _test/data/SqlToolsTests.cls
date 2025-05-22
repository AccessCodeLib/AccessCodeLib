VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit test class for SqlTools
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/data/SqlToolsTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Public Sub Setup()
   SqlTools.SqlDateFormat = vbNullString
End Sub

Public Sub TearDown()
   SqlTools.SqlDateFormat = vbNullString
End Sub

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

'
' TextToSqlText
'

'AccUnit:Row("abcdefg", "'", False, "'abcdefg'")
'AccUnit:Row("abcdefg", "'", True, "abcdefg")
'AccUnit:Row("abc'defg", "'", False, "'abc''defg'")
'AccUnit:Row("abc""defg", "'", False, "'abc""defg'")
'AccUnit:Row(Null, "'", False, "Null")
Public Function TextToSqlText_DefaultValueIfNullArg_ReturnsSqlText( _
                     ByVal Value As Variant, _
                     ByVal Delimiter As String, _
                     ByVal WithoutLeftRightDelim As Boolean, _
                     ByVal Expected As String) As String

   Dim Actual As String
   Actual = SqlTools.TextToSqlText(Value, Delimiter, WithoutLeftRightDelim)
   Assert.That Actual, Iz.EqualTo(Expected)

End Function


'
' DateToSqlText
'

'AccUnit:Row(Null, "", "Null")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#yyyy-mm-dd\#", "#2000-01-31#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#mm\/dd\/yyyy\#", "#01/31/2000#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "'yyyymmdd'", "'20000131'")
Public Sub DateToSqlText_UseFormatArg_ReturnsSqlText(ByVal Value As Variant, _
                     ByVal FormatString As String, _
                     ByVal Expected As String)
   Dim Actual As String
   Actual = SqlTools.DateToSqlText(Value, FormatString)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(Null, "", "Null")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#yyyy-mm-dd\#", "#2000-01-31#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "\#mm\/dd\/yyyy\#", "#01/31/2000#")
'AccUnit:Row(System.Convert.ToDateTime("2000-01-31"), "'yyyymmdd'", "'20000131'")
Public Sub DateToSqlText_UsePublicFormatVariable_ReturnsSqlText(ByVal Value As Variant, _
                     ByVal FormatString As String, _
                     ByVal Expected As String)
   Dim Actual As String
   SqlTools.SqlDateFormat = FormatString
   Actual = SqlTools.DateToSqlText(Value)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub DateToSqlText_UseNotDateValue_RaiseError()
   Dim Actual As String
   Assert.Throws vbObjectError
   Actual = SqlTools.DateToSqlText("1.1.1.1")
End Sub

Public Sub DateToSqlText_FormatVariableNotDefined_RaiseError()
   Dim Actual As String
   Assert.Throws -2147221503, "DateToSqlText"
   SqlTools.SqlDateFormat = vbNullString
   Actual = SqlTools.DateToSqlText(Date)
End Sub

'
' NumberToSqlText
'

'AccUnit:Row(Null, "Null")
'AccUnit:Row(0, "0")
'AccUnit:Row(000, "0")
'AccUnit:Row(12345, "12345")
'AccUnit:Row(1.2345, "1.2345")
'AccUnit:Row(0.123, "0.123")
'AccUnit:Row(.123, "0.123")
'AccUnit:Row(1E2, "100")
'AccUnit:Row(1E-2, "0.01")
Public Sub NumberToSqlText_UseNumbers_ReturnsSqlText(ByVal Value As Variant, ByVal Expected As String)
   Dim Actual As String
   Actual = SqlTools.NumberToSqlText(Value)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub NumberToSqlText_UseText_ThrowsErr13()
   Dim Actual As String
   Assert.Throws 13
   Actual = SqlTools.NumberToSqlText("abc")
End Sub

'
' BooleanToSqlText
'

'AccUnit:Row(Null, "True", "Null")
'AccUnit:Row(0, "True", "0")
'AccUnit:Row(12345, "True", "True")
'AccUnit:Row(1.2345, "True", "True")
'AccUnit:Row(-1, "True", "True")
'AccUnit:Row(True, "True", "True")
'AccUnit:Row(False, "True", "0")
Public Sub BooleanToSqlText_UseValues_ReturnsSqlText(ByVal Value As Variant, ByVal TrueString As String, ByVal Expected As String)
   Dim Actual As String
   Actual = SqlTools.BooleanToSqlText(Value, TrueString)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub BooleanToSqlText_UseText_ThrowsErr13()
   Dim Actual As String
   Assert.Throws 13
   Actual = SqlTools.BooleanToSqlText("abc")
End Sub

Public Sub BooleanToSqlText_UseTrue_ThrowsErrNOCONFIG()
   Dim Actual As String
   Assert.Throws vbObjectError + 1 ' = SqlToolsErrorNumbers.ERRNR_NOCONFIG
   Actual = SqlTools.BooleanToSqlText(True)
End Sub

Public Sub BooleanToSqlText_UseTrueSetTrueStringInSqlTools_ReturnSqlText()
   Const Expected As String = "TrueStringParam"
   Dim Actual As String
   Actual = SqlTools.Clone("", Expected, "*").BooleanToSqlText(True)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'
' ConvertToNumeric
'

'AccUnit:Row(Null, Null)
'AccUnit:Row(0, 0)
'AccUnit:Row("000", 0)
'AccUnit:Row(12345, 12345)
'AccUnit:Row("12345", 12345)
'AccUnit:Row(1.2345, 1.2345)
'AccUnit:Row("1.2345", 1.2345)
'AccUnit:Row("1,2345", 1.2345)
'AccUnit:Row(0.123, 0.123)
'AccUnit:Row(.123, 0.123)
'AccUnit:Row("", Null)
'AccUnit:Row("1.234.5", 1234.5)
'AccUnit:Row("1.234,5", 1234.5)
Public Sub ConvertToNumeric_UseValues_ReturnsValue(ByVal Value As Variant, ByVal Expected As Variant)
   Dim Actual As Variant
   Actual = SqlTools.ConvertToNumeric(Value)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ConvertToNumeric_UseText_ThrowsErr13()
   Dim Actual As String
   Assert.Throws 13
   Actual = SqlTools.ConvertToNumeric("abc")
End Sub

'
' ConvertToSqlText
'

'AccUnit:Row(Null, 2, "Null")
'AccUnit:Row(123.4, 2, "123.4")
'AccUnit:Row("abc", 3, "'abc'")
Public Sub ConvertToSqlText_UseValues_ReturnsString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType, ByVal Expected As String)
   Dim Actual As Variant
   Actual = SqlTools.ConvertToSqlText(Value, FieldDataType)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(True, "True")
Public Sub ConvertToSqlText_UseBoolean_ReturnsString(ByVal Value As Variant, ByVal Expected As String)
   Dim Actual As Variant
   Actual = SqlTools.DAO.ConvertToSqlText(Value, SQL_Boolean)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

'AccUnit:Row(#12/24/2023#, "#2023-12-24 00:00:00#")
Public Sub ConvertToSqlText_UseDate_ReturnsString(ByVal Value As Variant, ByVal Expected As String)
   Dim Actual As Variant
   Actual = SqlTools.DAO.ConvertToSqlText(Value, SQL_Date)
   Assert.That Actual, Iz.EqualTo(Expected)
End Sub

Public Sub ConvertToSqlText_NotExistingSqlFieldDataType_ThrowsErr13()
   Dim Actual As String
   Assert.Throws vbObjectError
   Actual = SqlTools.ConvertToSqlText("abc", -1)
End Sub
