VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime2Tests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'<codelib>
'  <file>_test\datetime\DateTime2Tests.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>

'AccUnit:TestClass

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

'AccUnit:Row("01.01.2013", 0, "01.01.2013")
'AccUnit:Row("01.01.2013", 1, "01.01.2014")
'AccUnit:Row("29.02.2012", 1, "28.02.2013")
'AccUnit:Row("29.02.2012", -1, "28.02.2011")
Public Sub NewValue_AddYear(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddYear(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2013", 0, "01.01.2013")
'AccUnit:Row("01.01.2013", 1, "01.04.2013")
'AccUnit:Row("01.01.2013", 4, "01.01.2014")
Public Sub NewValue_AddQuarter(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddQuarter(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2013", 1, "01.02.2013")
'AccUnit:Row("01.01.2013", 12, "01.01.2014")
'AccUnit:Row("01.01.2013", -12, "01.01.2012")
Public Sub NewValue_AddMonth(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddMonth(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("07.07.2012", 1, "08.07.2012")
'AccUnit:Row("01.01.2012", 31, "01.02.2012")
'AccUnit:Row("01.01.2012", 60, "01.03.2012")
'AccUnit:Row("01.01.2013", 59, "01.03.2013")
Public Sub NewValue_AddDayOfYear(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddDayOfYear(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("07.07.2012", 1, "14.07.2012")
'AccUnit:Row("07.07.2012 20:15:00", 1, "14.07.2012 20:15:00")
Public Sub NewValue_AddWeek(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddWeek(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:00:59", 1, "01:00:59")
'AccUnit:Row("01:00:01", -1, "00:00:01")
'AccUnit:Row("01.01.2000 00:00:00", -1, "31.12.1999 23:00:00")
'AccUnit:Row("31.12.1999 23:59:59", 1, "01.01.2000 00:59:59")
'AccUnit:Row("01.01.2000", -1, "31.12.1999 23:00:00")
Public Sub NewValue_AddHour(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddHour(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:00:59", 1, "00:01:59")
'AccUnit:Row("00:01:00", -1, "00:00:00")
'AccUnit:Row("01.01.2000 00:00:00", -1, "31.12.1999 23:59:00")
'AccUnit:Row("31.12.1999 23:59:59", 1, "01.01.2000 00:00:59")
'AccUnit:Row("01.01.2000", -1, "31.12.1999 23:59:00")
Public Sub NewValue_AddMinute(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddMinute(Number), Iz.EqualTo(Expected)
End Sub


'AccUnit:Row("01.01.2000 00:00:00", 1, "02.01.2000")
'AccUnit:Row("31.12.1999 23:59:59", 1, "01.01.2000 23:59:59")
'AccUnit:Row("01.01.2000", -1, "31.12.1999")
'AccUnit:Row("01.01.2000 00:00:00", -1, "31.12.1999")
Public Sub NewValue_AddDay(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddDay(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:00:59", 1, "00:01:00")
'AccUnit:Row("00:01:00", -1,"00:00:59")
'AccUnit:Row("01.01.2000 00:00:00", -1, "31.12.1999 23:59:59")
'AccUnit:Row("31.12.1999 23:59:59", 1, "01.01.2000")
'AccUnit:Row("01.01.2000", -1, "31.12.1999 23:59:59")
'AccUnit:Row("01.01.2000 00:00:00", -61, "31.12.1999 23:58:59")
Public Sub NewValue_AddSecond(ByVal vValue As Variant, ByVal Number As Integer, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).AddSecond(Number), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("31.12.2013 00:00:00", "31.12.2013 00:00:00")
'AccUnit:Row("29.02.2012 23:00:00", "29.02.2012 23:00:00")
'AccUnit:Row("28.02.2011 00:59:00", "28.02.2011 00:59:00")
'AccUnit:Row("31.12.2013 00:00:59", "31.12.2013 00:00:59")
'AccUnit:Row("31.12.2013", "31.12.2013")
'AccUnit:Row("29.02.2012", "29.02.2012")
'AccUnit:Row("28.02.2011", "28.02.2011")
'AccUnit:Row("31.12.2013", "31.12.2013")
'AccUnit:Row("00:00:00", "00:00:00")
'AccUnit:Row("23:00:00", "23:00:00")
'AccUnit:Row("00:59:00", "00:59:00")
'AccUnit:Row("00:00:59", "00:00:59")
Public Sub NewValue_ToString(ByVal vValue As Variant, ByVal Expected As String)
    Assert.That DateTime2.NewValue(vValue).ToString(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("31.12.2013 00:00:00", "31.12.2013 00:00:00")
'AccUnit:Row("29.02.2012 23:00:00", "29.02.2012 23:00:00")
'AccUnit:Row("28.02.2011 00:59:00", "28.02.2011 00:59:00")
'AccUnit:Row("31.12.2013 00:00:59", "31.12.2013 00:00:59")
'AccUnit:Row("31.12.2013", "31.12.2013")
'AccUnit:Row("29.02.2012", "29.02.2012")
'AccUnit:Row("28.02.2011", "28.02.2011")
'AccUnit:Row("31.12.2013", "31.12.2013")
'AccUnit:Row("00:00:00", "00:00:00")
'AccUnit:Row("23:00:00", "23:00:00")
'AccUnit:Row("00:59:00", "00:59:00")
'AccUnit:Row("00:00:59", "00:00:59")
Public Sub NewValue_ToDate(ByVal vValue As Variant, ByVal Expected As Date)
    Assert.That DateTime2.NewValue(vValue).ToDate(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2012", "31.01.2012")
'AccUnit:Row("15.02.2012", "29.02.2012")
'AccUnit:Row("15.02.2011", "28.02.2011")
'AccUnit:Row("31.12.2011", "31.12.2011")
Public Sub NewValue_LastDayOfMonth(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfMonth(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("15.03.2013", "31.03.2013")
'AccUnit:Row("15.04.2013", "30.06.2013")
'AccUnit:Row("15.07.2013", "30.09.2013")
'AccUnit:Row("15.11.2013", "31.12.2013")
Public Sub NewValue_LastDayOfQuarter(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfQuarter(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2011", "01.01.2011")
'AccUnit:Row("15.02.2012", "01.02.2012")
'AccUnit:Row("15.02.2011", "01.02.2011")
'AccUnit:Row("31.12.2011", "01.12.2011")
Public Sub NewValue_FirstDayOfMonth(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfMonth(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2010", "01.01.2010")
'AccUnit:Row("15.02.2012", "01.01.2012")
'AccUnit:Row("15.02.2011", "01.01.2011")
'AccUnit:Row("31.12.2013", "01.01.2013")
Public Sub NewValue_FirstDayOfYear(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("15.03.2013", "01.01.2013")
'AccUnit:Row("15.04.2013", "01.04.2013")
'AccUnit:Row("15.07.2013", "01.07.2013")
'AccUnit:Row("15.11.2013", "01.10.2013")
Public Sub NewValue_FirstDayOfQuarter(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfQuarter(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("20.02.2008", "18.02.2008")
'AccUnit:Row("31.08.2009", "31.08.2009")
'AccUnit:Row("26.08.2009", "24.08.2009")
'AccUnit:Row("03.10.2010", "27.09.2010")
'AccUnit:Row("15.10.2010", "11.10.2010")
'AccUnit:Row("08.11.2011", "07.11.2011")
'AccUnit:Row("05.07.2012", "02.07.2012")
'AccUnit:Row("05.02.2013", "04.02.2013")
'AccUnit:Row("20.02.2014", "17.02.2014")
'AccUnit:Row("20.02.2015", "16.02.2015")
'AccUnit:Row("01.03.2015", "23.02.2015")
'AccUnit:Row("20.02.2016", "15.02.2016")
'AccUnit:Row("29.02.2016", "29.02.2016")
Public Sub NewValue_FirstDayOfWeek(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).FirstDayOfWeek(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("08.07.2012", "08.07.2012")
'AccUnit:Row("20.02.2008", "24.02.2008")
'AccUnit:Row("31.08.2009", "06.09.2009")
'AccUnit:Row("26.08.2009", "30.08.2009")
'AccUnit:Row("03.10.2010", "03.10.2010")
'AccUnit:Row("15.10.2010", "17.10.2010")
'AccUnit:Row("08.11.2011", "13.11.2011")
'AccUnit:Row("05.07.2012", "08.07.2012")
'AccUnit:Row("05.02.2013", "10.02.2013")
'AccUnit:Row("20.02.2014", "23.02.2014")
'AccUnit:Row("20.02.2015", "22.02.2015")
'AccUnit:Row("01.03.2015", "01.03.2015")
'AccUnit:Row("20.02.2016", "21.02.2016")
'AccUnit:Row("29.02.2016", "06.03.2016")
Public Sub NewValue_LastDayOfWeek(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfWeek(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2010", "31.12.2010")
'AccUnit:Row("15.02.2012", "31.12.2012")
'AccUnit:Row("15.02.2011", "31.12.2011")
'AccUnit:Row("31.12.2013", "31.12.2013")
Public Sub NewValue_LastDayOfYear(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).LastDayOfYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("31.12.2013", "31.12.2013")
'AccUnit:Row("31.12.2013 00:00:00", "31.12.2013")
'AccUnit:Row("29.02.2012 23:00:00", "29.02.2012")
'AccUnit:Row("28.02.2011 00:59:00", "28.02.2011")
'AccUnit:Row("31.12.2013 00:00:59", "31.12.2013")
Public Sub NewValue_TrimTime(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).TrimTime(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("31.12.2013 00:00:00", "00:00:00")
'AccUnit:Row("31.12.2013 00:00:01", "00:00:01")
'AccUnit:Row("29.02.2012 23:00:00", "23:00:00")
'AccUnit:Row("28.02.2011 00:59:00", "00:59:00")
'AccUnit:Row("31.12.2013 00:00:59", "00:00:59")
Public Sub NewValue_GetTime(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).GetTime(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("31.12.2013 00:00:00", "00:00:00")
'AccUnit:Row("31.12.2013 00:00:01", "00:00:01")
'AccUnit:Row("29.02.2012 23:00:00", "23:00:00")
'AccUnit:Row("28.02.2011 00:59:00", "00:59:00")
'AccUnit:Row("31.12.2013 00:00:59", "00:00:59")
Public Sub NewValue_TrimDate(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).TrimDate(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("29.02.2012", "29.02.2012")
'AccUnit:Row("28.02.2011", "28.02.2011")
'AccUnit:Row("29.02.2012 23:00:00", "29.02.2012")
'AccUnit:Row("28.02.2011 00:59:00", "28.02.2011")
Public Sub NewValue_GetDate(ByVal vValue As Variant, ByVal Expected As Variant)
    Assert.That DateTime2.NewValue(vValue).GetDate(), Iz.EqualTo(Expected)
End Sub

Public Sub NewValue_Now()
    Assert.That DateTime2.NewValue("00:00:00").Now().ToDate(), Iz.EqualTo(Now())
End Sub

'AccUnit:Row("00:00:00", 1899)
'AccUnit:Row("28.02.2011 00:59:00", 2011)
'AccUnit:Row("31.12.9999 23:59:59", 9999)
Public Sub NewValue_GetYear(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("15.03.2013", 1)
'AccUnit:Row("15.04.2013", 2)
'AccUnit:Row("15.07.2013", 3)
'AccUnit:Row("15.11.2013", 4)
Public Sub NewValue_GetQuarter(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetQuarter(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("28.02.2011 00:59:00", 2)
'AccUnit:Row("31.12.9999 23:59:59", 12)
Public Sub NewValue_GetMonth(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetMonth(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("28.02.2011 00:59:00", 59)
'AccUnit:Row("31.12.2011 23:59:59", 365)
'AccUnit:Row("31.12.2012 23:59:59", 366)
Public Sub NewValue_GetDayOfYear(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetDayOfYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("28.02.2011 00:59:00", 28)
'AccUnit:Row("29.02.2012 00:59:00", 29)
'AccUnit:Row("31.12.9999 23:59:59", 31)
Public Sub NewValue_GetDay(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetDay(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("02.07.2012 23:59:59", 1).Name("Mo")
'AccUnit:Row("03.07.2012 23:59:59", 2).Name("Di")
'AccUnit:Row("04.07.2012 23:59:59", 3).Name("Mi")
'AccUnit:Row("05.07.2012 23:59:59", 4).Name("Do")
'AccUnit:Row("06.07.2012 23:59:59", 5).Name("Fr")
'AccUnit:Row("07.07.2012 23:59:59", 6).Name("Sa")
'AccUnit:Row("08.07.2012 23:59:59", 7).Name("So")
Public Sub NewValue_GetDayOfWeek_DefaultFirstDayOfWeekIsMonday(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetDayOfWeek(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("01.01.2012", 52)
'AccUnit:Row("06.02.2012", 6)
'AccUnit:Row("02.08.2012", 31)
'AccUnit:Row("31.12.2012", 1)
Public Sub NewValue_GetWeekOfYear_DefaultVbFirstWeekOfYearRuleIsVbFirstFourDays(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetWeekOfYear(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:59:00", 0)
'AccUnit:Row("28.02.2011 00:59:00", 0)
'AccUnit:Row("28.02.2011", 0)
'AccUnit:Row("23:59:59", 23)
'AccUnit:Row("31.12.9999 23:59:59", 23)
Public Sub NewValue_GetHour(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetHour(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:59:00", 59)
'AccUnit:Row("28.02.2011 00:00:00", 0)
'AccUnit:Row("28.02.2011", 0)
'AccUnit:Row("23:59:59", 59)
'AccUnit:Row("31.12.9999 23:59:59", 59)
Public Sub NewValue_GetMinute(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetMinute(), Iz.EqualTo(Expected)
End Sub

'AccUnit:Row("00:59:00", 0)
'AccUnit:Row("28.02.2011 00:00:00", 0)
'AccUnit:Row("28.02.2011", 0)
'AccUnit:Row("23:59:59", 59)
'AccUnit:Row("31.12.9999 23:59:59", 59)
Public Sub NewValue_GetSecond(ByVal vValue As Variant, ByVal Expected As Integer)
    Assert.That DateTime2.NewValue(vValue).GetSecond(), Iz.EqualTo(Expected)
End Sub
