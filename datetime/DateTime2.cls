VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DateTime2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: DateTime2
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Ein erweiterter DateTime-Datentyp
' \ingroup datetime
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>datetime/DateTime2.cls</file>
'  <test>_test\datetime\DateTime2Tests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_Value As String
Private m_FirstDayOfWeek As VbDayOfWeek
Private m_FirstWeekOfYear As VbFirstWeekOfYear

Private Sub Class_Initialize()
    m_Value = "00:00:00"
    m_FirstDayOfWeek = vbMonday
    m_FirstWeekOfYear = vbFirstFourDays
End Sub

Public Property Get NewValue(Optional ByVal Value As Variant) As DateTime2
    Dim newInstance As New DateTime2
        newInstance = Value
    Set NewValue = newInstance
    Set newInstance = Nothing
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_Value
End Property

Public Property Let Value(ByVal Value As Variant)
Attribute Value.VB_UserMemId = 0
    If Not VBA.IsNull(Value) And VBA.Information.IsDate(Value) Then
        m_Value = Value
    Else
         Err.Raise vbObjectError, "DateTime2", "Value is not a valid DateTime Value"
    End If
End Property

Public Property Get Add(ByVal Interval As String, ByVal Value As Integer) As DateTime2
    Set Add = DateTime2.NewValue(VBA.DateTime.DateAdd(Interval, Value, m_Value))
End Property

Public Property Get AddYear(ByVal Value As Integer) As DateTime2
    Set AddYear = Me.Add("yyyy", Value)
End Property

Public Property Get AddQuarter(ByVal Value As Integer) As DateTime2
    Set AddQuarter = Me.Add("q", Value)
End Property

Public Property Get AddMonth(ByVal Value As Integer) As DateTime2
    Set AddMonth = Me.Add("m", Value)
End Property

Public Property Get AddDayOfYear(ByVal Value As Integer) As DateTime2
    Set AddDayOfYear = Me.Add("y", Value)
End Property

Public Property Get AddDay(ByVal Value As Integer) As DateTime2
    Set AddDay = Me.Add("d", Value)
End Property

Public Property Get AddWeek(ByVal Value As Integer) As DateTime2
    Set AddWeek = Me.Add("ww", Value)
End Property

Public Property Get AddHour(ByVal Value As Integer) As DateTime2
    Set AddHour = Me.Add("h", Value)
End Property

Public Property Get AddMinute(ByVal Value As Integer) As DateTime2
    Set AddMinute = Me.Add("n", Value)
End Property

Public Property Get AddSecond(ByVal Value As Integer) As DateTime2
    Set AddSecond = Me.Add("s", Value)
End Property

Public Property Get GetYear() As Integer
    GetYear = VBA.DateTime.DatePart("yyyy", m_Value)
End Property

Public Property Get GetQuarter() As Integer
    GetQuarter = VBA.DateTime.DatePart("q", m_Value)
End Property

Public Property Get GetMonth() As Integer
    GetMonth = VBA.DateTime.DatePart("m", m_Value)
End Property

Public Property Get GetMonthName() As String
    GetMonthName = VBA.Strings.MonthName(Me.GetMonth(), False)
End Property

Public Property Get GetMonthNameShort() As String
    GetMonthNameShort = VBA.Strings.MonthName(Me.GetMonth(), True)
End Property

Public Property Get GetDayOfYear() As Integer
    GetDayOfYear = VBA.DateTime.DatePart("y", m_Value)
End Property

Public Property Get GetDay() As Integer
    GetDay = VBA.DateTime.DatePart("d", m_Value)
End Property

Public Property Get GetDayOfWeek() As Integer
    GetDayOfWeek = VBA.DateTime.DatePart("w", m_Value, m_FirstDayOfWeek, m_FirstWeekOfYear)
End Property

Public Property Get GetWeekDayName() As String
    GetWeekDayName = VBA.Strings.WeekdayName(Me.GetDayOfWeek, False, m_FirstDayOfWeek)
End Property

Public Property Get GetWeekDayNameShort() As String
    GetWeekDayNameShort = VBA.Strings.WeekdayName(Me.GetDayOfWeek, True, m_FirstDayOfWeek)
End Property

Public Property Get GetWeekOfYear() As Integer
    GetWeekOfYear = VBA.DateTime.DatePart("ww", m_Value, m_FirstDayOfWeek, m_FirstWeekOfYear)
End Property

Public Property Get GetHour() As Integer
    GetHour = VBA.DateTime.DatePart("h", m_Value)
End Property

Public Property Get GetMinute() As Integer
    GetMinute = VBA.DateTime.DatePart("n", m_Value)
End Property

Public Property Get GetSecond() As Integer
    GetSecond = VBA.DateTime.DatePart("s", m_Value)
End Property

Public Property Get ToString() As String
    ToString = VBA.Conversion.CStr(m_Value)
End Property

Public Property Get ToDate() As Date
    ToDate = VBA.Conversion.CDate(m_Value)
End Property

Public Property Get FirstDayOfYear() As DateTime2
    Set FirstDayOfYear = DateTime2.NewValue("01.01." & VBA.Strings.Mid(m_Value, 7))
End Property

Public Property Get FirstDayOfMonth() As DateTime2
    Set FirstDayOfMonth = DateTime2.NewValue("01." & VBA.Strings.Mid(m_Value, 4))
End Property

Public Property Get FirstDayOfQuarter() As DateTime2
    Set FirstDayOfQuarter = Me.FirstDayOfYear().AddQuarter(Me.GetQuarter() - 1)
End Property

Public Property Get FirstDayOfWeek() As DateTime2
    Set FirstDayOfWeek = Me.AddDay(1 - Me.GetDayOfWeek())
End Property

Public Property Get LastDayOfYear() As DateTime2
    Set LastDayOfYear = Me.FirstDayOfYear().AddYear(1).AddDay(-1)
End Property

Public Property Get LastDayOfMonth() As DateTime2
    Set LastDayOfMonth = Me.FirstDayOfMonth().AddMonth(1).AddDay(-1)
End Property

Public Property Get LastDayOfQuarter() As DateTime2
    Set LastDayOfQuarter = Me.FirstDayOfQuarter().AddQuarter(1).AddDay(-1)
End Property

Public Property Get LastDayOfWeek() As DateTime2
    Set LastDayOfWeek = Me.AddDay(7 - Me.GetDayOfWeek())
End Property

Public Property Get TrimTime() As DateTime2
    Set TrimTime = DateTime2.NewValue(VBA.Strings.Left(m_Value, 10))
End Property

Public Property Get GetTime() As DateTime2
    Set GetTime = Me.TrimDate()
End Property

Public Property Get TrimDate() As DateTime2
    Set TrimDate = DateTime2.NewValue(VBA.Strings.Right(m_Value, 8))
End Property

Public Property Get GetDate() As DateTime2
    Set GetDate = Me.TrimTime()
End Property

Public Property Get Now() As DateTime2
    Set Now = DateTime2.NewValue(VBA.DateTime.Now())
End Property
