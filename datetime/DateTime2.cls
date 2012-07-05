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
' \brief        Eine erweiterter DateTime-Datentyp
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

Private Sub Class_Initialize()
    m_Value = vbNullString
End Sub

Public Property Get NewValue(ByVal Value As Variant) As DateTime2
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
        m_Value = VBA.Conversion.CDate(Value)
    Else
         Err.Raise vbObjectError, "DateTime2", "Value is not a valid DateTime Value"
    End If
End Property

'/*
' * Operations
' */

Public Property Get Add(ByVal Interval As String, ByVal Value As Integer) As DateTime2
    Set Add = DateTime2.NewValue(VBA.DateTime.DateAdd(Interval, Value, m_Value))
End Property

Public Property Get AddYear(ByVal Value As Integer) As DateTime2
    Set AddYear = DateTime2.NewValue(m_Value).Add("yyyy", Value)
End Property

Public Property Get AddQuarter(ByVal Value As Integer) As DateTime2
    Set AddQuarter = DateTime2.NewValue(m_Value).Add("q", Value)
End Property

Public Property Get AddMonth(ByVal Value As Integer) As DateTime2
    Set AddMonth = DateTime2.NewValue(m_Value).Add("m", Value)
End Property

Public Property Get AddDayOfYear(ByVal Value As Integer) As DateTime2
    Set AddDayOfYear = DateTime2.NewValue(m_Value).Add("y", Value)
End Property

Public Property Get AddDay(ByVal Value As Integer) As DateTime2
    Set AddDay = DateTime2.NewValue(m_Value).Add("d", Value)
End Property

Public Property Get AddWeekDay(ByVal Value As Integer) As DateTime2
    Set AddWeekDay = DateTime2.NewValue(m_Value).Add("w", Value)
End Property

Public Property Get AddWeek(ByVal Value As Integer) As DateTime2
    Set AddWeek = DateTime2.NewValue(m_Value).Add("ww", Value)
End Property

Public Property Get AddHour(ByVal Value As Integer) As DateTime2
    Set AddHour = DateTime2.NewValue(m_Value).Add("h", Value)
End Property

Public Property Get AddMinute(ByVal Value As Integer) As DateTime2
    Set AddMinute = DateTime2.NewValue(m_Value).Add("n", Value)
End Property

Public Property Get AddSecond(ByVal Value As Integer) As DateTime2
    Set AddSecond = DateTime2.NewValue(m_Value).Add("s", Value)
End Property

'/*
' * Standards
' */

Public Property Get ToString() As String
    ToString = VBA.Conversion.CStr(m_Value)
End Property


'/*
' * Special Days
' */

Public Property Get FirstDayOfYear() As DateTime2
    Set FirstDayOfYear = DateTime2.NewValue("01.01." & VBA.Strings.Mid(m_Value, 7))
End Property

Public Property Get FirstDayOfMonth() As DateTime2
    Set FirstDayOfMonth = DateTime2.NewValue("01." & VBA.Strings.Mid(m_Value, 4))
End Property

Public Property Get LastDayOfYear() As DateTime2
    Set LastDayOfYear = DateTime2.NewValue(m_Value).FirstDayOfYear().AddYear(1).AddDay(-1)
End Property

Public Property Get LastDayOfMonth() As DateTime2
    Set LastDayOfMonth = DateTime2.NewValue(m_Value).FirstDayOfMonth().AddMonth(1).AddDay(-1)
End Property
