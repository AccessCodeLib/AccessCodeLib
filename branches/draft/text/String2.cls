VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "String2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: String2
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Eine erweiterter String-Datentyp
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/String2.cls</file>
'  <use>text/StringTools.bas</use>
'  <use>text/String2Factory.bas</use>
'  <test>_test\text\StringTests.cls</test>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_vvalue As String

Private Sub Class_Initialize()
    m_vvalue = vbNullString
End Sub

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_vvalue
End Property

Public Property Let Value(ByVal vValue As Variant)
Attribute Value.VB_UserMemId = 0
    If Not VBA.IsNull(vValue) Then: m_vvalue = VBA.Conversion.CStr(vValue)
End Property

Public Property Get IsNullOrEmpty(Optional ByVal IgnoreSpaces As Boolean = False) As Boolean
    IsNullOrEmpty = StringTools.IsNullOrEmpty(m_vvalue, IgnoreSpaces)
End Property

Public Property Get ToString(Optional ByVal FormatString As Variant = vbNullString, _
                         Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
                         Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
                         
        ToString = StringTools.Format(m_vvalue, FormatString, FirstDayOfWeek, FirstWeekOfYear)
End Property

Public Property Get ToString2() As String2
    Set ToString2 = String2Factory.Value(m_vvalue)
End Property

Public Property Get PadLeft(ByVal totalWidth As Integer, Optional ByVal padChar As String = " ") As String2
    Set PadLeft = String2Factory.Value(StringTools.PadLeft(m_vvalue, totalWidth, padChar))
End Property

Public Property Get PadRight(ByVal totalWidth As Integer, Optional ByVal padChar As String = " ") As String2
    Set PadRight = String2Factory.Value(StringTools.PadRight(m_vvalue, totalWidth, padChar))
End Property

Public Property Get Contains(ByVal searchValue As String) As Boolean
    Contains = StringTools.Contains(m_vvalue, searchValue)
End Property

Public Property Get EndsWith(ByVal searchValue As String) As Boolean
    EndsWith = StringTools.EndsWith(m_vvalue, searchValue)
End Property

Public Property Get StartsWith(ByVal searchValue As String) As Boolean
    StartsWith = StringTools.StartsWith(m_vvalue, searchValue)
End Property

Public Property Get Lenght() As Decimal2
    Set Lenght = Decimal2Factory.Value(StringTools.Lenght(m_vvalue))
End Property

Public Property Get Append(ByVal Value As Variant) As String2
    Set Append = String2Factory.Value(StringTools.Concat(m_vvalue, Value))
End Property

Public Property Get Trim(Optional ByVal TrimType As TrimOption = TrimOption.TrimBoth) As String2
    Set Trim = String2Factory.Value(StringTools.Trim(m_vvalue, TrimType))
End Property

Public Property Get Substring(ByVal startIndex As Long, Optional ByVal length As Long = 0) As String2
    Set Substring = String2Factory.Value(StringTools.Substring(m_vvalue, startIndex, length))
End Property
