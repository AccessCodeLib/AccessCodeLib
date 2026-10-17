VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Decimal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: math.Decimal2
'---------------------------------------------------------------------------------------
'
' A replacement class for the missing Decimal data type
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/Decimal2.cls</file>
'  <use>math/MathTools.bas</use>
'  <use>text/StringTools.cls</use>
'  <use>text/String2.cls</use>
'  <test>_test/math/Decimal2Tests.cls</test>
'  <license>_codelib/license.bas</license>
'  <description>Decimal number type as replacement for the missing VBA Decimal data type: value, rounding, formatted output</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private m_vvalue As Variant

Private Sub Class_Initialize()
    m_vvalue = VBA.Conversion.CDec(0)
End Sub

Public Property Get NewValue(ByVal vValue As Variant) As Decimal2
    Dim newInstance As New Decimal2
        newInstance = vValue
    Set NewValue = newInstance
    Set newInstance = Nothing
End Property

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_vvalue
End Property

Public Property Let Value(ByVal vValue As Variant)
Attribute Value.VB_UserMemId = 0
    m_vvalue = VBA.Conversion.CDec(vValue)
End Property

'---------------------------------------------------------------------------------------
' Property: Round
'---------------------------------------------------------------------------------------
'
' Returns the current value rounded.
'
' Parameters:
'     numDigitsAfterDecimal - Number of digits to the right of the decimal point to round to
'     midpointRoundingType  - Rounding method
'
' Returns:
'     Decimal2
'
'---------------------------------------------------------------------------------------
Public Property Get Round(Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Decimal2
        
        Set Round = Decimal2.NewValue(MathTools.Round(m_vvalue, numDigitsAfterDecimal, midpointRoundingType))
End Property

'---------------------------------------------------------------------------------------
' Property: ToString
'---------------------------------------------------------------------------------------
'
' Returns the current value as String.
'
' Parameters:
'     FormatString    - A valid named or user-defined format expression incl. extension for hours above 24 (for standard format expressions see VBA.Format)
'     FirstDayOfWeek  - Wird an VBA.Format weitergereicht
'     FirstWeekOfYear - Wird an VBA.Format weitergereicht
'
' Returns:
'     String
'
' Remarks:
'     Optionally, the available parameters of the Format function can be passed.
'
'---------------------------------------------------------------------------------------
Public Property Get ToString(Optional ByVal FormatString As Variant = vbNullString, _
                         Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
                         Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
        ToString = StringTools.Format(m_vvalue, FormatString, FirstDayOfWeek, FirstWeekOfYear)
End Property

'---------------------------------------------------------------------------------------
' Property: ToString2
'---------------------------------------------------------------------------------------
'
' Returns the current value as String2.
'
' Returns:
'     String2
'
'---------------------------------------------------------------------------------------
Public Property Get ToString2() As String2
    Set ToString2 = String2.NewValue(m_vvalue)
End Property
