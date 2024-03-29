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
' Class Module: Decimal2
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Eine Ersatz-Klasse f�r den fehlenden Decimal-Datentyp
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/Decimal2.cls</file>
'  <use>math/MathTools.bas</use>
'  <use>text/StringTools.bas</use>
'  <use>text/String2.cls</use>
'  <test>_test\math\Decimal2Tests.cls</test>
'  <license>_codelib/license.bas</license>
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
'/**
' <summary>
' Gibt den aktuellen Wert gerundet zur�ck.
' </summary>
' <param name="numDigitsAfterDecimal">Zahl, die angibt, wie viele Stellen rechts vom Dezimalpunkt beim Runden ber�cksichtigt werden</param>
' <param name="midpointRoundingType">Rundungsverfahren</param>
' <returns>Decimal2</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Round(Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Decimal2
        
        Set Round = Decimal2.NewValue(MathTools.Round(m_vvalue, numDigitsAfterDecimal, midpointRoundingType))
End Property

'---------------------------------------------------------------------------------------
' Property: ToString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt den aktuellen Wert als String zur�ck.
' </summary>
' <param name="FormatString">Ein g�ltiger benannter oder benutzerdefinierter Formatausdruck inkl. Erweiterung f�r Stundenanzeige �ber 24 (Standard-Formatanweisungen siehe VBA.Format)</param>
' <param name="FirstDayOfWeek">Wird an VBA.Format weitergereicht</param>
' <param name="FirstWeekOfYear">Wird an VBA.Format weitergereicht</param>
' <returns>String</returns>
' <remarks>
' Optional k�nnen die verf�garen Parameter der Format-Funktion �bergeben werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ToString(Optional ByVal FormatString As Variant = vbNullString, _
                         Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
                         Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
        ToString = StringTools.Format(m_vvalue, FormatString, FirstDayOfWeek, FirstWeekOfYear)
End Property

'---------------------------------------------------------------------------------------
' Property: ToString2
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt den aktuellen Wert als String2 zur�ck.
' </summary>
' <returns>String2</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ToString2() As String2
    Set ToString2 = String2.NewValue(m_vvalue)
End Property
