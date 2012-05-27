VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Decimal2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: Decimal2
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Eine Ersatz-Klasse für den fehlenden Decimal-Datentyp
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/Decimal2.cls</file>
'  <use>math/Decimal.bas</use>
'  <use>math/MathTools.bas</use>
'  <use>text/StringTools.bas</use>
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

Public Property Get Value() As Variant
Attribute Value.VB_UserMemId = 0
    Value = m_vvalue
End Property

Public Property Let Value(ByVal vvalue As Variant)
Attribute Value.VB_UserMemId = 0
    m_vvalue = VBA.Conversion.CDec(vvalue)
End Property

'---------------------------------------------------------------------------------------
' Function: Round
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt den aktuellen Wert gerundet zurück.
' </summary>
' <param name="numDigitsAfterDecimal">Zahl, die angibt, wie viele Stellen rechts vom Dezimalpunkt beim Runden berücksichtigt werden</param>
' <param name="midpointRoundingType">Rundungsverfahren</param>
' <returns>Decimal2</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Round(Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Decimal2
        
        Set Round = Decimal.Value(MathTools.Round(m_vvalue, numDigitsAfterDecimal, midpointRoundingType))
End Function

'---------------------------------------------------------------------------------------
' Function: ToString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt den aktuellen Wert als String zurück.
' </summary>
' <param name="FormatString">Ein gültiger benannter oder benutzerdefinierter Formatausdruck inkl. Erweiterung für Stundenanzeige über 24 (Standard-Formatanweisungen siehe VBA.Format)</param>
' <param name="FirstDayOfWeek">Wird an VBA.Format weitergereicht</param>
' <param name="FirstWeekOfYear">Wird an VBA.Format weitergereicht</param>
' <returns>String</returns>
' <remarks>
' Optional können die verfügaren Parameter der Format-Funktion übergeben werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ToString(Optional ByVal FormatString As Variant, _
                         Optional ByVal FirstDayOfWeek As VbDayOfWeek = vbSunday, _
                         Optional ByVal FirstWeekOfYear As VbFirstWeekOfYear = vbFirstJan1) As String
                         
        ToString = StringTools.Format(m_vvalue, FormatString, FirstDayOfWeek, FirstWeekOfYear)
End Function


