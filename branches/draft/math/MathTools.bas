Attribute VB_Name = "MathTools"
Attribute VB_Description = "Mathe-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Modul: MathTools
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Mathematische Hilfsfunktionen
' </summary>
' <remarks></remarks>
'
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/MathTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/math/MathToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Enum: MidpointRounding
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' Verfügbare Rundungsverfahren für die Round-Funktion
' </summary>
' <list type="table">
'   <item><term>ToEven (1)</term><description>Mathematisches Runden nach dem IEEE-754-Standard</description></item>
'   <item><term>AwayFromZero (2)</term><description>Kaufmännisches Runden</description></item>
' </list>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Enum MidpointRounding
    ToEven
    AwayFromZero
End Enum

'---------------------------------------------------------------------------------------
' Function: Round
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ersetzt die VBA Round-Funktion und erweitert diese um den Rundungstyp AwayFromZero
' </summary>
' <param name="number">Nummerischer Ausdruck der gerundet wird</param>
' <param name="numDigitsAfterDecimal">Zahl, die angibt, wie viele Stellen rechts vom Dezimalpunkt beim Runden berücksichtigt werden</param>
' <param name="midpointRoundingType">Rundungsverfahren</param>
' <returns>Variant (Nummeric)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Round(ByVal number As Variant, _
                      Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Variant
                         
    Select Case midpointRoundingType
        Case MidpointRounding.ToEven
            Round = VBA.Round(number, numDigitsAfterDecimal)
            Exit Function
        Case MidpointRounding.AwayFromZero
            Round = VBA.Sgn(number) * VBA.Int(MathTools.DecimalDivision(1, 2) + VBA.Abs(number) * 10 ^ numDigitsAfterDecimal) * 10 ^ -numDigitsAfterDecimal
            Exit Function
        Case Else
            Round = number
            Exit Function
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function: DecimalAddition
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Führt eine Addition unter Verwendung des Decimal-Datentyps durch
' </summary>
' <param name="valueA">1. Summand</param>
' <param name="valueB">2. Summand</param>
' <returns>Summe als Variant (Decimal)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DecimalAddition(ByVal valueA As Variant, ByVal valueB As Variant) As Variant
    DecimalAddition = MathTools.DecimalValue(valueA) + MathTools.DecimalValue(valueB)
End Function

'---------------------------------------------------------------------------------------
' Function: DecimalAddition
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Führt eine Subtraktion unter Verwendung des Decimal-Datentyps durch
' </summary>
' <param name="valueA">Minuend</param>
' <param name="valueB">Subtrahend</param>
' <returns>Differenz als Variant (Decimal)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DecimalSubtraction(ByVal valueA As Variant, ByVal valueB As Variant) As Variant
    DecimalSubtraction = MathTools.DecimalValue(valueA) - MathTools.DecimalValue(valueB)
End Function

'---------------------------------------------------------------------------------------
' Function: DecimalMultiplication
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Führt eine Multiplikation unter Verwendung des Decimal-Datentyps durch
' </summary>
' <param name="valueA">1. Faktor</param>
' <param name="valueB">2. Faktor</param>
' <returns>Produkt als Variant (Decimal)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DecimalMultiplication(ByVal valueA As Variant, ByVal valueB As Variant) As Variant
    DecimalMultiplication = MathTools.DecimalValue(valueA) * MathTools.DecimalValue(valueB)
End Function

'---------------------------------------------------------------------------------------
' Function: DecimalDivision
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Führt eine Division unter Verwendung des Decimal-Datentyps durch
' </summary>
' <param name="valueA">Dividend</param>
' <param name="valueB">Divisor</param>
' <returns>Quotient als Variant (Decimal)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DecimalDivision(ByVal valueA As Variant, ByVal valueB As Variant) As Variant
    DecimalDivision = MathTools.DecimalValue(valueA) / MathTools.DecimalValue(valueB)
End Function

'---------------------------------------------------------------------------------------
' Private Function: DecimalValue
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Führt eine Konvertierung in den Decimal-Datentyp durch
' </summary>
' <param name="value">Nummerischer Ausdruck</param>
' <returns>Nummerischer Ausdruck als Variant (Decimal)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Private Function DecimalValue(ByVal value As Variant) As Variant
    DecimalValue = VBA.Conversion.CDec(value)
End Function





