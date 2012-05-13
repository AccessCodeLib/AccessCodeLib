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
            Round = VBA.Sgn(number) * VBA.Int(VBA.Conversion.CDec(2 ^ -1 + VBA.Sgn(number) * number * 10 ^ numDigitsAfterDecimal)) * 10 ^ -numDigitsAfterDecimal
            Exit Function
        Case Else
            Round = number
            Exit Function
    End Select
End Function

