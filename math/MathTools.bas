Attribute VB_Name = "MathTools"
Attribute VB_Description = "Mathe-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Package: math.MathTools
'---------------------------------------------------------------------------------------
'
' Mathematical functions
'
' Author:
'     Sten Schmidt
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/MathTools.bas</file>
'  <use>math/Decimal2.cls</use>
'  <license>_codelib/license.bas</license>
'  <test>_test/math/MathToolsTests.cls</test>
'  <description>Math functions: rounding with selectable midpoint rounding (to even or away from zero), factorial</description>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit
Option Private Module

'---------------------------------------------------------------------------------------
' Enum: MidpointRounding
'---------------------------------------------------------------------------------------
'
' Available rounding methods for the Round function
'
' ToEven (1) - Mathematical rounding according to IEEE 754
' AwayFromZero (2) - Commercial rounding (away from zero)
'
'---------------------------------------------------------------------------------------
Public Enum MidpointRounding
    ToEven
    AwayFromZero
End Enum

'---------------------------------------------------------------------------------------
' Function: Round
'---------------------------------------------------------------------------------------
'
' Replaces the VBA Round function and extends it by the rounding type AwayFromZero
'
' Parameters:
'     number                - Numeric expression to be rounded
'     numDigitsAfterDecimal - Number of digits to the right of the decimal point to round to
'     midpointRoundingType  - Rounding method
'
' Returns:
'     Variant (Nummeric)
'
'---------------------------------------------------------------------------------------
Public Function Round(ByVal Number As Variant, _
                      Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Variant
                         
    Select Case midpointRoundingType
        Case MidpointRounding.ToEven
            Round = VBA.Round(Number, numDigitsAfterDecimal)
            Exit Function
        Case MidpointRounding.AwayFromZero
            Round = VBA.Sgn(Number) * VBA.Int(Decimal2.NewValue(0.5) + VBA.Abs(Number) * 10 ^ numDigitsAfterDecimal) * 10 ^ -numDigitsAfterDecimal
            Exit Function
        Case Else
            Round = Number
            Exit Function
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function: Fact
'---------------------------------------------------------------------------------------
'
' Calculates the factorial of a natural number
'
' Parameters:
'     number - Numeric expression whose factorial is calculated.
'
' Returns:
'     Variant (Nummeric)
'
' Remarks:
'     For values between 0 and 27, the exact result is returned as Decimal.
'     Between 28 and 170, the factorial is calculated with reduced precision
'     as Double. For all other values, Null is returned.
'
'---------------------------------------------------------------------------------------
Public Function Fact(ByVal Number As Integer) As Variant
    
    Dim n As Integer
    Dim Result As Variant
        Result = 1
        
    Select Case Number
        Case 0 To 27
            For n = 1 To Number Step 1
                Result = Decimal2.NewValue(Result) * n
            Next n
        Case 28 To 170
            For n = 1 To Number Step 1
                Result = CDbl(Result) * n
            Next n
        Case Else
            Result = Null
    End Select
    
    Fact = Result
    Set Result = Nothing
    
End Function
