Attribute VB_Name = "MathTools"
Attribute VB_Description = "Mathe-Hilfsfunktionen"
'---------------------------------------------------------------------------------------
' Modul: MathTools
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Mathematische Funktionen
' </summary>
' <remarks></remarks>
'
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/MathTools.bas</file>
'  <use>math/Decimal2Factory.bas</use>
'  <license>_codelib/license.bas</license>
'  <test>_test/math/MathToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------

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
Public Function Round(ByVal Number As Variant, _
                      Optional ByVal numDigitsAfterDecimal As Long = 0, _
                      Optional ByVal midpointRoundingType As MidpointRounding = MidpointRounding.ToEven) As Variant
                         
    Select Case midpointRoundingType
        Case MidpointRounding.ToEven
            Round = VBA.Round(Number, numDigitsAfterDecimal)
            Exit Function
        Case MidpointRounding.AwayFromZero
            Round = VBA.Sgn(Number) * VBA.Int(Decimal2Factory.Value(0.5) + VBA.Abs(Number) * 10 ^ numDigitsAfterDecimal) * 10 ^ -numDigitsAfterDecimal
            Exit Function
        Case Else
            Round = Number
            Exit Function
    End Select
End Function

'---------------------------------------------------------------------------------------
' Function: Fact
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Berechnet die Fakultät einer natürlichen Zahl
' </summary>
' <param name="number">Nummerischer Ausdruck vom dem die Fakultät berechnet wird.</param>
' <returns>Variant (Nummeric)</returns>
' <remarks>
' Im Wertebereich zwischen 0 und 27 wird das exakte Ergebnis als Decimal zurück
' gegeben. Zwischen 28 und 170 wird die Faktutät mit verringerter Genauigkeit
' als Double ermittelt. Bei allen übrigen Parametern wird Null zurückgegeben.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Fact(ByVal Number As Integer) As Variant
    
    Dim n As Integer
    Dim Result As Variant
        Result = 1
        
    Select Case Number
        Case 0 To 27
            For n = 1 To Number Step 1
                Result = Decimal2Factory.Value(Result) * n
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

Public Function DivWithNull() As Double
On Error GoTo HandleErr

    DivWithNull = 1 / 0
    
    
ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ExtAppFile.m_ApplicationHandler_ExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Function

