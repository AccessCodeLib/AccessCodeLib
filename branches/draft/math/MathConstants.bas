Attribute VB_Name = "MathConstants"
'---------------------------------------------------------------------------------------
' Modul: MathConstants
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Mathematische Konstanten
' </summary>
' <remarks></remarks>
'
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/MathConstants.bas</file>
'  <use>math/Decimal.bas</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Property Get MaxValue() As Decimal2
    Set MaxValue = Decimal.Value("+79228162514264337593543950335")
End Property

Public Property Get MinValue() As Decimal2
    Set MinValue = Decimal.Value("-79228162514264337593543950335")
End Property

Public Property Get e() As Decimal2
    Set e = Decimal.Value("2,71828182845904523536028747135")
End Property

Public Property Get EulerConstant() As Decimal2
    Set EulerConstant = Decimal.Value("0,5772156649015328606065120901")
End Property

Public Property Get epsilon() As Decimal2
    Set epsilon = Decimal.Value("0,0000000000000000000000000001")
End Property

Public Property Get GoldenRatio() As Decimal2
    Set GoldenRatio = Decimal.Value("1,61803398874989484820458683436")
End Property

Public Property Get pi() As Decimal2
    Set pi = Decimal.Value("3,14159265358979323846264338323")
End Property

