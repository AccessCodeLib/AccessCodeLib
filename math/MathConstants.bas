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
    Set e = Decimal.Value("2,718281828459045235360287")
End Property

Public Property Get pi() As Decimal2
    Set pi = Decimal.Value("3,14159265358979323846264338323")
End Property

