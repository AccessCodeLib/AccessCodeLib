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
'  <use>math/Decimal2.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Property Get MaxValue() As Decimal2
    Set MaxValue = Decimal2.NewValue("+79228162514264337593543950335")
End Property

Public Property Get MinValue() As Decimal2
    Set MinValue = Decimal2.NewValue("-79228162514264337593543950335")
End Property

Public Property Get e() As Decimal2
    Set e = Decimal2.NewValue("2,71828182845904523536028747135")
End Property

Public Property Get EulerConstant() As Decimal2
    Set EulerConstant = Decimal2.NewValue("0,5772156649015328606065120901")
End Property

Public Property Get epsilon() As Decimal2
    Set epsilon = Decimal2.NewValue("0,0000000000000000000000000001")
End Property

Public Property Get GoldenRatio() As Decimal2
    Set GoldenRatio = Decimal2.NewValue("1,61803398874989484820458683436")
End Property

Public Property Get pi() As Decimal2
    Set pi = Decimal2.NewValue("3,14159265358979323846264338323")
End Property
