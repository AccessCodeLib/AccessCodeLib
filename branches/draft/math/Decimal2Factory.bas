Attribute VB_Name = "Decimal2Factory"
'---------------------------------------------------------------------------------------
' Modul: Decimal
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Factory für den Decimal2-Datentyp
' </summary>
' <remarks></remarks>
'
' \ingroup math
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>math/Decimal.bas</file>
'  <use>math/Decimal2.cls</use>
'  <license>_codelib/license.bas</license>
'  <test>_test/math/DecimalTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Property Get Value(Optional ByVal initValue As Variant = 0) As Decimal2
    Dim decimal2Instance As New Decimal2
        decimal2Instance = initValue
    Set Value = decimal2Instance
    Set decimal2Instance = Nothing
End Property


