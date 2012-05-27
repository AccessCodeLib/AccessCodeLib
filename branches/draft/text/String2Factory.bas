Attribute VB_Name = "String2Factory"
'---------------------------------------------------------------------------------------
' Modul: String2Factory
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Factory für den String2-Datentyp
' </summary>
' <remarks></remarks>
'
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/String2Factory.bas</file>
'  <use>text/String2.cls</use>
'  <license>_codelib/license.bas</license>
'  <test>_test/text/String2FactoryTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Property Get Value(Optional ByVal initValue As Variant = vbNullString) As String2
    Dim string2Instance As New String2
        string2Instance = initValue
    Set Value = string2Instance
    Set string2Instance = Nothing
End Property
