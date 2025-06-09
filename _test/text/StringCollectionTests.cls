VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringCollectionTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' AccUnit:TestClass
' StringCollectionTests
'---------------------------------------------------------------------------------------
'/**
' <summary>
' AccUnit-Testklasse für StringCollection
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/text/StringCollectionTests.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>text/StringCollection.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'--------------------------------------------------------------------
' Tests
'--------------------------------------------------------------------

Public Sub Add_Strings_CheckArray()
   
   Dim Expected() As String
   Expected = Split("a,b,c", ",")
   
   Dim Actual As Variant
   
   With New StringCollection
      
      .Add "a"
      .Add "b"
      .Add "c"
      
      Actual = .ToStringArray()
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub Add_Strings_CheckToString()
   
   Const Expected As String = "a,b,c"
   
   Dim Actual As Variant
   
   With New StringCollection
      
      .Add "a"
      .Add "b"
      .Add "c"
      
      Actual = .ToString(",")
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub AddFromArray_StringArray_CheckToString()
   
   Const Expected As String = "a,b,c,x,y,z"
   
   Dim InputArray1() As String
   Dim InputArray2() As String
   
   InputArray1 = Split("a,b,c", ",")
   InputArray2 = Split("x,y,z", ",")
   
   Dim Actual As Variant
   
   With New StringCollection
      
      .AddFromArray InputArray1
      .AddFromArray InputArray2
      
      Actual = .ToString(",")
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

'AccUnit:Row("", ",", "1,2,3,4,5")
'AccUnit:Row("(#)", ", ", "(1), (2), (3), (4), (5)")
'AccUnit:Row("00", "; ", "01; 02; 03; 04; 05")
Public Sub AddFromArray_LongIntArrayWithFormat_CheckToString( _
         ByVal ItemStringFormat As String, _
         ByVal ToStringDelimiter As String, _
         ByVal Expected As String)
   
   Dim InputArray1(2) As Long
   Dim InputArray2(1) As Long
   
   Dim i As Long
   Dim varItm As Variant
   For i = LBound(InputArray1) To UBound(InputArray1)
      InputArray1(i) = i + 1
   Next
   
   For i = LBound(InputArray2) To UBound(InputArray2)
      InputArray2(i) = i + 4
   Next
   
   Dim Actual As Variant
   
   With New StringCollection
      
      .AddFromArray InputArray1, ItemStringFormat
      .AddFromArray InputArray2, ItemStringFormat
      
      Actual = .ToString(ToStringDelimiter)
      
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ToSring_AddPrefixAndSuffix_CheckToString()
   
   Const Expected As String = "*a*,**,*b*,*c*,**,*b*,*a*"
      
   Dim Actual As Variant
   
   With New StringCollection
      .AddFromArray Array("a", "", "b", "c", "", "b", "a")
      Actual = .ToString(",", "*", "*", False)
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ToSring_IgnoreEmptyStringsAddPrefix_CheckToString()
   
   Const Expected As String = "*a,*b,*c"
      
   Dim Actual As Variant
   
   With New StringCollection
      .AddFromArray Array("a", "", "b", "c")
      Actual = .ToString(",", "*", , True)
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ToSring_AddSuffixIgnoreDuplicateValues_CheckToString()
   
   Const Expected As String = "*a*,**,*b*,*c*"
      
   Dim Actual As Variant
   
   With New StringCollection
      .AddFromArray Array("a", "", "b", "c", "", "b", "a")
      Actual = .ToString(",", "*", "*", False, True)
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub

Public Sub ToSring_AddIgnoreDuplicateAndEmptyValues_CheckToString()
   
   Const Expected As String = "a,b,c"
      
   Dim Actual As Variant
   
   With New StringCollection
      .AddFromArray Array("a", "", "b", "c", "", "b", "a")
      Actual = .ToString(",", , , True, True)
   End With
   
   Assert.That Actual, Iz.EqualTo(Expected)

End Sub
