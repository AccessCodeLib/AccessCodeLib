Attribute VB_Name = "StringCollection_Examples"
'---------------------------------------------------------------------------------------
' Class Module: StringCollection_Examples
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Beispiel zur Verwendung der FilterStringBuilder-Klasse
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text\StringCollection_Examples.bas</file>
'  <use>text\StringCollection.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private Sub Werte_anfuegen_und_mit_Komma_getrennt_als_String_ausgeben()

   With New StringCollection
      .Add "a"
      .Add "b"
      .Add "c"
      Debug.Print .ToString(", ")
   End With

End Sub

Private Sub Werte_anfuegen_ändern_und_mit_Komma_getrennt_als_String_ausgeben()

   With New StringCollection

      .Add "a"
      .Add "b"
      .Add "c"

      .Item(2) = "bx"

      Debug.Print .ToString(", ")

   End With

End Sub
