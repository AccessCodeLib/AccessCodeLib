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

Private Sub Werte_anfuegen_löschen_und_per_For_Each_Schleife_durchlaufen()

    Dim coll As StringCollection
    Set coll = New StringCollection
    
    With coll
        .Add "row 1"
        .Add "row 2"
        .Add "row 3"
        .Add "remove me"
        
        .Items.Remove 4
    End With
    
    Dim varItm As Variant

    For Each varItm In coll.Items
        Debug.Print varItm
    Next

    Set coll = Nothing

End Sub

Private Sub Werte_anfuegen_löschen_und_per_For_Schleife_rueckwärts_durchlaufen()

    Dim n As Long
    
    With New StringCollection
    
        .Add "row 1"
        .Add "row 2"
        .Add "row 3"
        .Add "remove me"
        
        .Items.Remove 4
    
        For n = .Items.Count To 1 Step -1
            Debug.Print .Item(n)
        Next n
    
    End With

End Sub
