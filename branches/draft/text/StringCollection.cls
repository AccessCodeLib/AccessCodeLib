VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StringCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: StringCollection
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Collection-Funktionen für Strings
' </summary>
' <remarks></remarks>
' \ingroup text
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>text/StringCollection.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_Items As Collection

Private Sub Class_Initialize()
   Set m_Items = New Collection
End Sub

Private Sub Class_Terminate()
   Set m_Items = Nothing
End Sub

Public Property Get Self() As StringCollection
   Set Self = Me
End Property

Public Property Get Items() As Collection
   Set Items = m_Items
End Property

Public Property Get Item(ByVal Index As Variant) As String
   Item = m_Items.Item(Index)
End Property

Public Property Let Item(ByVal Index As Variant, ByVal NewValue As String)
   m_Items.Add NewValue, , , Index
   m_Items.Remove Index
End Property

Public Sub Add(ByVal Item As String)
   m_Items.Add Item
End Sub

Public Sub AddFromArray(ByRef ArrayToAdd As Variant, Optional ByVal ItemStringFormat As String = vbNullString)

   Dim i As Long

   For i = LBound(ArrayToAdd) To UBound(ArrayToAdd)
      m_Items.Add Format(ArrayToAdd(i), ItemStringFormat)
   Next

End Sub

Public Sub AddFromCollection(ByVal CollectionRef As Object, Optional ByVal ItemStringFormat As String = vbNullString)
'Object, damit alle Collections mit Enumarable- u. Item(index)-Interface durchlaufen werden können

   Dim varItm As Variant

   For Each varItm In CollectionRef
      m_Items.Add Format(varItm, ItemStringFormat)
   Next

End Sub

Public Function ToString(Optional ByVal Delimiter As String = ", ", _
                         Optional ByVal ItemPrefix As String = vbNullString, _
                         Optional ByVal ItemSuffix As String = vbNullString, _
                         Optional ByVal IgnoreEmptyValue As Boolean = False, _
                         Optional ByVal IgnoreDuplicateValues As Boolean = False) As String
   
   Dim s As String

   s = VBA.Join(ToStringArray(IgnoreEmptyValue, IgnoreDuplicateValues), ItemSuffix & Delimiter & ItemPrefix)
   If Len(s) > 0 Then s = ItemPrefix & s & ItemSuffix

   ToString = s

End Function

Public Function ToStringArray(Optional ByVal IgnoreEmptyValue As Boolean = False, _
                              Optional ByVal IgnoreDuplicateValues As Boolean = False) As String()

   Dim arrItems() As String
   Dim MaxArrayIndex As Long
   Dim i As Long

   MaxArrayIndex = m_Items.Count - 1

   If MaxArrayIndex < 0 Then
      ToStringArray = arrItems
      Exit Function
   End If
   
   If IgnoreEmptyValue Then
      If IgnoreDuplicateValues Then
         ToStringArray = RemoveDuplicateValues(GetArrayWithoutEmptyValues())
      Else
         ToStringArray = GetArrayWithoutEmptyValues()
      End If
      Exit Function
   End If

   ReDim arrItems(0 To MaxArrayIndex)
   For i = 0 To MaxArrayIndex
      arrItems(i) = m_Items.Item(i + 1)
   Next

   If IgnoreDuplicateValues Then
      ToStringArray = RemoveDuplicateValues(arrItems)
   Else
      ToStringArray = arrItems
   End If
   
End Function

Private Function GetArrayWithoutEmptyValues() As String()

   Dim arrItems() As String
   Dim MaxArrayIndex As Long
   Dim ItemIndex As Long
   Dim varItm As Variant

   MaxArrayIndex = m_Items.Count - 1

   If MaxArrayIndex < 0 Then
      GetArrayWithoutEmptyValues = arrItems
      Exit Function
   End If

   ReDim arrItems(0 To MaxArrayIndex)
   ItemIndex = -1
   For Each varItm In m_Items
      If Len(varItm) > 0 Then
         ItemIndex = ItemIndex + 1
         arrItems(ItemIndex) = varItm
      End If
   Next
   
   If ItemIndex = -1 Then
      Erase arrItems
      GetArrayWithoutEmptyValues = arrItems
      Exit Function
   End If
   
   If ItemIndex < (m_Items.Count - 1) Then
      ReDim Preserve arrItems(0 To ItemIndex)
   End If
   
   GetArrayWithoutEmptyValues = arrItems

End Function

Private Function RemoveDuplicateValues(ByRef ArrayToCheck() As String) As String()

   Dim arrItems() As String
   Dim MaxArrayIndex As Long
   Dim ItemIndex As Long
   Dim varItm As Variant

   MaxArrayIndex = UBound(ArrayToCheck)

   If MaxArrayIndex = 0 Then
      RemoveDuplicateValues = ArrayToCheck
      Exit Function
   End If
   
   ReDim arrItems(MaxArrayIndex)
   
   ItemIndex = -1
   For Each varItm In ArrayToCheck
      If Not ValueExistsInArray(arrItems, varItm, ItemIndex) Then
         ItemIndex = ItemIndex + 1
         arrItems(ItemIndex) = varItm
         Debug.Print varItm
      End If
   Next
   
   If ItemIndex < (m_Items.Count - 1) Then
      ReDim Preserve arrItems(0 To ItemIndex)
   End If
   
   RemoveDuplicateValues = arrItems
   
End Function

Private Function ValueExistsInArray(ByRef ArrayToCheck() As String, ByVal ValueToCheck As String, ByVal CheckUntilArrayIndex As Long) As Boolean
   
   Dim i As Long
   
   If CheckUntilArrayIndex < 0 Then
      Exit Function
   End If
   
   For i = LBound(ArrayToCheck) To CheckUntilArrayIndex
      If StrComp(ArrayToCheck(i), ValueToCheck, vbBinaryCompare) = 0 Then
         ValueExistsInArray = True
         Exit Function
      End If
   Next
   
End Function
