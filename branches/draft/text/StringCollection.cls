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
Option Compare Database
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
                         Optional ByVal ItemSuffix As String = vbNullString) As String
   
   Dim s As String

   s = VBA.Join(ToStringArray(), ItemSuffix & Delimiter & ItemPrefix)
   If Len(s) > 0 Then s = ItemPrefix & s & ItemSuffix

   ToString = s

End Function

Public Function ToStringArray() As String()

   Dim arrItems() As String
   Dim MaxArrayIndex As Long
   Dim i As Long

   MaxArrayIndex = m_Items.Count - 1

   If MaxArrayIndex < 0 Then
      ToStringArray = arrItems
      Exit Function
   End If

   ReDim arrItems(0 To MaxArrayIndex)
   For i = 0 To MaxArrayIndex
      arrItems(i) = m_Items.Item(i + 1)
   Next

   ToStringArray = arrItems

End Function
