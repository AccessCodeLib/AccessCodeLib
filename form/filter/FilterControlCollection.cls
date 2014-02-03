VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlCollection
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Sammlung von Formular-Steuerelementen zur Filtererstellung
' </summary>
' <remarks>Wird in FilterControlManager verwendet</remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControlCollection.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControl.cls</use>
'  <use>form/filter/FilterControlEventBridge.cls<</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_FilterControls As Collection
Private WithEvents m_EventBridge As FilterControlEventBridge
Attribute m_EventBridge.VB_VarHelpID = -1
Private m_ItemConcatOperator As SqlLogicalOperator

Public Event FilterValuesChanged()

Private Sub Class_Initialize()
   Set m_FilterControls = New Collection
   Set m_EventBridge = New FilterControlEventBridge
   m_ItemConcatOperator = SqlLogicalOperator.SQL_And
End Sub

Friend Property Get EventBridge() As FilterControlEventBridge
   Set EventBridge = m_EventBridge
End Property

Friend Property Set EventBridge(ByVal NewRef As FilterControlEventBridge)

   Set m_EventBridge = NewRef
   
   Dim fc As FilterControl
   For Each fc In m_FilterControls
      Set fc.EventBridge = m_EventBridge
   Next

End Property

Public Property Get Self() As FilterControlCollection
   Set Self = Me
End Property

Public Sub Add(ByVal DataFieldName As String, ByVal DataType As SqlFieldDataType, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal ControlRef As Control, _
      Optional ByVal Control2Ref As Control = Nothing, _
      Optional ByVal IgnoreValue As Variant = Null)
   
   With New FilterControl
      Set .EventBridge = m_EventBridge
      Set .Control = ControlRef
      .DataField = DataFieldName
      .DataType = DataType
      .RelationalOperator = RelationalOperator
      If Not (Control2Ref Is Nothing) Then
         Set .Control2 = Control2Ref
      End If
      .IgnoreValue = IgnoreValue
      m_FilterControls.Add .Self
   End With
   
End Sub

Public Function AddGroup(Optional ByVal ConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_Or) As FilterControlCollection
   
   With New FilterControlCollection
      .ItemConcatOperator = ConcatOperator
      Set .EventBridge = m_EventBridge
      m_FilterControls.Add .Self
      Set AddGroup = .Self
   End With

End Function

Public Property Get ItemConcatOperator() As SqlLogicalOperator
   ItemConcatOperator = m_ItemConcatOperator
End Property

Public Property Let ItemConcatOperator(ByVal NewValue As SqlLogicalOperator)
   m_ItemConcatOperator = NewValue
End Property

Public Sub AddMultiControlCriteria( _
               ByVal DataFieldName As String, ByVal DataType As SqlFieldDataType, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal IgnoreValue As Variant, _
               ParamArray Controls() As Variant)
   
   With New FilterControl
      Set .EventBridge = m_EventBridge
      .SetSubControls Controls
      .DataField = DataFieldName
      .DataType = DataType
      .RelationalOperator = RelationalOperator
      .IgnoreValue = IgnoreValue
      m_FilterControls.Add .Self
   End With
   
End Sub

Public Property Get Items() As Collection
   Set Items = m_FilterControls
End Property

Public Sub RemoveFilterValues(Optional ByVal RaiseFilterValuesChangedEvent As Boolean = True)
   
   Dim fc As FilterControl
   For Each fc In Items
      fc.RemoveFilterValue False
   Next
   
   If RaiseFilterValuesChangedEvent Then
      RaiseEvent FilterValuesChanged
   End If
   
End Sub

Private Sub m_EventBridge_FilterValuesChanged(ByVal FilterControlRef As FilterControl)
   RaiseEvent FilterValuesChanged
End Sub

Public Sub Clear()

   Set m_EventBridge = Nothing
   Set m_FilterControls = Nothing
   Set m_FilterControls = New Collection
   Set m_EventBridge = New FilterControlEventBridge
   
End Sub

Public Function CriteriaString(ByVal SqlToolsRef As SqlTools) As String
   
   Dim fc As Object ' schöner wäre As FilterControl kominiert mit FilterControlCollection implements FilterControl
                    ' .. aber Implements ist aus Erfahrung mit Vorsicht zu genießen, wenn Code/Interface in Access/VBA geändert wird.

   With New FilterStringBuilder
      For Each fc In m_FilterControls
         .AddCriteria fc.CriteriaString(SqlToolsRef)
      Next
      CriteriaString = .ToString(m_ItemConcatOperator)
   End With

End Function
