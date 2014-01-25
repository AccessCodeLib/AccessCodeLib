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

Public Event FilterValuesChanged()

Private Sub Class_Initialize()
   Set m_FilterControls = New Collection
   Set m_EventBridge = New FilterControlEventBridge
End Sub

Public Sub Add(ByVal ControlRef As Control, _
               ByVal DataFieldName As String, ByVal DataType As SqlFieldDataType, ByVal RelationalOperator As SqlRelationalOperators, _
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

Public Sub AddMultiControlCriteria( _
               ByVal DataFieldName As String, ByVal DataType As SqlFieldDataType, ByVal RelationalOperator As SqlRelationalOperators, _
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
