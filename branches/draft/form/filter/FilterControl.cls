VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControl
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Formular-Steuerelement mit Filtereinstellungen
' </summary>
' <remarks>Wird in FilterControlCollection und FilterControlManager verwendet</remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControl.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/SqlTools.cls</use>
'  <use>form/filter/FilterControlEventBridge.cls<</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_EventBridge As FilterControlEventBridge
Attribute m_EventBridge.VB_VarHelpID = -1
Private m_SqlTool As SqlTools

Private m_DataField As String
Private m_DataType As SqlFieldDataType
Private m_RelationalOperator As SqlRelationalOperators
Private m_IgnoreValue As Variant

Private m_Control As Control
Private WithEvents m_Textbox As TextBox
Attribute m_Textbox.VB_VarHelpID = -1
Private WithEvents m_ComboBox As ComboBox
Attribute m_ComboBox.VB_VarHelpID = -1
Private WithEvents m_ListBox As ListBox
Attribute m_ListBox.VB_VarHelpID = -1
Private WithEvents m_ToggleButton As ToggleButton
Attribute m_ToggleButton.VB_VarHelpID = -1
Private WithEvents m_CheckBox As CheckBox
Attribute m_CheckBox.VB_VarHelpID = -1

Private Const EventProcedurePropertyTag As String = "[Event Procedure]"

Private m_FilterControl2 As FilterControl
Private m_SubFilterControls As Collection

Private Sub InitControl()

   RemoveWithEventsControlRefs

   Select Case m_Control.ControlType
      Case acTextBox
         Set m_Textbox = m_Control
      Case acComboBox
         Set m_ComboBox = m_Control
      Case acListBox
         Set m_ListBox = m_Control
      Case acToggleButton
         Set m_ToggleButton = m_Control
      Case acCheckBox
         Set m_CheckBox = m_Control
      Case Else
   
   End Select
   
   CheckAfterUpdateEventHandler

End Sub

Private Sub CheckAfterUpdateEventHandler()

   If Len(m_Control.AfterUpdate) = 0 Then
      m_Control.AfterUpdate = EventProcedurePropertyTag
   End If

End Sub

Private Sub RemoveWithEventsControlRefs()
   Set m_Textbox = Nothing
   Set m_ComboBox = Nothing
   Set m_ListBox = Nothing
   Set m_ToggleButton = Nothing
   Set m_CheckBox = Nothing
End Sub

Private Sub Class_Initialize()
   m_IgnoreValue = Null
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Friend Sub Dispose()
   Set m_EventBridge = Nothing
   Set m_SubFilterControls = Nothing
   Set m_FilterControl2 = Nothing
   Set m_Control = Nothing
End Sub

Friend Property Get EventBridge() As FilterControlEventBridge
   Set EventBridge = m_EventBridge
End Property

Friend Property Set EventBridge(ByVal NewRef As FilterControlEventBridge)
   Set m_EventBridge = NewRef
   If Not (m_FilterControl2 Is Nothing) Then
      Set m_FilterControl2.EventBridge = m_EventBridge
   End If

   If m_SubFilterControls Is Nothing Then Exit Property
   If m_SubFilterControls.Count = 0 Then Exit Property
   
   Dim fc As FilterControl
   For Each fc In m_SubFilterControls
      fc.EventBridge = m_EventBridge
   Next

End Property


Friend Property Get SqlTool() As SqlTools
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.Clone
   End If
   Set SqlTool = m_SqlTool
End Property

Friend Property Set SqlTool(ByVal NewRef As SqlTools)
   Set m_SqlTool = NewRef
End Property

Friend Sub ConfigSqlFormat(ByVal SqlDateFormat As String, _
                           ByVal SqlBooleanTrueString As String, _
                           ByVal SqlWildCardString As String)
   
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = SqlTools.NewInstance(SqlDateFormat, SqlBooleanTrueString, SqlWildCardString)
      Exit Sub
   End If

   With SqlTool
      .SqlDateFormat = SqlDateFormat
      .SqlBooleanTrueString = SqlBooleanTrueString
      .SqlWildCardString = SqlWildCardString
   End With

End Sub

Private Sub RaiseFilterValueChanged()
   If m_EventBridge Is Nothing Then Exit Sub
   m_EventBridge.RaiseFilterValueChanged Me
End Sub

Public Property Get Self() As FilterControl
   Set Self = Me
End Property

Public Property Get Control() As Control
   Set Control = m_Control
End Property

Public Property Set Control(ByVal NewRef As Control)
   Set m_Control = NewRef
   InitControl
End Property

Public Property Get Control2() As Control
   If m_FilterControl2 Is Nothing Then
      Set Control2 = Nothing
   Else
      Set Control2 = m_FilterControl2.Control
   End If
End Property

Public Property Set Control2(ByVal NewRef As Control)
   Set m_FilterControl2 = New FilterControl
   With m_FilterControl2
      Set .Control = NewRef
      Set .EventBridge = m_EventBridge
   End With
End Property

Public Property Get SubControls() As Collection
   If m_SubFilterControls Is Nothing Then
      Set SubControls = Nothing
   Else
      Set SubControls = m_SubFilterControls
   End If
End Property

Public Sub SetSubControls(ByVal SubControlArray As Variant)

   Dim ctl As Variant

   Set m_SubFilterControls = New Collection
   For Each ctl In SubControlArray
      With New FilterControl
         Set .Control = ctl
         Set .EventBridge = m_EventBridge
         m_SubFilterControls.Add .Self
      End With
   Next

End Sub

Public Property Get DataField() As String
   DataField = m_DataField
End Property

Public Property Let DataField(ByVal NewValue As String)
   m_DataField = NewValue
End Property

Public Property Get DataType() As SqlFieldDataType
   DataType = m_DataType
End Property

Public Property Let DataType(ByVal NewValue As SqlFieldDataType)
   m_DataType = NewValue
End Property

Public Property Get RelationalOperator() As SqlRelationalOperators
   RelationalOperator = m_RelationalOperator
End Property

Public Property Let RelationalOperator(ByVal NewValue As SqlRelationalOperators)
   m_RelationalOperator = NewValue
End Property

Public Property Get IgnoreValue() As Variant
   IgnoreValue = m_IgnoreValue
End Property

Public Property Let IgnoreValue(ByVal NewValue As Variant)
   m_IgnoreValue = NewValue
End Property

Public Property Get CriteriaString() As String
   If m_SubFilterControls Is Nothing Then
      CriteriaString = SqlTool.BuildCriteria(m_DataField, m_DataType, m_RelationalOperator, GetControlValue(Control), GetControlValue(Control2), IgnoreValue)
   Else
      CriteriaString = SqlTool.BuildCriteria(m_DataField, m_DataType, m_RelationalOperator, GetSubControlsValues, , IgnoreValue)
   End If
End Property

Private Function GetSubControlsValues() As Variant()
   
   Dim fc As FilterControl
   Dim i As Long
   Dim Values() As Variant
   ReDim Values(0 To (m_SubFilterControls.Count - 1))

   For Each fc In m_SubFilterControls
      Select Case fc.Control.ControlType
         Case acCheckBox, acToggleButton
            If fc.Control.Value = True Then
               Values(i) = fc.Control.Tag
            Else
               Values(i) = m_IgnoreValue
            End If
         Case Else
            Values(i) = fc.Control.Value
      End Select
      i = i + 1
   Next

   GetSubControlsValues = Values

End Function

Private Function GetControlValue(ByVal ctl As Control) As Variant
   
   If ctl Is Nothing Then
      GetControlValue = Null
      Exit Function
   End If
   
   If TypeOf ctl Is ListBox Then
      GetControlValue = GetListBoxFilterValue(ctl)
      Exit Function
   End If

   GetControlValue = ctl.Value
   
End Function

Private Function GetListBoxFilterValue(ByVal lb As ListBox) As Variant
   
   If lb.MultiSelect > 0 Then
      GetListBoxFilterValue = GetSelectedListboxItemArray(lb, lb.BoundColumn - 1)
   Else
      GetListBoxFilterValue = lb.Value
   End If
   
End Function

Private Function GetSelectedListboxItemArray(ByVal lst As ListBox, ByVal ColumnIndex As Long) As Variant()
   
   Dim arrItems() As Variant
   Dim MaxArrIndex As Long
   Dim SelectedItems As [_ItemsSelected]

   Set SelectedItems = lst.ItemsSelected
   MaxArrIndex = SelectedItems.Count - 1
   If MaxArrIndex < 0 Then
      GetSelectedListboxItemArray = arrItems
      Exit Function
   End If

   ReDim arrItems(MaxArrIndex)
   Dim i As Long

   For i = 0 To MaxArrIndex
      arrItems(i) = lst.Column(ColumnIndex, SelectedItems.Item(i))
   Next

   GetSelectedListboxItemArray = arrItems

End Function

Private Sub m_CheckBox_AfterUpdate()
   RaiseFilterValueChanged
End Sub

Private Sub m_ComboBox_AfterUpdate()
   RaiseFilterValueChanged
End Sub

Private Sub m_ListBox_AfterUpdate()
   RaiseFilterValueChanged
End Sub

Private Sub m_Textbox_AfterUpdate()
   RaiseFilterValueChanged
End Sub

Private Sub m_ToggleButton_AfterUpdate()
   RaiseFilterValueChanged
End Sub

Public Sub RemoveFilterValue(Optional ByVal RaiseFilterValueChangedEvent As Boolean = True)

   If Not (m_Control Is Nothing) Then
      Select Case m_Control.ControlType
         Case acCheckBox, acToggleButton
            m_Control.Value = False
         Case Else
            m_Control.Value = Null
      End Select
   End If

   If Not (m_FilterControl2 Is Nothing) Then
      m_FilterControl2.RemoveFilterValue False
   End If

   If Not (m_SubFilterControls Is Nothing) Then
      Dim fc As FilterControl
      For Each fc In m_SubFilterControls
         fc.RemoveFilterValue False
      Next
   End If

   If RaiseFilterValueChangedEvent Then
      RaiseFilterValueChanged
   End If
End Sub
