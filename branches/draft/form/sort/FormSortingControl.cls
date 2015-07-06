VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormSortingControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: FormSortingControl
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsklasse für FormSortingExtension
' </summary>
' <remarks>
' </remarks>
' \ingroup form
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/sort/FormSortingControl.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/sort/FormSortingControlEventBridge.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const EventProcedureString = "[Event Procedure]"
Private Const OrderByAscDescPlaceholder As String = "[ASC/DESC]"

Private Const DefaultSortSymbolAscStringUnicode As Long = &H2191  ' Pfeil nach oben
Private Const DefaultSortSymbolDescStringUnicode As Long = &H2193 ' Pfeil nach unten
Private m_SortSymbolStrings(0 To 1) As String

Private WithEvents m_EventBridge As FormSortingControlEventBridge
Attribute m_EventBridge.VB_VarHelpID = -1

Private WithEvents m_Label As Access.Label
Attribute m_Label.VB_VarHelpID = -1
Private WithEvents m_Button As Access.CommandButton
Attribute m_Button.VB_VarHelpID = -1
Private m_InitialControlCaption As String

Private m_OrderBy As String

Private Sub Class_Initialize()
   SortSymbolString(0) = ChrW(DefaultSortSymbolAscStringUnicode)
   SortSymbolString(1) = ChrW(DefaultSortSymbolDescStringUnicode)
End Sub

Public Property Get This() As FormSortingControl
   Set This = Me
End Property

Public Property Get EventBridge() As FormSortingControlEventBridge
   Set EventBridge = m_EventBridge
End Property

Public Property Set EventBridge(ByVal NewRef As FormSortingControlEventBridge)
   Set m_EventBridge = NewRef
End Property

Public Property Get Control() As Access.Control
   If Not (m_Label Is Nothing) Then
      Set Control = m_Label
   ElseIf Not (m_Button Is Nothing) Then
      Set Control = m_Button
   End If
End Property

Public Property Set Control(ByVal NewRef As Access.Control)
   
   ClearControlReferences
   
   Select Case NewRef.ControlType
      Case acLabel
         SetSortingLabel NewRef
      Case acCommandButton
         SetSortingButton NewRef
      Case Else
         Err.Raise vbObjectError, "FormSortingControl.Control", "Type of control '" & NewRef.Name & "' is not supported!'"
   End Select
   
   m_InitialControlCaption = NewRef.Caption
   
End Property

Private Sub ClearControlReferences()
   Set m_Label = Nothing
   Set m_Button = Nothing
End Sub

Private Sub SetSortingLabel(ByVal LabelRef As Access.Label)
   Set m_Label = LabelRef
   With m_Label
      If Len(.OnClick) = 0 Then
         .OnClick = EventProcedureString
      End If
   End With
End Sub

Private Sub SetSortingButton(ByVal ButtonRef As Access.CommandButton)
   Set m_Button = ButtonRef
   With m_Button
      If Len(.OnClick) = 0 Then
         .OnClick = EventProcedureString
      End If
   End With
End Sub

Public Property Get OrderByStatement() As String
   OrderByStatement = m_OrderBy
End Property

Public Property Let OrderByStatement(ByVal NewValue As String)

   Dim OrderByAttributes() As String
   Dim OrderByAttribute As String
   Dim i As Long
   
   NewValue = Replace(NewValue, " " & OrderByAscDescPlaceholder, vbNullString)
   OrderByAttributes = Split(NewValue, ",")

   For i = LBound(OrderByAttributes) To UBound(OrderByAttributes)
      OrderByAttribute = Trim(OrderByAttributes(i))
      If Right(OrderByAttribute, 4) = " ASC" Then
         'hier nichts ändern / ASC ist fixierte Sortierung
      ElseIf Right(OrderByAttribute, 5) = " DESC" Then
         'hier nichts ändern / DESC ist fixierte Sortierung
      Else ' Platzhalter zum Ändern der Sortierrichtung einfügen
          OrderByAttribute = OrderByAttribute & " " & OrderByAscDescPlaceholder
      End If
      OrderByAttributes(i) = OrderByAttribute
   Next
   m_OrderBy = Join(OrderByAttributes, ", ")
   
End Property

Public Function GetOrderByString(ByVal UseDesc As Boolean)
   
   Dim DescString As String
   If UseDesc Then
      DescString = " DESC"
   End If
   
   GetOrderByString = Replace(OrderByStatement, " " & OrderByAscDescPlaceholder, DescString)

End Function

Private Sub m_Button_Click()
   RaiseSortingControlClick
End Sub

Private Sub m_Label_Click()
   RaiseSortingControlClick
End Sub

Private Sub RaiseSortingControlClick()
   m_EventBridge.RaiseSortingControlClick This
End Sub

Public Property Get SortSymbolString(ByVal SortDesc As Boolean) As String
   SortSymbolString = m_SortSymbolStrings(Abs(SortDesc))
End Property

Public Property Let SortSymbolString(ByVal SortDesc As Boolean, ByVal NewString As String)
   m_SortSymbolStrings(Abs(SortDesc)) = NewString
End Property

Public Sub SetSortSymbol(ByVal SortDesc As Boolean)
   Me.Control.Caption = m_InitialControlCaption & SortSymbolString(SortDesc)
End Sub

Public Sub RemoveSortSymbol()
   Me.Control.Caption = m_InitialControlCaption
End Sub
