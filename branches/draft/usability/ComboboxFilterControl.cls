VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ComboboxFilterControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: ComboboxFilterControl
'---------------------------------------------------------------------------------------
'/**
' \author   Josef Poetzl
' <summary>
' Combobox-Steuerung zum Suchen/Auswählen von Einträgen innerhalb der Combobox-Liste
' </summary>
' <remarks>Beispiel für Verwendung siehe ComboboxFilterControl_Beispiele.frm</remarks>
'\ingroup usability
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>usability/ComboboxFilterControl.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private WithEvents m_Combobox As ComboBox
Attribute m_Combobox.VB_VarHelpID = -1

Private m_DataRecordset As Object 'DAO.Recordset oder ADODB.Recordset
Private m_BoundColumnIndex As Long
Private m_SearchColumnArray As Variant

Private m_InSearchMode As Boolean

Public Event Changed()
Public Event EmptySearchRecordset(ByRef CancelSearch As Boolean)

Public Sub Dispose()
   Set m_Combobox = Nothing
   Set m_DataRecordset = Nothing
   m_BoundColumnIndex = 0
   m_SearchColumnArray = Array()
   m_InSearchMode = 0
End Sub

Public Sub Init(ByVal ComboboxRef As ComboBox, _
                ByVal SearchColumns As Variant, _
                ByVal DataRecordset As Object)

   Const EventProcedurePropertyTag As String = "[Event Procedure]"

   Set m_Combobox = ComboboxRef

   With m_Combobox
   
      m_BoundColumnIndex = .BoundColumn - 1

      Set m_DataRecordset = DataRecordset

      If Len(.AfterUpdate) = 0 Then
         .AfterUpdate = EventProcedurePropertyTag
      End If

      If Len(.OnEnter) = 0 Then
         .OnEnter = EventProcedurePropertyTag
      End If

      If Len(.OnExit) = 0 Then
         .OnExit = EventProcedurePropertyTag
      End If

      If Len(.OnNotInList) = 0 Then
         .OnNotInList = EventProcedurePropertyTag
      End If

      If Len(.OnKeyDown) = 0 Then
         .OnKeyDown = EventProcedurePropertyTag
      End If

   End With

   m_SearchColumnArray = SearchColumns

End Sub

Private Sub m_Combobox_Enter()
   With m_Combobox
      If Len(.RowSource) = 0 Then
         Set .Recordset = m_DataRecordset
      End If
   End With
End Sub

Private Sub m_Combobox_KeyDown(KeyCode As Integer, Shift As Integer)

   Select Case KeyCode
      Case vbKeyReturn, vbKeyTab, vbKeyLeft, vbKeyRight
         'Damit Change-Ereignis auf jeden Fall eintritt
         With m_Combobox
            If .LimitToList = False Or Len(.Text) = 0 Then
               .Text = m_Combobox.Text
            End If
         End With
   End Select
   
End Sub

Private Sub m_Combobox_NotInList(NewData As String, Response As Integer)
   CheckFilterValue m_Combobox.Text
   Response = acDataErrContinue
End Sub

Private Sub m_Combobox_AfterUpdate()

   With m_Combobox
   
      CheckFilterValue .Value
      If Not m_InSearchMode Then
         RaiseEvent Changed
      End If

   End With

End Sub

Private Sub m_Combobox_Exit(Cancel As Integer)
   'Ausstieg verhindern, bis der Suchmodus deaktiviert wurde
   ' (Ende Suchmodus = ein vorhandener Eintrag wurde ausgewählt oder Value ist NULL)
   If m_InSearchMode Then
      m_Combobox.Dropdown
      Cancel = True
   End If
End Sub

Private Sub CheckFilterValue(Optional ByVal FilterValue As Variant)

   Dim QueryAlias As String
   Dim QueryRecordset As DAO.Recordset
   Dim bolUseDataRecordset As Boolean

   If IsMissing(FilterValue) Then
      FilterValue = m_Combobox.Value
   End If

   bolUseDataRecordset = IsNull(FilterValue) Or (Len(FilterValue) = 0)
   If Not bolUseDataRecordset Then
      bolUseDataRecordset = FilterValueInBoundColumnData(FilterValue, Not m_Combobox.LimitToList)
   End If
   
   'Problem bei ADODB-Recordset:
   With m_Combobox
      Debug.Assert Nz(.Value = .Column(m_BoundColumnIndex), True)
      If .Value <> .Column(m_BoundColumnIndex) Then
         ' Column(x) zeigt einen Wert, obwohl .Value keinem Wert in der gebundenen Spalte entspricht
         Debug.Print "Value = " & .Value, "Column(" & m_BoundColumnIndex & ") = " & .Column(m_BoundColumnIndex)
         ' Wie kann man das vermeiden?
      End If
   End With
   
   If bolUseDataRecordset Then
      If m_InSearchMode Then
         DeactivateSearchMode
      Else
      '   SetComboboxRecordset m_DataRecordset
      End If
      Exit Sub
   End If

   ActivateSearchMode FilterValue
  
End Sub

Private Sub ActivateSearchMode(Optional ByVal FilterValue As Variant)

   Dim Cancel As Boolean
   Dim rst As Object 'DAO.Recordset oder ADODB.Recordset

   m_Combobox.Value = Null
      
   Set rst = GetFilterRecordset(FilterValue, m_SearchColumnArray)
   If rst.EOF Then
      RaiseEvent EmptySearchRecordset(Cancel)
      If Cancel Then
         If m_InSearchMode Then
            Set rst = m_Combobox.Recordset
         Else
            Set rst = m_DataRecordset
         End If
      End If
   End If
   
   m_InSearchMode = True
   SetComboboxRecordset rst
   
End Sub

Private Sub DeactivateSearchMode()
   m_InSearchMode = False
   SetComboboxRecordset m_DataRecordset
End Sub

Private Sub SetComboboxRecordset(ByRef rst As Object)

   With m_Combobox
   
      If Not (.Recordset Is rst) Then
         If Not (.Recordset Is m_DataRecordset) Then
            .Recordset.Close
         End If
      End If
      
      Set .Recordset = rst
      
   End With
   
End Sub

Private Function FilterValueInBoundColumnData(ByVal FilterValue As Variant, ByVal UseWildCardFilterOption As Boolean) As Boolean

   Dim FilterColArray(0) As Long
   Dim rst As Object 'DAO.Recordset oder ADODB.Recordset

   If m_Combobox.Column(m_BoundColumnIndex) = FilterValue Then
   'direkter Treffer
      FilterValueInBoundColumnData = True
      Exit Function
   End If
   
   If Not UseWildCardFilterOption Then
      FilterValueInBoundColumnData = False
      Exit Function
   End If
   
   FilterColArray(0) = m_BoundColumnIndex
   Set rst = GetFilterRecordset(FilterValue, FilterColArray)
   If Not rst.EOF Then
      'rst.Close
      FilterValueInBoundColumnData = True
   End If
   
End Function

Private Function GetFilterRecordset(ByVal FilterValue As Variant, ByVal SearchColumnArray As Variant) As Object 'DAO.Recordset
   
   Dim rst As Object 'DAO.Recordset oder ADODB.Recordset
   Dim FilterString As String
   Dim i As Long
   
   Set rst = m_Combobox.Recordset.Clone
   
   If IsNull(FilterValue) Then
      Set GetFilterRecordset = rst
      Exit Function
   End If
   
   For i = LBound(SearchColumnArray) To UBound(SearchColumnArray)
      FilterString = FilterString & " OR " & rst.Fields(SearchColumnArray(i)).Name & " like '" & FilterValue & "'"
   Next
   If Len(FilterString) > 0 Then
      FilterString = Mid(FilterString, 5) ' " OR " abschneiden
   End If
   
   If TypeOf rst Is DAO.Recordset Then
      rst.Filter = FilterString
      Set rst = rst.OpenRecordset
   Else 'ADODB
      '??? Warum muss man hier mit * statt mit % eine Wildcard-Suche machen?
      Set rst = rst.Clone
      rst.Filter = FilterString
   End If
   
   Set GetFilterRecordset = rst

End Function
