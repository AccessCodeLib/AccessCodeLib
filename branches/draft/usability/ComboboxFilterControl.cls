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
Private m_SearchRecordset As Object

Private m_FilterValues As Collection

Public Event Changed()
Public Event EmptySearchRecordset(ByRef CancelSearch As Boolean)
Public Event MissingDataRecordset(ByRef rs As Object)

'Möglichkeiten für die Recordsetermittlung von außen (z. B. falls Filterkritien an eine MSSQL-SErver-Prozedur übergeben werden sollen)
Public Event BeforeCreateSearchResultRecordset(ByVal FilterValue As Variant, ByVal FilterColumns As Variant, ByRef rs As Object)
             ' rs = Rückgabemöglichkeit des zu verwendenden Recordsets

Private Sub Class_Initialize()
   Set m_FilterValues = New Collection
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Public Sub Dispose()
   Set m_Combobox = Nothing
   Set m_DataRecordset = Nothing
   Set m_SearchRecordset = Nothing
   m_BoundColumnIndex = 0
   m_SearchColumnArray = Array()
   m_InSearchMode = 0
End Sub

Public Sub Init(ByVal ComboboxRef As ComboBox, _
       Optional ByVal SearchColumns As Variant, _
       Optional ByVal DataRecordset As Object = Nothing)

   Const EventProcedurePropertyTag As String = "[Event Procedure]"

   Set m_Combobox = ComboboxRef

   With m_Combobox
   
      m_BoundColumnIndex = .BoundColumn - 1
      
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

   If Not IsMissing(SearchColumns) Then
      m_SearchColumnArray = SearchColumns
   End If
   
   If Not DataRecordset Is Nothing Then
      Set m_DataRecordset = DataRecordset
   End If

End Sub

Private Sub m_Combobox_Enter()
   CheckDataSource
End Sub

Private Sub m_Combobox_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case vbKeyReturn, vbKeyTab, vbKeyLeft, vbKeyRight
         'Damit Change-Ereignis auf jeden Fall eintritt
         With m_Combobox
            If .LimitToList = False Or Len(.Text) = 0 Then
               .Text = m_Combobox.Text
               If m_InSearchMode Then
                  KeyCode = 0
               End If
               CheckDropDownStatus
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
      Cancel = True
      CheckDropDownStatus
   End If
End Sub

Private Property Get DataRecordset() As Object

   If m_DataRecordset Is Nothing Then
      RaiseEvent MissingDataRecordset(m_DataRecordset)
      If m_DataRecordset Is Nothing And Len(m_Combobox.RowSource) > 0 Then
         UseRowSourceForDataRecordset
         Exit Sub
      End If
   End If
   
   Set DataRecordset = m_DataRecordset
   
End Property

Private Sub CheckDataSource()
   With m_Combobox
      If Len(.RowSource) = 0 Then
         SetComboboxDataSource DataRecordset
      ElseIf .RowSourceType = "Table/Query" Then
         UseRowSourceForDataRecordset
      End If
   End With
End Sub

Private Sub UseRowSourceForDataRecordset()
   With m_Combobox
      If .Recordset Is Nothing Then
         Set m_DataRecordset = CurrentDb.OpenRecordset(.RowSource)
      Else
         Set m_DataRecordset = .Recordset
      End If
   End With
End Sub

Private Sub CheckDropDownStatus()
   If m_InSearchMode Then
      m_Combobox.Dropdown
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
   
   'Problem bei ADODB-Recordset + RowSourceType = "Table/Query":
   With m_Combobox
      
'      Debug.Assert Nz(.Value = .Column(m_BoundColumnIndex), True)
      If .Value <> .Column(m_BoundColumnIndex) Then
         ' Column(x) zeigt einen Wert, obwohl .Value keinem Wert in der gebundenen Spalte entspricht.
         ' Das tritt an dieser Stelle nur auf, wenn die Combox nicht aufgeklappt war, als Text eingeben wurde.
         ' Aber ist .Column(x) enthält spätestens nach dem Zuklappen der Combobox einen Wert.
         Debug.Print "Value = " & .Value, "Column(" & m_BoundColumnIndex & ") = " & .Column(m_BoundColumnIndex)
         ' Wie kann man das vermeiden? => Problemumgehung: RowSourceType auf "Value List" einstellen.
      End If

   End With
   
   If bolUseDataRecordset Then
      If m_InSearchMode Then
         DeactivateSearchMode
      End If
      Exit Sub
   End If

   ActivateSearchMode FilterValue
  
End Sub

Private Sub ActivateSearchMode(Optional ByVal FilterValue As Variant)

   Dim Cancel As Boolean
   Dim rst As Object 'DAO.Recordset oder ADODB.Recordset

   m_Combobox.Value = Null

   If Not IsArray(m_SearchColumnArray) Then
      m_SearchColumnArray = GetSearchColumnsFromCombobox
   End If

   Set rst = GetSeachResultRecordset(FilterValue, m_SearchColumnArray)
   If rst.EOF Then
      RaiseEvent EmptySearchRecordset(Cancel)
      If Cancel Then
         If m_InSearchMode Then
            Set rst = m_SearchRecordset
         Else
            Set rst = m_DataRecordset
            Set m_SearchRecordset = Nothing
         End If
      End If
   Else
      Set m_SearchRecordset = rst
   End If
   
   m_InSearchMode = True
   SetComboboxDataSource rst
   
End Sub

Private Sub DeactivateSearchMode()
   m_InSearchMode = False
   SetComboboxDataSource m_DataRecordset
   Set m_SearchRecordset = Nothing
   ClearFilterCollection
End Sub

Private Sub SetComboboxDataSource(ByRef rst As Object)

   With m_Combobox

      If .RowSourceType = "Value List" Then
         .RowSource = GetValueListFromRecordset(rst)
      Else
         If .Recordset Is rst Then
            Exit Sub
         ElseIf Not .Recordset Is Nothing Then
            If Not (.Recordset Is m_DataRecordset) Then
               .Recordset.Close
            End If
         End If
         Set .Recordset = rst
      End If
      
   End With
   
End Sub

Private Function GetValueListFromRecordset(ByVal rst As Object) As String
   If Not rst.BOF Then
      rst.MoveFirst
   End If

   If TypeOf rst Is DAO.Recordset Then
      GetValueListFromRecordset = GetValueListFromDaoRecordset(rst)
   Else
      GetValueListFromRecordset = rst.GetString(2, , ";", ";") 'adClipString = 2
   End If

End Function

Private Function GetValueListFromDaoRecordset(ByVal rst As DAO.Recordset) As String
   
   Dim RecordValues() As Variant
   Dim Values() As Variant
   Dim MaxRowIndex As Long
   Dim MaxFieldIndex As Long

   rst.MoveLast
   rst.MoveFirst
   RecordValues = rst.GetRows(rst.RecordCount)

   MaxFieldIndex = UBound(RecordValues, 1)
   MaxRowIndex = UBound(RecordValues, 2)
  
   ReDim Values((MaxFieldIndex + 1) * (MaxRowIndex + 1) - 1)
   
   Dim c As Long, r As Long, i As Long
   For r = 0 To MaxRowIndex
      For c = 0 To MaxFieldIndex
         Values(i) = RecordValues(c, r)
         i = i + 1
      Next
   Next

   GetValueListFromDaoRecordset = Join(Values, ";")

End Function

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
   Set rst = GetSeachResultRecordset(FilterValue, FilterColArray, False)
   If Not rst.EOF Then
      FilterValueInBoundColumnData = True
   End If
   
End Function

Private Function GetSeachResultRecordset(ByVal FilterValue As Variant, ByVal SearchColumnArray As Variant, _
                                Optional ByVal SaveFilterValue As Boolean = True) As Object 'DAO.Recordset
   
   Dim rst As Object 'DAO.Recordset oder ADODB.Recordset
   Dim FilterString As String
   Dim i As Long
   
   If SaveFilterValue Then
      AddFilterValueToFilterCollection FilterValue
   End If
   
   RaiseEvent BeforeCreateSearchResultRecordset(FilterValue, SearchColumnArray, rst)
   If Not (rst Is Nothing) Then
   ' Abbruch, da rst von außen übergeben wurde
      Set GetSeachResultRecordset = rst
      Exit Function
   End If

   If m_SearchRecordset Is Nothing Then
      Set m_SearchRecordset = DataRecordset.Clone
   End If
   Set rst = m_SearchRecordset
   If Not rst.BOF Then
      rst.MoveFirst
   End If

   If IsNull(FilterValue) Then
      Set GetSeachResultRecordset = rst
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
   
   Set GetSeachResultRecordset = rst

End Function

Private Sub AddFilterValueToFilterCollection(ByVal FilterValue As Variant)
   m_FilterValues.Add FilterValue
End Sub

Private Sub ClearFilterCollection()
   Set m_FilterValues = Nothing
   Set m_FilterValues = New Collection
End Sub

Private Function GetSearchColumnsFromCombobox() As Variant()
   
   Dim SearchColumns() As Variant
   Dim ColumnWidthArray() As String
   Dim i As Long
   Dim c As Long
   
   With m_Combobox
      ColumnWidthArray = Split(.ColumnWidths, ";")
      ReDim SearchColumns(UBound(ColumnWidthArray))
      
      For i = 0 To UBound(ColumnWidthArray)
         If Val(ColumnWidthArray(i)) > 0 Then
            SearchColumns(c) = i
            c = c + 1
         End If
      Next
      
      ReDim Preserve SearchColumns(c - 1)
   
   End With
   
   GetSearchColumnsFromCombobox = SearchColumns
   
End Function

Public Property Get FilterValues() As Collection
   Set FilterValues = m_FilterValues
End Property
