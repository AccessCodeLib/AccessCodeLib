VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterStringBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterStringBuilder
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' SQL-Filter-Ausdruck erstellen
' </summary>
' <remarks></remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/FilterStringBuilder.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>text/StringCollection.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_Items As StringCollection

Public Enum SqlRelationalOperators
   SQL_Not = 1
   SQL_Equal = 2
   SQL_LessThan = 4
   SQL_GreaterThan = 8
   SQL_Like = 256
   SQL_Between = 512
   SQL_In = 1024
End Enum

Private Sub Class_Initialize()
   Set m_Items = New StringCollection
End Sub

Private Sub Class_Terminate()
   Set m_Items = Nothing
End Sub

Public Property Get Self() As FilterStringBuilder
   Set Self = Me
End Property

Public Sub AddTextFilter(ByVal FieldName As String, ByVal Value As Variant, _
                Optional ByVal RelationalOperator As SqlRelationalOperators = SQL_Equal, _
                Optional ByVal Value2 As Variant = Null, _
                Optional ByVal IgnoreValue As Variant = Null)

   If NullFilterOrEmptyFilterAdded(FieldName, Value, IgnoreValue) Then
      Exit Sub
   End If

   If Len(Value) > 0 Then
      AddFilterItem FieldName, RelationalOperator, SqlTools.TextToSqlText(Value), SqlTools.TextToSqlText(Value2)
   End If

End Sub

Public Sub AddNumericFilter(ByVal FieldName As String, ByVal Value As Variant, _
                   Optional ByVal RelationalOperator As SqlRelationalOperators = SQL_Equal, _
                   Optional ByVal Value2 As Variant = Null, _
                   Optional ByVal IgnoreValue As Variant = Null)

   If NullFilterOrEmptyFilterAdded(FieldName, Value, IgnoreValue) Then
      Exit Sub
   End If

   AddFilterItem FieldName, RelationalOperator, SqlTools.NumberToSqlText(Value), SqlTools.NumberToSqlText(Value2)

End Sub

Public Sub AddDateFilter(ByVal FieldName As String, ByVal Value As Variant, _
                Optional ByVal RelationalOperator As SqlRelationalOperators = SQL_Equal, _
                Optional ByVal Value2 As Variant = Null, _
                Optional ByVal IgnoreValue As Variant = Null)
   
   If NullFilterOrEmptyFilterAdded(FieldName, Value, IgnoreValue) Then
      Exit Sub
   End If

   If Len(Value) > 0 Then
      AddFilterItem FieldName, RelationalOperator, SqlTools.DateToSqlText(Value), SqlTools.DateToSqlText(Value2)
   End If

End Sub

Private Sub AddFilterItem(ByVal FieldName As String, _
                          ByVal RelationalOperator As SqlRelationalOperators, _
                          ByVal FilterValueString As String, _
                          ByVal FilterValue2String As String)
   m_Items.Add BuildCriteria(FieldName, RelationalOperator, FilterValueString, FilterValue2String)
End Sub

Private Function BuildCriteria(ByVal FieldName As String, _
                               ByVal RelationalOperator As SqlRelationalOperators, _
                               ByVal FilterValueString As String, _
                      Optional ByVal FilterValue2String As String = vbNullString) As String

   Dim OperatorString As String
   Dim Criteria As String

   If (RelationalOperator And SQL_In) = SQL_In Then
      BuildCriteria = FieldName & " In (" & FilterValueString & ")"
      Exit Function
   End If

   If (RelationalOperator And SQL_Between) = SQL_Between Then
      BuildCriteria = FieldName & " Between " & FilterValueString & " And " & FilterValue2String
      Exit Function
   End If

   If (RelationalOperator And SQL_Like) = SQL_Like Then
      BuildCriteria = FieldName & " like " & FilterValueString
      Exit Function
   End If
   

   If (RelationalOperator And SQL_LessThan) = SQL_LessThan Then
      OperatorString = OperatorString & "<"
   End If
   
   If (RelationalOperator And SQL_GreaterThan) = SQL_GreaterThan Then
      OperatorString = OperatorString & ">"
   End If

   If (RelationalOperator And SQL_Equal) = SQL_Equal Then
      OperatorString = OperatorString & "="
   End If

   Criteria = FieldName & " " & OperatorString & " " & FilterValueString
   If (RelationalOperator And SQL_Not) = SQL_Not Then
      Criteria = "Not " & Criteria
   End If

   BuildCriteria = Criteria

End Function


Private Function NullFilterOrEmptyFilterAdded(ByVal FieldName As String, ByVal Value As Variant, ByVal IgnoreValue As Variant) As Boolean
   
   If IsNull(Value) Then
      If Not IsNull(IgnoreValue) Then
         m_Items.Add FieldName & " Is Null"
      End If
      NullFilterOrEmptyFilterAdded = True
   ElseIf Value = IgnoreValue Then
      NullFilterOrEmptyFilterAdded = True
   End If

End Function

Public Sub AddFilterFromListboxSelectedItems( _
                 ByVal FieldName As String, ByVal lst As ListBox, ByVal ColumnIndexToAdd As Long, _
        Optional ByVal FieldValuePrefix As String = vbNullString)
   
   Dim FilterValues As String

   If lst.ItemsSelected.Count = 0 Then
      Exit Sub
   End If

   With New StringCollection
      .AddFromArray GetSelectedListboxItemArray(lst, ColumnIndexToAdd)
      FilterValues = .ToString(", ", FieldValuePrefix, FieldValuePrefix)
   End With

   If Len(FilterValues) = 0 Then
      Exit Sub
   End If

   m_Items.Add FieldName & " In (" & FilterValues & ")"

End Sub

Public Function ToString(Optional ByVal Delimiter As String = " AND ") As String
   ToString = m_Items.ToString(Delimiter)
End Function
