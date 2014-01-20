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
'  <use>data/SqlTools.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
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

Public Enum SqlFieldDataType
   SQL_Boolean = 1
   SQL_Numeric = 2
   SQL_Text = 3
   SQL_Date = 4
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

Public Sub Add(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
               ByVal Value As Variant, _
      Optional ByVal RelationalOperator As SqlRelationalOperators = SQL_Equal, _
      Optional ByVal Value2 As Variant = Null, _
      Optional ByVal IgnoreValue As Variant = Null)

   If NullFilterOrEmptyFilterAdded(FieldName, Nz(Value, Value2), IgnoreValue) Then
      Exit Sub
   End If

   m_Items.Add "(" & BuildCriteria(FieldName, FieldDataType, RelationalOperator, Value, Value2) & ")"

End Sub

Private Function BuildCriteria(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
                               ByVal RelationalOperator As SqlRelationalOperators, _
                               ByVal FilterValue As Variant, _
                      Optional ByVal FilterValue2 As Variant = Null) As String

   Dim FilterValueString As String
   Dim FilterValue2String As String
   Dim OperatorString As String
   Dim Criteria As String
   
   If (RelationalOperator And SQL_In) = SQL_In Then
      If IsArray(FilterValue) Then
         BuildCriteria = FieldName & " In (" & GetValueArrayString(FilterValue, FieldDataType, ",") & ")"
      ElseIf VarType(FilterValue) = vbString Then ' Value ist bereits die Auflistung als String
         BuildCriteria = FieldName & " In (" & FilterValue & ")"
      Else
         BuildCriteria = FieldName & " In (" & GetFilterValueString(FilterValue, FieldDataType) & ")"
      End If
      Exit Function
   End If

   FilterValueString = GetFilterValueString(FilterValue, FieldDataType)
   FilterValue2String = GetFilterValueString(FilterValue2, FieldDataType)
      

   If (RelationalOperator And SQL_Between) = SQL_Between Then
      If IsNull(FilterValue2) Or IsMissing(FilterValue2) Then
         RelationalOperator = SQL_GreaterThan + SQL_Equal
      ElseIf IsNull(FilterValue) Then
         RelationalOperator = SQL_LessThan + SQL_Equal
         FilterValueString = FilterValue2String
      Else
         BuildCriteria = FieldName & " Between " & FilterValueString & " And " & FilterValue2String
         Exit Function
      End If
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

Private Function GetValueArrayString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType, ByVal Delimiter As String) As String
   
   Dim i As Long

   With New StringCollection
      For i = LBound(Value) To UBound(Value)
         .Add GetFilterValueString(Value(i), FieldDataType)
      Next
      GetValueArrayString = .ToString(Delimiter)
   End With

End Function

Private Function GetFilterValueString(ByVal Value As Variant, ByVal FieldDataType As SqlFieldDataType) As String

   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Numeric
         GetFilterValueString = SqlTools.NumberToSqlText(Value)
      Case SqlFieldDataType.SQL_Text
         GetFilterValueString = SqlTools.TextToSqlText(Value)
      Case SqlFieldDataType.SQL_Date
         GetFilterValueString = SqlTools.DateToSqlText(Value)
      Case SqlFieldDataType.SQL_Boolean
         GetFilterValueString = SqlTools.BooleanToSqlText(Value)
      Case Else
         Err.Raise vbObjectError, "FilterStringBuilder.GetFilterValueString", "SqlFieldDataType '" & FieldDataType & "' wird nicht unterstützt."

   End Select
End Function

Private Function NullFilterOrEmptyFilterAdded(ByVal FieldName As String, ByVal Value As Variant, ByVal IgnoreValue As Variant) As Boolean
   
   If IsNull(Value) Then
      If Not IsNull(IgnoreValue) Then
         m_Items.Add FieldName & " Is Null"
      End If
      NullFilterOrEmptyFilterAdded = True
   ElseIf IsArray(Value) Then
      Dim a() As Variant
      a = Value
      If (0 / 1) + (Not Not a) = 0 Then ' leerer Array
         NullFilterOrEmptyFilterAdded = True
      End If
   ElseIf Value = IgnoreValue Then
      NullFilterOrEmptyFilterAdded = True
   End If

End Function

Public Function ToString(Optional ByVal Delimiter As String = " AND ") As String
   ToString = m_Items.ToString(Delimiter)
End Function
