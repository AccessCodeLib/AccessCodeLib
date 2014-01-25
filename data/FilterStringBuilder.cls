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
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_Items As StringCollection
Private m_SqlTool As SqlTools

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

   AddCriteria SqlTool.BuildCriteria(FieldName, FieldDataType, RelationalOperator, Value, Value2, IgnoreValue)

End Sub

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


Public Sub AddCriteria(ByVal Criteria As String)
   If Len(Criteria) = 0 Then Exit Sub
   m_Items.Add "(" & Criteria & ")"
End Sub

Public Function ToString(Optional ByVal Delimiter As String = " AND ") As String
   ToString = m_Items.ToString(Delimiter)
End Function
