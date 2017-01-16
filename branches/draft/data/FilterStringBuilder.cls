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
'  <test>_test/data/FilterStringBuilderTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private m_Items As StringCollection
Private m_SqlTool As SqlTools
Private m_DefaultConcatOperator As SqlLogicalOperator

Private m_ConditionGroups As Collection

Private m_SqlPrefix As String
Private m_SqlSuffix As String
Private m_ConditionPrefix As String
Private m_IgnoreSqlPreSuffixIfEmptyFilter As Boolean

Private Const WhereReplacementText As String = "[WhereStatement]"

Private Sub Class_Initialize()
   Set m_Items = New StringCollection
   m_DefaultConcatOperator = SqlLogicalOperator.SQL_And
End Sub

Private Sub Class_Terminate()
   Set m_Items = Nothing
End Sub

Public Property Get Self() As FilterStringBuilder
   Set Self = Me
End Property

Public Property Get DefaultConcatOperator() As SqlLogicalOperator
   DefaultConcatOperator = m_DefaultConcatOperator
End Property

Public Property Let DefaultConcatOperator(ByVal NewValue As SqlLogicalOperator)
   m_DefaultConcatOperator = NewValue
End Property

Public Sub Add(ByVal FieldName As String, ByVal FieldDataType As SqlFieldDataType, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal Value As Variant, _
      Optional ByVal Value2 As Variant = Null, _
      Optional ByVal IgnoreValue As Variant = Null)

   AddCriteria SqlTool.BuildCriteria(FieldName, FieldDataType, RelationalOperator, Value, Value2, IgnoreValue)

End Sub

Public Sub AddCriteria(ByVal Criteria As String)
   If Len(Criteria) = 0 Then Exit Sub
   m_Items.Add Criteria
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

Public Function ToString(Optional ByVal ConcatOperator As SqlLogicalOperator = SqlLogicalOperator.[_SQL_Default], _
                         Optional ByVal IgnoreDuplicateFilters As Boolean = False, _
                         Optional ByVal UseConditionBracket As Boolean = False) As String
   
   Dim FilterString As String
   Dim ItemPrefix As String
   Dim ItemSuffix As String

   If ConcatOperator = SqlLogicalOperator.[_SQL_Default] Then
      ConcatOperator = DefaultConcatOperator
   End If
   
   If UseConditionBracket Then
      ItemPrefix = "("
      ItemSuffix = ")"
   End If
   
   FilterString = m_Items.ToString(GetConcatOperatorString(ConcatOperator), ItemPrefix, ItemSuffix, , IgnoreDuplicateFilters)
   FilterString = AppendFilterGroupsString(FilterString, ConcatOperator, IgnoreDuplicateFilters, UseConditionBracket)
   If Len(FilterString) > 0 Then
      FilterString = m_ConditionPrefix & FilterString
   End If
   
   If m_IgnoreSqlPreSuffixIfEmptyFilter Then
   If Len(FilterString) = 0 Then
      ToString = vbNullString
      Exit Function
   End If
   End If
   
   ToString = m_SqlPrefix & FilterString & m_SqlSuffix
   
End Function

Private Function GetConcatOperatorString(ByVal ConcatOperator As SqlLogicalOperator) As String

   Select Case ConcatOperator
      Case SqlLogicalOperator.SQL_And
         GetConcatOperatorString = " And "
      Case SqlLogicalOperator.SQL_Or
         GetConcatOperatorString = " Or "
      Case SqlLogicalOperator.SQL_CommaSeparator
         GetConcatOperatorString = ", "
      Case Else
         
   End Select
 
End Function

Private Property Get ConditionGroups() As Collection
   If m_ConditionGroups Is Nothing Then
      Set m_ConditionGroups = New Collection
   End If
   Set ConditionGroups = m_ConditionGroups
End Property

Public Function NewConditionGroup(ByVal ConcatOperator As SqlLogicalOperator) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder
   
   Set NewBuilder = New FilterStringBuilder
   Set NewBuilder.SqlTool = m_SqlTool
   NewBuilder.DefaultConcatOperator = ConcatOperator
   
   ConditionGroups.Add NewBuilder
   
   Set NewConditionGroup = NewBuilder
   
End Function

Private Function AppendFilterGroupsString(ByVal BaseFilterString As String, ByVal ConcatOperator As SqlLogicalOperator, _
                                 Optional ByVal IgnoreDuplicateFilters As Boolean = False, _
                                 Optional ByVal UseConditionBracket As Boolean = False) As String
   
   Dim CondPrefix As String
   Dim CondSuffix As String
   Dim Cond As String
   
   If m_ConditionGroups Is Nothing Then
      AppendFilterGroupsString = BaseFilterString
      Exit Function
   End If
   
   If m_ConditionGroups.Count = 0 Then
      AppendFilterGroupsString = BaseFilterString
      Exit Function
   End If
   
   Dim FSB As FilterStringBuilder
   
   With New StringCollection
   
      .Add BaseFilterString
      
      If Not UseConditionBracket Then ' Gruppen immer Klammern
         CondPrefix = "("
         CondSuffix = ")"
      End If
   
      For Each FSB In m_ConditionGroups
'         Cond = FSB.ToString(, , UseConditionBracket)
'         If UseConditionBracket Then
'            If Len(Cond) > 0 Then
'               Cond = CondPrefix & Cond & CondSuffix
'            End If
'         End If
         .Add FSB.ToString(, , UseConditionBracket)
      Next

      AppendFilterGroupsString = .ToString(GetConcatOperatorString(ConcatOperator), "(", ")", True, IgnoreDuplicateFilters)
      
   End With
   
End Function

Public Function AddSubSelectCriteria( _
               ByVal FieldName As String, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfSubSelectHasNoCriteria As Boolean = False, _
               Optional ByVal UseWhereReplacementTextInFromText As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And _
      ) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder
   Dim SqlPrefix As String
   Dim SqlSuffix As String
   Dim WhereReplacementPos As Long
   
   SqlPrefix = FieldName & " " & SqlTools.GetRelationalOperatorString(RelationalOperator) & " ("
   SqlSuffix = ")"
   If UseWhereReplacementTextInFromText Then
      WhereReplacementPos = InStr(1, SelectFromText, WhereReplacementText, vbTextCompare)
   End If
   
   If WhereReplacementPos > 0 Then
      SqlPrefix = SqlPrefix & Trim$(Left(SelectFromText, WhereReplacementPos - 1))
      SqlSuffix = " " & Trim$(Mid$(SelectFromText, WhereReplacementPos + Len(WhereReplacementText))) & SqlSuffix
   Else
      SqlPrefix = SqlPrefix & SelectFromText
   End If
   
   Set NewBuilder = New FilterStringBuilder
   NewBuilder.DefaultConcatOperator = SubSelectConcatOperator
   Set NewBuilder.SqlTool = m_SqlTool
   
   NewBuilder.ConfigSqlStatement SqlPrefix:=SqlPrefix, SqlSuffix:=SqlSuffix, _
                                 ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfSubSelectHasNoCriteria
   ConditionGroups.Add NewBuilder
   
   Set AddSubSelectCriteria = NewBuilder
   
End Function

Public Function AddExistsCriteria( _
                        ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfExistsStatementHasNoCriteria As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And, _
               Optional ByVal UseNotExists As Boolean = False _
      ) As FilterStringBuilder

   Dim NewBuilder As FilterStringBuilder
   Dim ExistsSqlPrefix As String
   
   Set NewBuilder = New FilterStringBuilder
   NewBuilder.DefaultConcatOperator = SubSelectConcatOperator
   Set NewBuilder.SqlTool = m_SqlTool
   
   ExistsSqlPrefix = "Exists ("
   If UseNotExists Then ExistsSqlPrefix = "Not " & ExistsSqlPrefix
   
   NewBuilder.ConfigSqlStatement SqlPrefix:=ExistsSqlPrefix & SelectFromText, _
                                 SqlSuffix:=")", ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfExistsStatementHasNoCriteria
   
   ConditionGroups.Add NewBuilder
   
   Set AddExistsCriteria = NewBuilder
   
End Function

Friend Sub ConfigSqlStatement(ByVal SqlPrefix As String, ByVal SqlSuffix As String, _
                              ByVal ConditionPrefix As String, _
                     Optional ByVal IgnoreSqlPreSuffixIfEmptyFilter As Boolean = False)
   m_SqlPrefix = SqlPrefix
   m_SqlSuffix = SqlSuffix
   m_ConditionPrefix = ConditionPrefix
   m_IgnoreSqlPreSuffixIfEmptyFilter = IgnoreSqlPreSuffixIfEmptyFilter

End Sub
