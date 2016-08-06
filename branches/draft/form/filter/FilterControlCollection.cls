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
'  <use>form/filter/FilterControlEventBridge.cls</use>
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_FilterControls As Collection
Private WithEvents m_EventBridge As FilterControlEventBridge
Attribute m_EventBridge.VB_VarHelpID = -1
Private m_ItemConcatOperator As SqlLogicalOperator

Private m_SqlPrefix As String
Private m_SqlSuffix As String
Private m_ConditionPrefix As String
Private m_IgnoreSqlPreSuffixIfEmptyFilter As Boolean

Private Const WhereReplacementText As String = "[WhereStatement]"

Public Event FilterValuesChanged()
Public Event FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)

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
      .DataField = DataFieldName
      .DataType = DataType
      .RelationalOperator = RelationalOperator
      .IgnoreValue = IgnoreValue
      .SetSubControls Controls
      m_FilterControls.Add .Self
   End With
   
End Sub

Public Property Get Items() As Collection
   Set Items = m_FilterControls
End Property

Public Sub RemoveFilterValues(Optional ByVal RaiseFilterValuesChangedEvent As Boolean = True)
   
   Dim fc As Object ' FilterControl oder FilterControlCollection
   For Each fc In Items
      If TypeOf fc Is FilterControlCollection Then
         fc.RemoveFilterValues False
      Else
         fc.RemoveFilterValue False
      End If
   Next
   
   If RaiseFilterValuesChangedEvent Then
      RaiseEvent FilterValuesChanged
   End If
   
End Sub

Private Sub m_EventBridge_FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)
   RaiseEvent FilterValueRequest(FilterValueCode, FilterValue)
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
      .ConfigSqlStatement InsertEventValues(m_SqlPrefix), InsertEventValues(m_SqlSuffix), m_ConditionPrefix, m_IgnoreSqlPreSuffixIfEmptyFilter
      For Each fc In m_FilterControls
         .AddCriteria fc.CriteriaString(SqlToolsRef)
      Next
      CriteriaString = .ToString(m_ItemConcatOperator)
   End With

End Function

Private Function InsertEventValues(ByVal StringToCheck As String) As String

   Dim CheckStringArray() As String
   
   If Len(StringToCheck) = 0 Then
      Exit Function
   End If
   
   CheckStringArray = Split(StringToCheck, "[EventValue(", , vbTextCompare)
   
   Dim i As Long
   Dim ArraySize As Long
   
   ArraySize = UBound(CheckStringArray)
   
   If ArraySize = 0 Then
      InsertEventValues = StringToCheck
      Exit Function
   End If
   
   Dim FilterValueCode As String
   Dim FilterValue As String
   Dim Pos As String
   Dim TempString As String
   
   For i = 1 To ArraySize
   
      Pos = InStr(1, CheckStringArray(i), ")")
      FilterValueCode = Left(CheckStringArray(i), Pos - 1)
      
      Pos = InStr(Pos, CheckStringArray(i), "]")
      
      TempString = Mid(CheckStringArray(i), Pos + 1)
      
      FilterValue = "Null"
      RaiseEvent FilterValueRequest(FilterValueCode, FilterValue)
      If FilterValue = "Null" Then
         m_EventBridge.RaiseFilterValueRequest FilterValueCode, FilterValue
      End If
      
      CheckStringArray(i) = FilterValue & TempString
      
   Next
   
   InsertEventValues = Join(CheckStringArray, vbNullString)

End Function

Public Function AddSubSelectCriteria( _
               ByVal FieldName As String, _
               ByVal RelationalOperator As SqlRelationalOperators, _
               ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfSubSelectHasNoCriteria As Boolean = False, _
               Optional ByVal UseWhereReplacementTextInFromText As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And _
      ) As FilterControlCollection
      
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
      SqlSuffix = " " & Trim$(Mid$(SelectFromText, WhereReplacementPos + Len(WhereReplacementText) + 1)) & SqlSuffix
   Else
      SqlPrefix = SqlPrefix & SelectFromText
   End If
   
   With New FilterControlCollection
      .ItemConcatOperator = SubSelectConcatOperator
      .ConfigSqlStatement SqlPrefix:=SqlPrefix, SqlSuffix:=SqlSuffix, _
                          ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfSubSelectHasNoCriteria
      
      Set .EventBridge = m_EventBridge
      m_FilterControls.Add .Self
      Set AddSubSelectCriteria = .Self
   End With
   
End Function

Public Function AddExistsCriteria( _
                        ByVal SelectFromText As String, _
               Optional ByVal IgnoreIfExistsStatementHasNoCriteria As Boolean = False, _
               Optional ByVal SubSelectConcatOperator As SqlLogicalOperator = SqlLogicalOperator.SQL_And, _
               Optional ByVal UseNotExists As Boolean = False _
      ) As FilterControlCollection

   Dim ExistsSqlPrefix As String
   
   ExistsSqlPrefix = "Exists ("
   If UseNotExists Then ExistsSqlPrefix = "Not " & ExistsSqlPrefix

   With New FilterControlCollection
      .ItemConcatOperator = SubSelectConcatOperator
      .ConfigSqlStatement SqlPrefix:=ExistsSqlPrefix & SelectFromText, _
                          SqlSuffix:=")", ConditionPrefix:=" Where ", IgnoreSqlPreSuffixIfEmptyFilter:=IgnoreIfExistsStatementHasNoCriteria
      
      Set .EventBridge = m_EventBridge
      m_FilterControls.Add .Self
      Set AddExistsCriteria = .Self
   End With

End Function

Friend Sub ConfigSqlStatement(ByVal SqlPrefix As String, ByVal SqlSuffix As String, _
                              ByVal ConditionPrefix As String, _
                     Optional ByVal IgnoreSqlPreSuffixIfEmptyFilter As Boolean = False)
   
   m_SqlPrefix = SqlPrefix
   m_SqlSuffix = SqlSuffix
   m_ConditionPrefix = ConditionPrefix
   m_IgnoreSqlPreSuffixIfEmptyFilter = IgnoreSqlPreSuffixIfEmptyFilter

End Sub
