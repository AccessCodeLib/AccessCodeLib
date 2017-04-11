VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlTagConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlTagConverter
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Hilfsklasse zum Befüllen der FilterControls (für den FilterControlManager) über Control-Tag-Werte
' </summary>
' <remarks></remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControlTagConverter.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControlCollection.cls</use>
'  <use>data/SqlTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const DefaultFilterControlNamePrefix As String = "fctl"
Private Const DefaultFilterControlTagValueSeparator As String = "|"

Private Const DEF_DefaultNumericRelationalOperator As Long = SqlRelationalOperators.SQL_Equal
Private Const DEF_DefaultTextRelationalOperator As Long = SqlRelationalOperators.SQL_Like + SqlRelationalOperators.SQL_Add_WildCardSuffix
Private Const DEF_DefaultDateRelationalOperator As Long = SqlRelationalOperators.SQL_Equal
Private Const DEF_DefaultBooleanRelationalOperator As Long = SqlRelationalOperators.SQL_Equal

Private m_FilterControlNamePrefix As String
Private m_FilterControlTagValueSeparator As String

Private m_DefaultNumericRelationalOperator As SqlRelationalOperators
Private m_DefaultTextRelationalOperator As SqlRelationalOperators
Private m_DefaultDateRelationalOperator As SqlRelationalOperators
Private m_DefaultBooleanRelationalOperator As SqlRelationalOperators

Private m_DataSourceRecordset As Object ' wegen DAO und ADODB

Private Type FctlDef
   DataFieldName As String
   DataType As SqlFieldDataType
   RelationalOperator As SqlRelationalOperators
   IgnoreValue As Variant
   FilterControl2Name As String
End Type

Private Sub Class_Initialize()

   m_FilterControlNamePrefix = DefaultFilterControlNamePrefix
   m_FilterControlTagValueSeparator = DefaultFilterControlTagValueSeparator
   
   m_DefaultNumericRelationalOperator = DEF_DefaultNumericRelationalOperator
   m_DefaultTextRelationalOperator = DEF_DefaultTextRelationalOperator
   m_DefaultDateRelationalOperator = DEF_DefaultDateRelationalOperator
   m_DefaultBooleanRelationalOperator = DEF_DefaultBooleanRelationalOperator
   
End Sub

Public Sub AddFilterControls(ByVal FctlCol As FilterControlCollection, ByVal ControlsRef As Object, Optional ByVal DataSourceRecordset As Object = Nothing)
   
   Dim ctl As Control
   
   Set m_DataSourceRecordset = DataSourceRecordset
   
   For Each ctl In ControlsRef
      If ctl.Name Like m_FilterControlNamePrefix & "*" Then
         AddFilterControl FctlCol, ctl
      End If
   Next
   
   Set m_DataSourceRecordset = Nothing
   
End Sub

Private Sub AddFilterControl(ByVal FctlCol As FilterControlCollection, ByVal FilterControlRef As Control)
   
   Dim FcDef As FctlDef
   Dim FilterControl2 As Control
   
   FcDef = GetFilterControlDefinition(FilterControlRef)
   
   With FcDef
      
      If Len(.FilterControl2Name) > 0 Then
         Set FilterControl2 = FilterControlRef.Parent.Form.Controls(.FilterControl2Name)
      Else
         Set FilterControl2 = Nothing
      End If
      
      FctlCol.Add .DataFieldName, .DataType, .RelationalOperator, FilterControlRef, FilterControl2, .IgnoreValue
      
   End With
   
End Sub

Private Function GetFilterControlDefinition(ByVal FilterControlRef As Control) As FctlDef
   
   Dim FcDef As FctlDef
   
   Dim TagValues() As String
   
   If Len(FilterControlRef.Tag) > 0 Then
      TagValues = Split(FilterControlRef.Tag, m_FilterControlTagValueSeparator)
      FillFilterControlDefinitionFromTagArray FcDef, TagValues
   End If
   
   If Len(FcDef.DataFieldName) = 0 Then
      FcDef.DataFieldName = Mid(FilterControlRef.Name, Len(m_FilterControlNamePrefix) + 1)
      If Len(FcDef.DataFieldName) = 0 Then
         Err.Raise vbObjectError, "FilterControlTagConverter.GetFilterControlDefinition", "DataFieldName is not available"
      End If
   End If
      
   If FcDef.DataType = 0 Then
      If Not TryGetDataTypeFromDataSource(FcDef.DataFieldName, FcDef.DataType) Then
         Err.Raise vbObjectError, "FilterControlTagConverter.GetFilterControlDefinition", "DataType is not available"
      End If
   End If
   
   If FcDef.RelationalOperator = 0 Then
      FcDef.RelationalOperator = GetDefaultRelationalOperator(FcDef.DataType)
   End If
   
   If IsEmpty(FcDef.IgnoreValue) Then
      FcDef.IgnoreValue = Null
   End If
   
   GetFilterControlDefinition = FcDef
   
End Function

Private Function TryGetDataTypeFromDataSource(ByVal FieldName As String, ByRef FieldDataType As SqlFieldDataType) As Boolean
   
   If m_DataSourceRecordset Is Nothing Then
      TryGetDataTypeFromDataSource = False
      Exit Function
   End If
   
   Err.Raise vbObjectError, "TryGetDataTypeFromDataSource", "Not implemented yet"
   
End Function

Private Function GetDefaultRelationalOperator(ByVal FieldDataType As SqlFieldDataType) As SqlRelationalOperators

   Select Case FieldDataType
      Case SqlFieldDataType.SQL_Numeric
         GetDefaultRelationalOperator = m_DefaultNumericRelationalOperator
      Case SqlFieldDataType.SQL_Text
         GetDefaultRelationalOperator = m_DefaultTextRelationalOperator
      Case SqlFieldDataType.SQL_Date
         GetDefaultRelationalOperator = m_DefaultDateRelationalOperator
      Case SqlFieldDataType.SQL_Boolean
         GetDefaultRelationalOperator = m_DefaultBooleanRelationalOperator
      Case Else
          Err.Raise vbObjectError, "FilterControlTagConverter.GetFilterControlDefinition", "DataType '" & FieldDataType & "' is not supported"
   End Select
   
End Function

Private Sub FillFilterControlDefinitionFromTagArray(ByRef FcDef As FctlDef, ByRef TagValues() As String)
' <DataType>|<RelationalOperator>|<FieldName>|<Control2Name>|<IgnoreValue>
'
' Samples:
'    D|between|InvoiceDate|xfctlInvoiceDate_Max
'    N|>=
'    T|*like*
'    T
   
   Dim MaxArrayIndex As Long
   
   MaxArrayIndex = UBound(TagValues)
   
   If Len(TagValues(0)) > 0 Then
      SetDataTypeFromTagValue FcDef.DataType, TagValues(0)
   End If
   If MaxArrayIndex = 0 Then Exit Sub
   
   If Len(TagValues(1)) > 0 Then
      SetRelationalOperatorFromTagValue FcDef.RelationalOperator, TagValues(1)
   End If
   If MaxArrayIndex = 1 Then Exit Sub
   
   If Len(TagValues(2)) > 0 Then
      FcDef.DataFieldName = TagValues(2)
   End If
   If MaxArrayIndex = 2 Then Exit Sub
   
   If Len(TagValues(3)) > 0 Then
      FcDef.FilterControl2Name = TagValues(3)
   End If
   If MaxArrayIndex = 3 Then Exit Sub
   
   Err.Raise vbObjectError, "FillFilterControlDefinitionFromTagArray", "NullValue is not implemented yet"
   
End Sub

Private Function SetDataTypeFromTagValue(ByRef FcDefDataType As SqlFieldDataType, ByVal TagValue As String) As Boolean
   
   SetDataTypeFromTagValue = True

   Select Case TagValue
      Case "N", "Numeric"
         FcDefDataType = SQL_Numeric
      Case "T", "Text"
         FcDefDataType = SQL_Text
      Case "D", "Date"
         FcDefDataType = SQL_Date
      Case "B", "Boolean", "Bool"
         FcDefDataType = SQL_Boolean
      Case Else
         SetDataTypeFromTagValue = False
   End Select
   
End Function

Private Function SetRelationalOperatorFromTagValue(ByRef FcDefRelationalOperator As SqlRelationalOperators, ByVal TagValue As String) As Boolean

   Dim i As Long

   SetRelationalOperatorFromTagValue = True
   
   If TagValue Like "Not*" Then
      FcDefRelationalOperator = SQL_Not
      TagValue = Mid(TagValue, Len("NOT") + 1)
   End If
   
   If TagValue Like "*between*" Then
      FcDefRelationalOperator = FcDefRelationalOperator + SQL_Between
      If TagValue Like "*between[*]*" Then ' Sonderfall für Datumswerte
         FcDefRelationalOperator = FcDefRelationalOperator + SQL_Add_WildCardSuffix
      End If
      Exit Function
   ElseIf TagValue Like "*like*" Then
      FcDefRelationalOperator = FcDefRelationalOperator + SQL_Like
      If TagValue Like "*like[*]*" Then
         FcDefRelationalOperator = FcDefRelationalOperator + SQL_Add_WildCardSuffix
      End If
      If TagValue Like "*[*]like*" Then
         FcDefRelationalOperator = FcDefRelationalOperator + SQL_Add_WildCardPrefix
      End If
      Exit Function
   ElseIf TagValue Like "*In*" Then
      FcDefRelationalOperator = FcDefRelationalOperator + SQL_In
      Exit Function
   End If
   
   For i = 1 To Len(TagValue)
      Select Case Mid(TagValue, i, 1)
         Case "="
            FcDefRelationalOperator = FcDefRelationalOperator Or SQL_Equal
         Case "<"
            FcDefRelationalOperator = FcDefRelationalOperator Or SQL_LessThan
         Case ">"
            FcDefRelationalOperator = FcDefRelationalOperator Or SQL_GreaterThan
      End Select
   Next
   
   If FcDefRelationalOperator = 0 Then
      SetRelationalOperatorFromTagValue = False
   End If

End Function
