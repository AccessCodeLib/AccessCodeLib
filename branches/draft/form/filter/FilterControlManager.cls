VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlManager
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Filterstring mit Formular-Steuerelementen erstellen
' </summary>
' <remarks></remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControlManager.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControlCollection.cls</use>
'  <use>data/FilterStringBuilder.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_SqlTool As SqlTools
Private WithEvents m_FilterControls As FilterControlCollection
Attribute m_FilterControls.VB_VarHelpID = -1
Private m_AutoFilterModeOn As Boolean

Public Event FilterStringChanged(ByVal NewFilterString As String)
Public Event FilterValuesChanged()
Public Event FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)

Private Sub Class_Initialize()
   Set m_FilterControls = New FilterControlCollection
End Sub

Private Sub Class_Terminate()
   Set m_FilterControls = Nothing
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

Public Property Get FilterControls() As FilterControlCollection
   Set FilterControls = m_FilterControls
End Property

Public Function CreateFilterString() As String
   
   CreateFilterString = FilterControls.CriteriaString(SqlTool)

End Function

Private Sub m_FilterControls_FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)
   RaiseEvent FilterValueRequest(FilterValueCode, FilterValue)
End Sub

Private Sub m_FilterControls_FilterValuesChanged()
   RaiseEvent FilterValuesChanged
   UseAutoFilterMode
End Sub

Public Property Get AutoFilterOn() As Boolean
   AutoFilterOn = m_AutoFilterModeOn
End Property

Public Property Let AutoFilterOn(ByVal NewValue As Boolean)
   m_AutoFilterModeOn = NewValue
   UseAutoFilterMode
End Property

Public Sub UseAutoFilterMode()
   If m_AutoFilterModeOn Then
      ApplyFilter
   End If
End Sub

Public Sub UseFilter()
' DEPRECATED --> New: ApplyFilter
   Stop
   ApplyFilter
End Sub

Public Sub ApplyFilter()
   
   Dim NewFilterString As String
   NewFilterString = CreateFilterString
   
   RaiseEvent FilterStringChanged(NewFilterString)
   
End Sub

Public Sub RemoveFilter(Optional ByVal NewFilterString As String = vbNullString)
   
   FilterControls.RemoveFilterValues False
   RaiseEvent FilterStringChanged(NewFilterString)
   
End Sub
