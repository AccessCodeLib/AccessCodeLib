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
   
   Dim fc As FilterControl

   With New FilterStringBuilder
      For Each fc In FilterControls.Items
         .AddCriteria fc.CriteriaString(SqlTool)
      Next
      CreateFilterString = .ToString(" AND ")
   End With

End Function

Private Sub m_FilterControls_FilterValuesChanged()
   UseAutoFilterMode
End Sub

Private Property Get AutoFilterOn() As Boolean
   AutoFilterOn = m_AutoFilterModeOn
End Property

Public Property Let AutoFilterOn(ByVal NewValue As Boolean)
   m_AutoFilterModeOn = NewValue
   UseAutoFilterMode
End Property

Private Sub UseAutoFilterMode()
   If m_AutoFilterModeOn Then
      UseFilter
   End If
End Sub

Public Sub UseFilter()

   Dim NewFilterString As String
   NewFilterString = CreateFilterString
   
   RaiseEvent FilterStringChanged(NewFilterString)
   
End Sub

Public Sub RemoveFilter()
   
   FilterControls.RemoveFilterValues False
   RaiseEvent FilterStringChanged(vbNullString)
   
End Sub
