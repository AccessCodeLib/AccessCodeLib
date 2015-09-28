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
'  <file>form/filter/FilterControlManager_DotNetLib.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_dotnetlib/integration/DotNetLibsSetup.bas</use>
'  <use>form/filter/FilterControlCollection_DotNetLib.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_SqlTool As ACLibSqlTools.ISqlConverter
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

Friend Property Get SqlTool() As ACLibSqlTools.ISqlConverter
   If m_SqlTool Is Nothing Then
      Set m_SqlTool = NetSqlTools.SqlConverters.DaoSqlConverter
   End If
   Set SqlTool = m_SqlTool
End Property

Friend Property Set SqlTool(ByVal NewRef As ACLibSqlTools.ConditionStringBuilder)
   Set m_SqlTool = NewRef
End Property

Public Property Get FilterControls() As FilterControlCollection
   Set FilterControls = m_FilterControls
End Property

Public Function CreateFilterString() As String
   CreateFilterString = FilterControls.CriteriaString(SqlTool)
End Function

Private Sub m_FilterControls_FilterValuesChanged()
   UseAutoFilterMode
End Sub

Public Property Get AutoFilterOn() As Boolean
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
