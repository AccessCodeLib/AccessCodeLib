VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormSortingControlEventBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: FormSortingControlEventBridge
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsklasse für FormSortingExtension
' </summary>
' <remarks>
' </remarks>
' \ingroup form
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/sort/FormSortingControlEventBridge.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/sort/FormSortingControl.cls</use>
'  <description>Internal helper for FormSortingExtension that forwards the click events of the sorting controls</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Event SortingControlClick(ByVal SortingControl As FormSortingControl)

Public Sub RaiseSortingControlClick(ByVal SortingControl As FormSortingControl)
   RaiseEvent SortingControlClick(SortingControl)
End Sub
