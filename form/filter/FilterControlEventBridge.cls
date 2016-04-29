VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlEventBridge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlEventBridge
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Hilfsklasse zur Weitergabe von Ereignissen
' </summary>
' <remarks>Wird für FilterControlCollection benötigt</remarks>
' \ingroup data
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControlEventBridge.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControl.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Event FilterValuesChanged(ByVal FilterControlRef As FilterControl)
Public Event FilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)

Friend Sub RaiseFilterValueChanged(ByVal FilterControlRef As FilterControl)
   RaiseEvent FilterValuesChanged(FilterControlRef)
End Sub

Friend Sub RaiseFilterValueRequest(ByVal FilterValueCode As String, ByRef FilterValue As String)
   RaiseEvent FilterValueRequest(FilterValueCode, FilterValue)
End Sub
