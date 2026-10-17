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
' Class: form.filter.FilterControlEventBridge
'---------------------------------------------------------------------------------------
'
' Helper class for passing on events
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     Required for FilterControlCollection
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>form/filter/FilterControlEventBridge.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControl.cls</use>
'  <description>Internal helper that forwards value-change events from the FilterControls to the collection</description>
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
