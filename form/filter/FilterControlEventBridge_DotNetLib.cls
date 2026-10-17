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
'  <file>form/filter/FilterControlEventBridge_DotNetLib.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControl_DotNetLib.cls</use>
'  <description>Internal event helper for the .NET (ACLib DotNetLib) filter controls</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Event FilterValuesChanged(ByVal FilterControlRef As FilterControl)

Friend Sub RaiseFilterValueChanged(ByVal FilterControlRef As FilterControl)
   RaiseEvent FilterValuesChanged(FilterControlRef)
End Sub
