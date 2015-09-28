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
'  <file>form/filter/FilterControlEventBridge_DotNetLib.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>form/filter/FilterControl_DotNetLib.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Event FilterValuesChanged(ByVal FilterControlRef As FilterControl)

Friend Sub RaiseFilterValueChanged(ByVal FilterControlRef As FilterControl)
   RaiseEvent FilterValuesChanged(FilterControlRef)
End Sub
