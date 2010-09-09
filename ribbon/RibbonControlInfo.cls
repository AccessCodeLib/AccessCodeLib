VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RibbonControlInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: RibbonControlInfo (Josef Pötzl, 2010-04-10)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Klasse zum Speichern der RibbonControl-Daten
' </summary>
' <remarks>
' Implements IRibbonControl
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>ribbon/RibbonControlInfo.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements IRibbonControl

Public ID As String
Public Tag As String
Public Label As String
Public OnAction As String
Public ImageMso As String
Public visible As Boolean

Private Property Get IRibbonControl_Context() As Object
   Set IRibbonControl_Context = Nothing
End Property

Private Property Get IRibbonControl_Id() As String
   IRibbonControl_Id = Me.ID
End Property

Private Property Get IRibbonControl_Tag() As String
   IRibbonControl_Tag = Me.Tag
End Property
