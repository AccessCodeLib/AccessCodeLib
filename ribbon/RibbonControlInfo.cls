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
' Class: ribbon.RibbonControlInfo
'---------------------------------------------------------------------------------------
'
' Class for storing the RibbonControl data
'
' Author:
'     Josef Pötzl
'
' Remarks:
'     Implements IRibbonControl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>ribbon/RibbonControlInfo.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>Office</name><major>2</major><minor>4</minor><guid>{2DF8D04C-5BFA-101B-BDE5-00AA0044DE52}</guid></ref>
'  <description>Data of a ribbon control used by RibbonWatcher</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Implements IRibbonControl

Public ID As String
Public Tag As String
Public Label As String
Public OnAction As String
Public ImageMso As String
Public Visible As Boolean

Private Property Get IRibbonControl_Context() As Object
   Set IRibbonControl_Context = Nothing
End Property

Private Property Get IRibbonControl_Id() As String
   IRibbonControl_Id = Me.ID
End Property

Private Property Get IRibbonControl_Tag() As String
   IRibbonControl_Tag = Me.Tag
End Property
