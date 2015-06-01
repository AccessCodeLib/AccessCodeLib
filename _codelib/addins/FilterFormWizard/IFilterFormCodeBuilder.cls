VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFilterFormCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Interface: IFilterFormCodeBuilder
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' CodeBuilder-Schnittstelle zum Einfügen von Code für FilterControlManager, FilterStringBuilder, ...
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/IFilterFormCodeBuilder.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Property Get Self() As IFilterFormCodeBuilder
'
End Property

Public Sub CheckRequiredCodeModules()
'
End Sub

Public Property Get DeclarationCode() As String
'
End Property

Public Property Get ClassInitalisationCode( _
                        ByVal UseFilterMethodeCode As String, _
                        ByVal FilterControlCodeLines As String) As String
'
End Property

Public Property Get FormLoadCode() As String
'
End Property

Public Property Get FormUnLoadCode() As String
'
End Property

Public Function GetFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                    ByVal Control As String, ByVal Control2 As String) As String
'
End Function

Public Function GetMultiFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                         ByRef ControlArray() As String) As String
'
End Function
