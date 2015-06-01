VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterStringBuilderCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterStringBuilderImporter
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Code für Import von FilterStringBuilder
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/FilterStringBuilderCodeBuilder.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/FilterFormWizard/IFilterFormCodeBuilder.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements IFilterFormCodeBuilder

Private Property Get IFilterFormCodeBuilder_Self() As IFilterFormCodeBuilder
   Set IFilterFormCodeBuilder_Self = Me
End Property

Private Sub IFilterFormCodeBuilder_CheckRequiredCodeModules()
   With New AppFileCodeModulTransfer
      .TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder"
   End With
End Sub

Private Property Get IFilterFormCodeBuilder_DeclarationCode() As String
   IFilterFormCodeBuilder_DeclarationCode = vbNullString
End Property

Private Property Get IFilterFormCodeBuilder_ClassInitalisationCode( _
                        ByVal UseFilterMethodeCode As String, _
                        ByVal FilterControlCodeLines As String) As String
   
   Dim Code As String

   Code = _
      "Private Sub RefreshData()" & vbNewLine & _
      "   UseFilter GetFilterString()" & vbNewLine & _
      "End Sub" & vbNewLine

   Code = Code & vbNewLine & _
      UseFilterMethodeCode & vbNewLine

   Code = Code & vbNewLine & _
      "Private Function GetFilterString() As String" & vbNewLine & _
      vbNewLine & _
      "   With New FilterStringBuilder" & vbNewLine & _
      vbNewLine & _
      FilterControlCodeLines & vbNewLine & _
      vbNewLine & _
      "       GetFilterString = .ToString()" & vbNewLine & _
      vbNewLine & _
      "   End With" & vbNewLine & _
      vbNewLine & _
      "End Function" & vbNewLine

   IFilterFormCodeBuilder_ClassInitalisationCode = Code

End Property

Private Property Get IFilterFormCodeBuilder_FormLoadCode() As String
   IFilterFormCodeBuilder_FormLoadCode = vbNullString
End Property

Private Property Get IFilterFormCodeBuilder_FormUnLoadCode() As String
   IFilterFormCodeBuilder_FormUnLoadCode = vbNullString
End Property

Private Function IFilterFormCodeBuilder_GetFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                    ByVal Control As String, ByVal Control2 As String) As String
   
   Dim fcLine As String

   CheckControlRefCode Control

   fcLine = "Add """ & DataField & """, " & DataType & ", " & RelationalOperator & ", " & Control
   
   If Len(Control2) > 0 Then
      CheckControlRefCode Control2
      fcLine = fcLine & ", " & Control2
   End If

   IFilterFormCodeBuilder_GetFilterControlCodeLine = fcLine

End Function

Private Function IFilterFormCodeBuilder_GetMultiFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                         ByRef ControlArray() As String) As String

   Dim fcLine As String

   Dim i As Long
   For i = LBound(ControlArray) To UBound(ControlArray)
      CheckControlRefCode ControlArray(i)
   Next
   
   fcLine = "Add """ & DataField & """, " & DataType & ", , " & RelationalOperator & ", Array(" & Join(ControlArray, ", ") & "), , Null"
   IFilterFormCodeBuilder_GetMultiFilterControlCodeLine = fcLine

End Function

Private Sub CheckControlRefCode(ByRef ControlCode As String)

   ControlCode = Trim(ControlCode)

   If Len(ControlCode) = 0 Then
      Exit Sub
   End If

   If InStr(1, ControlCode, ".") = 0 Then
      ControlCode = "Me." & ControlCode & ".Value"
   End If

End Sub
