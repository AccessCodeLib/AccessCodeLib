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
                        ByVal FilterControlCodeLines As String, _
                        ByVal FilterControlNames As StringCollection, _
               Optional ByVal SqlDateFormat As String = "\#yyyy-mm-dd hh:nn:ss\#", _
               Optional ByVal SqlBooleanTrueString As String = "True", _
               Optional ByVal SqlWildCardString As String = "*", _
               Optional ByVal UseFilterCtlName As String, _
               Optional ByVal RemoveFilterCtlName As String, _
               Optional ByVal AutoFilterCtlName As String) As String
   
   Dim Code As String

   Code = GetRefreshFilterCode(UseFilterCtlName) & vbNewLine
      
   Code = Code & vbNewLine & _
      UseFilterMethodeCode & vbNewLine

   Code = Code & vbNewLine & _
      "Private Function GetFilterString() As String" & vbNewLine & _
      vbNewLine & _
      "   With New FilterStringBuilder" & vbNewLine & vbNewLine & _
      "      .ConfigSqlFormat """ & SqlDateFormat & """, """ & SqlBooleanTrueString & """, """ & SqlWildCardString & """" & vbNewLine & _
      vbNewLine & _
      FilterControlCodeLines & vbNewLine & _
      vbNewLine & _
      "      GetFilterString = .ToString()" & vbNewLine & _
      vbNewLine & _
      "   End With" & vbNewLine & _
      vbNewLine & _
      "End Function" & vbNewLine
      
   Code = Code & vbNewLine & _
      GetRemoveFilterCode(UseFilterCtlName, FilterControlNames)

   IFilterFormCodeBuilder_ClassInitalisationCode = Code

End Property

Private Function GetRefreshFilterCode(ByVal UseFilterCtlName As String) As String

   Dim Code As String

   Code = "Private Sub RefreshFilter()" & vbNewLine & _
          "   UseFilter GetFilterString()" & vbNewLine & _
          "End Sub"

   GetRefreshFilterCode = Code

End Function

Private Function GetRemoveFilterCode(ByVal UseFilterCtlName As String, ByVal FilterControlNames As StringCollection) As String

   Dim Code As String

   Code = "Private Sub RemoveFilter()" & vbNewLine & _
          "   RemoveFilterValues" & vbNewLine & _
          "   UseFilter ""0=1"" ' Anzeige leeren" & vbNewLine
          
   If Len(UseFilterCtlName) > 0 Then
       Code = Code & _
          "   Me." & UseFilterCtlName & ".FontBold = False" & vbNewLine
   End If
   
   Code = Code & _
      "End Sub" & vbNewLine
      
   Code = Code & vbNewLine & GetRemoveFilterValuesCode(FilterControlNames)

   GetRemoveFilterCode = Code

End Function

Private Function GetRemoveFilterValuesCode(ByVal FilterControlNames As StringCollection) As String
   
   Dim Code As String

   Code = "Private Sub RemoveFilterValues()" & vbNewLine & _
          "   Dim fctl As Control" & vbNewLine & _
          "   For Each fctl In GetFilterControls()" & vbNewLine & _
          "      fctl.Value = Null" & vbNewLine & _
          "   Next" & vbNewLine & _
          "End Sub" & vbNewLine
              
   Code = Code & vbNewLine & GetGetFilterControlsCode(FilterControlNames)

   GetRemoveFilterValuesCode = Code
   
End Function

Private Function GetGetFilterControlsCode(ByVal FilterControlNames As StringCollection) As String
   
   Dim Code As String

   Code = "Private Function GetFilterControls() As Collection" & vbNewLine & _
          "   Dim fctlCol As Collection" & vbNewLine & vbNewLine & _
          "   Set fctlCol = New Collection" & vbNewLine & _
          "   'Filter-Steuerelemente anfügen:" & vbNewLine
          
   Code = Code & FilterControlNames.ToString(vbNewLine, "   fctlCol.Add Me.", , True) & vbNewLine
          
   Code = Code & _
          "   Set GetFilterControls = fctlCol" & vbNewLine & _
          "End Function" '& vbNewLine
          
   GetGetFilterControlsCode = Code
   
End Function


Private Property Get IFilterFormCodeBuilder_FormLoadCode() As String
   IFilterFormCodeBuilder_FormLoadCode = vbNullString
End Property

Private Property Get IFilterFormCodeBuilder_FormUnLoadCode() As String
   IFilterFormCodeBuilder_FormUnLoadCode = vbNullString
End Property

Private Property Get IFilterFormCodeBuilder_UseFilterCtlCode() As String
   IFilterFormCodeBuilder_UseFilterCtlCode = "   RefreshFilter"
End Property

Private Property Get IFilterFormCodeBuilder_RemoveFilterCtlCode() As String
   IFilterFormCodeBuilder_RemoveFilterCtlCode = "   RemoveFilter"
End Property

Private Function IFilterFormCodeBuilder_AutoFilterCtlCode(Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   IFilterFormCodeBuilder_AutoFilterCtlCode = vbNullString
End Function

Private Function IFilterFormCodeBuilder_FilterControlValueChangedEventFunction(Optional ByVal UseFilterCtlName As String = vbNullString, Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   IFilterFormCodeBuilder_FilterControlValueChangedEventFunction = GetControlValueChangedCode(UseFilterCtlName, AutoFilterCtlName)
End Function

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

Private Function GetControlValueChangedCode(ByVal UseFilterCtlName As String, ByVal AutoFilterCtlName As String) As String

   Dim Code As String
   Dim UseFilterCtlLine As String
 
   If Len(UseFilterCtlName) + Len(AutoFilterCtlName) = 0 Then
      GetControlValueChangedCode = vbNullString
      Exit Function
   End If

   Code = "Private Function FilterControlValueChanged()" & vbNewLine
   
   If Len(UseFilterCtlName) > 0 Then
      UseFilterCtlLine = "   Me." & UseFilterCtlName & ".FontBold = True" & vbNewLine
   End If
   
   If Len(AutoFilterCtlName) > 0 Then
      Code = Code & _
          "   If Me." & AutoFilterCtlName & ".Value = True Then" & vbNewLine & _
          "      RefreshFilter" & vbNewLine
          
      If Len(UseFilterCtlName) > 0 Then
         Code = Code & _
          "   Else" & vbNewLine & _
          "   " & UseFilterCtlLine
      End If
      
      Code = Code & _
          "   End If" & vbNewLine
   Else
      Code = Code & UseFilterCtlLine
   End If
   
   Code = Code & _
      "End Function"
   
   GetControlValueChangedCode = Code
   
End Function
