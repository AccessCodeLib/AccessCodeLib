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
'  <use>_codelib/addins/shared/AppFileCodeModulTransfer.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Implements IFilterFormCodeBuilder

Private Property Get IFilterFormCodeBuilder_Self() As IFilterFormCodeBuilder
   Set IFilterFormCodeBuilder_Self = Me
End Property

Private Sub IFilterFormCodeBuilder_CheckRequiredCodeModules(ByVal CodemoduleImporter As AppFileCodeModulTransfer)
   With CodemoduleImporter
      .TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder"
   End With
End Sub

Private Property Get IFilterFormCodeBuilder_DeclarationCode() As String
   IFilterFormCodeBuilder_DeclarationCode = vbNullString
End Property

Private Property Get IFilterFormCodeBuilder_ClassInitalisationCode( _
                        ByVal ApplyFilterMethodeCode As String, _
                        ByVal FilterControlCodeLines As String, _
                        ByVal FilterControlNames As StringCollection, _
               Optional ByVal SqlDateFormat As String = "\#yyyy-mm-dd hh:nn:ss\#", _
               Optional ByVal SqlBooleanTrueString As String = "True", _
               Optional ByVal SqlWildCardString As String = "*", _
               Optional ByVal ApplyFilterCtlName As String, _
               Optional ByVal RemoveFilterCtlName As String, _
               Optional ByVal AutoFilterCtlName As String) As String
   
   Dim Code As String

   Code = GetRefreshFilterCode() & vbNewLine
      
   Code = Code & vbNewLine & _
      ApplyFilterMethodeCode & vbNewLine

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
      GetRemoveFilterCode(ApplyFilterCtlName, FilterControlNames)

   IFilterFormCodeBuilder_ClassInitalisationCode = Code

End Property

Private Function GetRefreshFilterCode() As String

   Dim Code As String

   Code = "Private Sub RefreshFilter()" & vbNewLine & _
          "   ApplyFilter GetFilterString()" & vbNewLine & _
          "End Sub"

   GetRefreshFilterCode = Code

End Function

Private Function GetRemoveFilterCode(ByVal ApplyFilterCtlName As String, ByVal FilterControlNames As StringCollection) As String

   Dim Code As String

   Code = "Private Sub RemoveFilter()" & vbNewLine & _
          "   RemoveFilterValues" & vbNewLine & _
          "   ApplyFilter ""0=1"" ' " & L10n.Text("Anzeige leeren") & vbNewLine
          
   If Len(ApplyFilterCtlName) > 0 Then
       Code = Code & _
          "   Me." & ApplyFilterCtlName & ".FontBold = False" & vbNewLine
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
          "   '" & L10n.Text("Filter-Steuerelemente anfügen:") & vbNewLine
          
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

Private Property Get IFilterFormCodeBuilder_ApplyFilterCtlCode() As String
   IFilterFormCodeBuilder_ApplyFilterCtlCode = "   RefreshFilter"
End Property

Private Property Get IFilterFormCodeBuilder_RemoveFilterCtlCode() As String
   IFilterFormCodeBuilder_RemoveFilterCtlCode = "   RemoveFilter"
End Property

Private Function IFilterFormCodeBuilder_AutoFilterCtlCode(Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   IFilterFormCodeBuilder_AutoFilterCtlCode = vbNullString
End Function

Private Function IFilterFormCodeBuilder_FilterControlValueChangedEventFunction(Optional ByVal ApplyFilterCtlName As String = vbNullString, Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   IFilterFormCodeBuilder_FilterControlValueChangedEventFunction = GetControlValueChangedCode(ApplyFilterCtlName, AutoFilterCtlName)
End Function

Private Function IFilterFormCodeBuilder_GetFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                    ByVal Control As String, ByVal Control2 As String) As String
   
   Dim FcLine As String

   CheckControlRefCode Control

   FcLine = "Add """ & DataField & """, " & DataType & ", " & RelationalOperator & ", " & Control
   
   If Len(Control2) > 0 Then
      CheckControlRefCode Control2
      FcLine = FcLine & ", " & Control2
   End If

   IFilterFormCodeBuilder_GetFilterControlCodeLine = FcLine

End Function

Private Function IFilterFormCodeBuilder_GetMultiFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                         ByRef ControlArray() As String) As String

   Dim FcLine As String

   Dim i As Long
   For i = LBound(ControlArray) To UBound(ControlArray)
      CheckControlRefCode ControlArray(i)
   Next
   
   FcLine = "Add """ & DataField & """, " & DataType & ", , " & RelationalOperator & ", Array(" & Join(ControlArray, ", ") & "), , Null"
   IFilterFormCodeBuilder_GetMultiFilterControlCodeLine = FcLine

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

Private Function GetControlValueChangedCode(ByVal ApplyFilterCtlName As String, ByVal AutoFilterCtlName As String) As String

   Dim Code As String
   Dim ApplyFilterCtlLine As String
 
   If Len(ApplyFilterCtlName) + Len(AutoFilterCtlName) = 0 Then
      GetControlValueChangedCode = vbNullString
      Exit Function
   End If

   Code = "Private Function FilterControlValueChanged()" & vbNewLine
   
   If Len(ApplyFilterCtlName) > 0 Then
      ApplyFilterCtlLine = "   Me." & ApplyFilterCtlName & ".FontBold = True" & vbNewLine
   End If
   
   If Len(AutoFilterCtlName) > 0 Then
      Code = Code & _
          "   If Me." & AutoFilterCtlName & ".Value = True Then" & vbNewLine & _
          "      RefreshFilter" & vbNewLine
          
      If Len(ApplyFilterCtlName) > 0 Then
         Code = Code & _
          "   Else" & vbNewLine & _
          "   " & ApplyFilterCtlLine
      End If
      
      Code = Code & _
          "   End If" & vbNewLine
   Else
      Code = Code & ApplyFilterCtlLine
   End If
   
   Code = Code & _
      "End Function"
   
   GetControlValueChangedCode = Code
   
End Function
