VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlManagerCodeBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlManagerCodeBuilder
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Code für Import von FilterControlManager
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/FilterControlManagerCodeBuilder.cls</file>
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
      .TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager"
   End With
End Sub

Private Property Get IFilterFormCodeBuilder_DeclarationCode() As String
   IFilterFormCodeBuilder_DeclarationCode = _
      "Private WithEvents m_FilterControlManager As FilterControlManager" & vbNewLine
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
               Optional ByVal AutoFilterCtlName As String, _
               Optional ByVal FilterControlTagConverterMode As Long = 0) As String
   
   Dim Code As String

   Code = _
      "Private Sub m_FilterControlManager_FilterStringChanged(ByVal NewFilterString As String)" & vbNewLine & _
      "   ApplyFilter NewFilterString" & vbNewLine & _
      "End Sub" & vbNewLine
      
   If Len(ApplyFilterCtlName) > 0 Then
      Code = Code & vbNewLine & _
         "Private Sub m_FilterControlManager_FilterValuesChanged()" & vbNewLine & _
         "   SetApplyFilterSignal" & vbNewLine & _
         "End Sub" & vbNewLine
      
      Code = Code & vbNewLine & _
         "Private Sub SetApplyFilterSignal()" & vbNewLine & _
         "' " & L10n.Text("Notwendigen Klick auf ""Filter anwenden"" signalisieren") & vbNewLine
      
      If Len(AutoFilterCtlName) > 0 Then
         Code = Code & _
            "   If Me." & AutoFilterCtlName & ".Value = False Then" & vbNewLine & _
            "      Me." & ApplyFilterCtlName & ".FontBold = True" & vbNewLine & _
            "   End If" & vbNewLine
      Else
         Code = Code & _
            "   Me." & ApplyFilterCtlName & ".FontBold = True" & vbNewLine
      End If
     
      Code = Code & _
         "End Sub" & vbNewLine
      
   End If

   Code = Code & vbNewLine & _
      ApplyFilterMethodeCode & vbNewLine

   Code = Code & vbNewLine & _
      "Private Sub InitFilterControlManager()" & vbNewLine & _
      "   Set m_FilterControlManager = New FilterControlManager" & vbNewLine & _
      "   m_FilterControlManager.ConfigSqlFormat """ & SqlDateFormat & """, """ & SqlBooleanTrueString & """, """ & SqlWildCardString & """" & vbNewLine & _
      "   InitFilterControls" & vbNewLine
      
   If Len(AutoFilterCtlName) Then
      Code = Code & "   m_FilterControlManager.AutoFilterOn = Nz(Me." & AutoFilterCtlName & ".Value, False)" & vbNewLine
   Else
      Code = Code & "   m_FilterControlManager.AutoFilterOn = True '" & L10n.Text("oder z. B. mittels Checkbox:") & " = Nz(Me.cbAutoFilter.Value, False)" & vbNewLine
   End If
   Code = Code & "End Sub" & vbNewLine

   Code = Code & vbNewLine & _
      "Private Property Get FilterControlManager() As FilterControlManager" & vbNewLine & _
      "   If m_FilterControlManager Is Nothing Then InitFilterControlManager" & vbNewLine & _
      "   Set FilterControlManager = m_FilterControlManager" & vbNewLine & _
      "End Property" & vbNewLine

   Code = Code & vbNewLine & _
      GetInitFilterControlsCode(FilterControlCodeLines, FilterControlTagConverterMode <> 0, FilterControlTagConverterMode = 2)
   
   Code = Code & vbNewLine & _
      GetRemoveFilterCode()

   IFilterFormCodeBuilder_ClassInitalisationCode = Code

End Property

Private Function GetInitFilterControlsCode(ByVal FilterControlCodeLines As String, ByVal UseFilterControlTagConverter As Boolean, ByVal UseHeaderControls As Boolean) As String
   
   Dim Code As String
   Dim ControlsCollectionCode As String

   Code = _
      "Private Sub InitFilterControls()" & vbNewLine & _
      vbNewLine
      
   If UseFilterControlTagConverter Then
   
      If UseHeaderControls Then
         ControlsCollectionCode = "Me.Section(acHeader).Controls"
      Else
         ControlsCollectionCode = "Me.Controls"
      End If
   
      Code = Code & _
         "   With New FilterControlTagConverter" & vbNewLine & _
         "      .AddFilterControls FilterControlManager.FilterControls, " & ControlsCollectionCode & vbNewLine & _
         "   End With" & vbNewLine
   Else
      Code = Code & _
         "   With FilterControlManager.FilterControls" & vbNewLine & _
         vbNewLine & _
         FilterControlCodeLines & vbNewLine & _
         vbNewLine & _
         "   End With" & vbNewLine
   End If
   
   Code = Code & _
      vbNewLine & _
      "End Sub" & vbNewLine
      
   GetInitFilterControlsCode = Code

End Function

Private Property Get IFilterFormCodeBuilder_FormLoadCode() As String
   IFilterFormCodeBuilder_FormLoadCode = "   InitFilterControlManager"
End Property

Private Property Get IFilterFormCodeBuilder_FormUnLoadCode() As String
   IFilterFormCodeBuilder_FormUnLoadCode = "   Set m_FilterControlManager = Nothing"
End Property

Private Property Get IFilterFormCodeBuilder_ApplyFilterCtlCode() As String
   IFilterFormCodeBuilder_ApplyFilterCtlCode = "   FilterControlManager.ApplyFilter"
End Property

Private Property Get IFilterFormCodeBuilder_RemoveFilterCtlCode() As String
   IFilterFormCodeBuilder_RemoveFilterCtlCode = "   RemoveFilter"
End Property

Private Function IFilterFormCodeBuilder_AutoFilterCtlCode(Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   If Len(AutoFilterCtlName) = 0 Then Exit Function
   IFilterFormCodeBuilder_AutoFilterCtlCode = "   m_FilterControlManager.AutoFilterOn = Me." & AutoFilterCtlName & ".Value"
End Function

Private Function IFilterFormCodeBuilder_FilterControlValueChangedEventFunction(Optional ByVal ApplyFilterCtlName As String = vbNullString, Optional ByVal AutoFilterCtlName As String = vbNullString) As String
   IFilterFormCodeBuilder_FilterControlValueChangedEventFunction = vbNullString
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
   
   FcLine = "AddMultiControlCriteria """ & DataField & """, " & DataType & ", " & RelationalOperator & ", Null, " & Join(ControlArray, ", ")
   
   IFilterFormCodeBuilder_GetMultiFilterControlCodeLine = FcLine

End Function

Private Sub CheckControlRefCode(ByRef ControlCode As String)

   ControlCode = Trim(ControlCode)

   If Len(ControlCode) = 0 Then
      Exit Sub
   End If

   If InStr(1, ControlCode, ".") = 0 Then
      ControlCode = "Me." & ControlCode
   End If

End Sub

Private Function GetRemoveFilterCode() As String

   Dim Code As String

   Code = "Private Sub RemoveFilter()" & vbNewLine & _
          "'  FilterControlManager.RemoveFilter ""1=0""  ' 1=0 .. " & L10n.Text("damit nach dem Entfernen der Filterwerte keine Datensätze angezeigt werden") & vbNewLine & _
          "   FilterControlManager.RemoveFilter        ' " & L10n.Text("ohne Parameterübergabe werden alle Datensätze angezeigt") & vbNewLine & _
          "End Sub" & vbNewLine

   GetRemoveFilterCode = Code

End Function
