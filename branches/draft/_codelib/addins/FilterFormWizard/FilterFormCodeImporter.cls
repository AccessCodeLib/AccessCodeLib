VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterFormCodeImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterFormCodeImporter
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
'  <file>_codelib/addins/FilterFormWizard/FilterFormCodeImporter.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'  <use>_codelib/addins/shared/AppFileCodeModulTransfer.cls</use>
'  <use>_codelib/addins/FilterFormWizard/IFilterFormCodeBuilder.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_FilterControlNames As StringCollection
Private m_FilterControlLines As StringCollection

Private m_Importer As IFilterFormCodeBuilder
Private m_CodeModulImporter As AppFileCodeModulTransfer

Private Sub Class_Initialize()
   Set m_FilterControlNames = New StringCollection
   Set m_FilterControlLines = New StringCollection
End Sub

Friend Property Get Self() As FilterFormCodeImporter
   Set Self = Me
End Property

Public Property Get Importer() As IFilterFormCodeBuilder
   Set Importer = m_Importer
End Property

Public Property Set Importer(ByVal NewRef As IFilterFormCodeBuilder)
   Set m_Importer = NewRef
End Property

Private Property Get CodeModulImporter() As AppFileCodeModulTransfer
   If m_CodeModulImporter Is Nothing Then
      Set m_CodeModulImporter = New AppFileCodeModulTransfer
      m_CodeModulImporter.UseVbComponentsImport = APPLICATION_FILTERCODEMODULE_USEVBCOMPONENTSIMPORT
   End If
   Set CodeModulImporter = m_CodeModulImporter
End Property

Friend Sub WriteToForm(ByVal FormName As String, _
              Optional ByVal KindOfApplyFilterMethode As Long = 0, _
              Optional ByVal ApplyFilterMethodeSubFormName As String, _
              Optional ByVal SqlDateFormat As String = vbNullString, _
              Optional ByVal SqlBooleanTrueString As String = vbNullString, _
              Optional ByVal SqlWildCardString As String = vbNullString, _
              Optional ByVal ApplyFilterCtlName As String, _
              Optional ByVal RemoveFilterCtlName As String, _
              Optional ByVal AutoFilterCtlName As String, _
              Optional ByVal FilterControlTagConverterMode As Long = 0)
   
   CheckRequiredCodeModules CodeModulImporter
   
   If CurrentProject.AllForms(FormName).IsLoaded Then
      DoCmd.Close acForm, FormName
   End If

   DoCmd.OpenForm FormName, acDesign, , , , acWindowNormal

   AddCodeToForm Application.Forms(FormName), KindOfApplyFilterMethode, ApplyFilterMethodeSubFormName, _
                  SqlDateFormat, SqlBooleanTrueString, SqlWildCardString, _
                 ApplyFilterCtlName, RemoveFilterCtlName, AutoFilterCtlName, FilterControlTagConverterMode

   DoCmd.Save acForm, FormName

End Sub

Private Sub CheckRequiredCodeModules(ByVal CodemoduleImporter As AppFileCodeModulTransfer)
   m_Importer.CheckRequiredCodeModules CodemoduleImporter
End Sub

Private Sub AddCodeToForm(ByVal FrmRef As Form, _
                          ByVal KindOfApplyFilterMethode As Long, _
                          ByVal ApplyFilterMethodeSubFormName As String, _
                          ByVal SqlDateFormat As String, _
                          ByVal SqlBooleanTrueString As String, _
                          ByVal SqlWildCardString As String, _
                          ByVal ApplyFilterCtlName As String, _
                          ByVal RemoveFilterCtlName As String, _
                          ByVal AutoFilterCtlName As String, _
                          ByVal FilterControlTagConverterMode As Long)

   Dim TempCode As String

   CheckCodeModuleBehindForm FrmRef

   With New CodeModuleWriter

      .Load "Form_" & FrmRef.Name
      
      TempCode = DeclarationCode
      If Len(TempCode) > 0 Then
         TempCode = TempCode & vbNewLine
      End If
      
      TempCode = TempCode & _
                 ClassInitalisationCode(KindOfApplyFilterMethode, ApplyFilterMethodeSubFormName, _
                           SqlDateFormat, SqlBooleanTrueString, SqlWildCardString, _
                           ApplyFilterCtlName, RemoveFilterCtlName, AutoFilterCtlName, FilterControlTagConverterMode)

      .AddCode TempCode
   
      TempCode = FormLoadCode
      If Len(TempCode) > 0 Then
         .AddCodeToEventProcedure "Load", "Form", TempCode
      End If
      
      TempCode = FormUnLoadCode
      If Len(TempCode) > 0 Then
         .AddCodeToEventProcedure "Unload", "Form", TempCode
      End If
      
      If Len(ApplyFilterCtlName) > 0 Then
         TempCode = ApplyFilterCtlCode
         If Len(TempCode) > 0 Then
            .AddCodeToEventProcedure "Click", ApplyFilterCtlName, TempCode
         End If
      End If
      
      If Len(AutoFilterCtlName) > 0 Then
         TempCode = AutoFilterCtlCode(AutoFilterCtlName)
         If Len(TempCode) > 0 Then
            .AddCodeToEventProcedure "Click", AutoFilterCtlName, TempCode
         End If
      End If
      
      If Len(RemoveFilterCtlName) > 0 Then
         TempCode = RemoveFilterCtlCode
         If Len(TempCode) > 0 Then
            .AddCodeToEventProcedure "Click", RemoveFilterCtlName, TempCode
         End If
      End If
      
      TempCode = FilterControlValueChangedEventFunction(ApplyFilterCtlName, AutoFilterCtlName)
      If Len(TempCode) > 0 Then
         .AddCode TempCode
         InsertFilterControlValueChangedEventFunction FrmRef
      End If
      
   End With

End Sub

Private Sub CheckCodeModuleBehindForm(ByVal FrmRef As Form)

   If Not FrmRef.HasModule Then
      FrmRef.HasModule = True
      DoCmd.Save acForm, FrmRef.Name
   End If

End Sub

Private Property Get DeclarationCode() As String
   DeclarationCode = m_Importer.DeclarationCode
End Property

Private Property Get ClassInitalisationCode( _
                        ByVal KindOfApplyFilterMethode As Long, _
                        ByVal ApplyFilterMethodeSubFormName As String, _
                        ByVal SqlDateFormat As String, _
                        ByVal SqlBooleanTrueString As String, _
                        ByVal SqlWildCardString As String, _
                        ByVal ApplyFilterCtlName As String, _
                        ByVal RemoveFilterCtlName As String, _
                        ByVal AutoFilterCtlName As String, _
                        ByVal FilterControlTagConverterMode As Long) As String
   
   ClassInitalisationCode = m_Importer.ClassInitalisationCode( _
                  GetApplyFilterMethodeCode(KindOfApplyFilterMethode, ApplyFilterMethodeSubFormName, ApplyFilterCtlName, False), _
                  GetFilterControlCodeLines, m_FilterControlNames, _
                  SqlDateFormat, SqlBooleanTrueString, SqlWildCardString, _
                  ApplyFilterCtlName, RemoveFilterCtlName, AutoFilterCtlName, FilterControlTagConverterMode)

End Property

Private Function GetApplyFilterMethodeCode( _
                        ByVal KindOfApplyFilterMethode As Long, _
                        ByVal ApplyFilterMethodeSubFormName As String, _
                        ByVal ApplyFilterCtlName As String, _
                        ByVal AddDebugLine As Boolean) As String
   
   Dim MethodeCode As String
   Dim SufFormRefCode As String
   Dim LinePrefix As String
   
   Select Case KindOfApplyFilterMethode
      Case 0
         LinePrefix = "'"
         SufFormRefCode = " ' oder Me.EinUnterformular.Form"
         AddDebugLine = True
      Case 1
         LinePrefix = vbNullString
         SufFormRefCode = vbNullString
      Case 2
         LinePrefix = vbNullString
         If Len(ApplyFilterMethodeSubFormName) = 0 Then
            Err.Raise vbObjectError, "FilterControlManagerImporter", "Bitte Unterformular angeben"
         End If
         SufFormRefCode = "." & ApplyFilterMethodeSubFormName & ".Form"
   End Select

   MethodeCode = _
         "Private Sub ApplyFilter(ByVal NewFilterString As String)" & vbNewLine
         
   If AddDebugLine Then
      MethodeCode = MethodeCode & vbNewLine & _
         "   Debug.Print ""FilterText:""; NewFilterString" & vbNewLine
   End If
   
   MethodeCode = MethodeCode & _
         vbNewLine & _
         LinePrefix & "   With Me" & SufFormRefCode & vbNewLine & _
         LinePrefix & "      .Filter = NewFilterString" & vbNewLine & _
         LinePrefix & "      .FilterOn = (Len(.Filter) > 0)" & vbNewLine & _
         LinePrefix & "   End With" & vbNewLine
         
         
   If Len(ApplyFilterCtlName) > 0 Then
       MethodeCode = MethodeCode & vbNewLine & _
         "   Me." & ApplyFilterCtlName & ".FontBold = False" & vbNewLine
   End If
         
   MethodeCode = MethodeCode & vbNewLine & _
         "End Sub"

   GetApplyFilterMethodeCode = MethodeCode

End Function

Private Function GetFilterControlCodeLines() As String

   If m_FilterControlLines.Items.Count = 0 Then
      GetFilterControlCodeLines = _
         "'      .Add ""DatenfeldName"", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.Steuerelement" & vbNewLine & _
         "'      .Add ..." & vbNewLine & _
         "'      .Add ..."
         Exit Function
   End If

   GetFilterControlCodeLines = _
      m_FilterControlLines.ToString(vbNewLine, "      .")

End Function

Private Property Get FormLoadCode() As String
   FormLoadCode = m_Importer.FormLoadCode
End Property

Private Property Get FormUnLoadCode() As String
   FormUnLoadCode = m_Importer.FormUnLoadCode
End Property

Private Property Get ApplyFilterCtlCode() As String
   ApplyFilterCtlCode = m_Importer.ApplyFilterCtlCode
End Property

Private Property Get RemoveFilterCtlCode() As String
   RemoveFilterCtlCode = m_Importer.RemoveFilterCtlCode
End Property

Private Property Get AutoFilterCtlCode(ByVal AutoFilterCtlName As String) As String
   AutoFilterCtlCode = m_Importer.AutoFilterCtlCode(AutoFilterCtlName)
End Property

Private Property Get FilterControlValueChangedEventFunction(ByVal ApplyFilterCtlName As String, ByVal AutoFilterCtlName As String) As String
   FilterControlValueChangedEventFunction = m_Importer.FilterControlValueChangedEventFunction(ApplyFilterCtlName, AutoFilterCtlName)
End Property

Private Sub InsertFilterControlValueChangedEventFunction(ByVal FormRef As Form)
   
   Dim FcltName As Variant
   Dim ctl As Control

   For Each FcltName In m_FilterControlNames.Items
      Set ctl = FormRef.Controls(FcltName)
      If Len(ctl.AfterUpdate) = 0 Then
         ctl.AfterUpdate = "=FilterControlValueChanged()"
      End If
   Next
   
End Sub

Public Sub AddFilterControlDefinition(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                      ByVal Control As String, ByVal Control2 As String)
   
   Dim ControlArray() As String

   If Len(Control2) > 0 Then
      If Not (RelationalOperator Like "*SQL_Between*") Then
      ' Control + Control2 werden nur bei Between benötigt => Rest: MultiFilterControlCodeLine
         If Len(Control) > 0 Then
            If Len(Control2) > 0 Then
               Control2 = "," & Control2
            End If
            Control2 = Control & Control2
            Control = vbNullString
         End If
      End If
   End If

   If Len(Control2) > 0 Then
      CleanControlArrayString Control2
      ControlArray = Split(Control2, ",")
      If UBound(ControlArray) > 0 Then
         m_FilterControlLines.Add GetMultiFilterControlCodeLine(DataField, DataType, RelationalOperator, ControlArray)
         Exit Sub
      End If
   End If

   m_FilterControlNames.Add Control
   If Len(Control2) > 0 Then
      m_FilterControlNames.Add Control2
   End If
   
   m_FilterControlLines.Add GetFilterControlCodeLine(DataField, DataType, RelationalOperator, Control, Control2)

End Sub

Private Function GetFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                    ByVal Control As String, ByVal Control2 As String) As String

   GetFilterControlCodeLine = m_Importer.GetFilterControlCodeLine(DataField, DataType, RelationalOperator, Control, Control2)

End Function

Private Function GetMultiFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                         ByRef ControlArray() As String) As String

   GetMultiFilterControlCodeLine = m_Importer.GetMultiFilterControlCodeLine(DataField, DataType, RelationalOperator, ControlArray)

End Function

Private Sub CleanControlArrayString(ByRef CtlString As String)
   Do While InStr(1, CtlString, ", ") > 0
      CtlString = Replace(CtlString, ", ", ",")
   Loop
End Sub
