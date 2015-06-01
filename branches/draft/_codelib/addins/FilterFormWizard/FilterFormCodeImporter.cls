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

Private m_FilterControlLines As StringCollection

Private m_Importer As IFilterFormCodeBuilder

Private Sub Class_Initialize()
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

Friend Sub WriteToForm(ByVal FormName As String, _
              Optional ByVal KindOfUseFilterMethode As Long = 0, Optional ByVal UseFilterMethodeSubFormName As String)
   
   Dim frm As Form

   CheckRequiredCodeModules
   
   If CurrentProject.AllForms(FormName).IsLoaded Then
      DoCmd.Close acForm, FormName
   End If

   DoCmd.OpenForm FormName, acDesign, , , , acWindowNormal

   AddCodeToFrom Application.Forms(FormName), KindOfUseFilterMethode, UseFilterMethodeSubFormName

   DoCmd.Save acForm, FormName

End Sub

Private Function CheckRequiredCodeModules() As Boolean
   m_Importer.CheckRequiredCodeModules
End Function

Private Sub AddCodeToFrom(ByVal frm As Form, _
                                              ByVal KindOfUseFilterMethode As Long, _
                                              ByVal UseFilterMethodeSubFormName As String)

#If EARLYBINDING Then
   Dim cm As CodeModule
#Else
   Dim cm As Object
#End If

   Dim TempCode As String

   CheckCodeModuleBehindForm frm

   With New CodeModuleWriter

      .Load "Form_" & frm.Name

      .AddCode DeclarationCode & vbNewLine & _
               ClassInitalisationCode(KindOfUseFilterMethode, UseFilterMethodeSubFormName)
   
      
      TempCode = FormLoadCode
      If Len(TempCode) > 0 Then
         .AddCodeToEventProcedure "Load", "Form", TempCode
      End If
      
      TempCode = FormUnLoadCode
      If Len(TempCode) > 0 Then
         .AddCodeToEventProcedure "Unload", "Form", TempCode
      End If
      
   End With

End Sub

Private Sub CheckCodeModuleBehindForm(ByVal frm As Form)

   If Not frm.HasModule Then
      frm.HasModule = True
      DoCmd.Save acForm, frm.Name
   End If

End Sub

Private Property Get DeclarationCode() As String
   DeclarationCode = m_Importer.DeclarationCode
End Property

Private Property Get ClassInitalisationCode( _
                        ByVal KindOfUseFilterMethode As Long, _
                        ByVal UseFilterMethodeSubFormName As String) As String
   
   ClassInitalisationCode = m_Importer.ClassInitalisationCode(GetUseFilterMethodeCode(KindOfUseFilterMethode, UseFilterMethodeSubFormName), GetFilterControlCodeLines)

End Property

Private Function GetUseFilterMethodeCode( _
                        ByVal KindOfUseFilterMethode As Long, _
                        ByVal UseFilterMethodeSubFormName As String) As String
   
   Dim MethodeCode As String
   Dim SufFormRefCode As String
   Dim LinePrefix As String

   Select Case KindOfUseFilterMethode
      Case 0
         LinePrefix = "'"
         SufFormRefCode = " ' oder Me.EinUnterformular.Form"
      Case 1
         LinePrefix = vbNullString
         SufFormRefCode = vbNullString
      Case 2
         LinePrefix = vbNullString
         If Len(UseFilterMethodeSubFormName) = 0 Then
            Err.Raise vbObjectError, "FilterControlManagerImporter", "Bitte Unterformular angeben"
         End If
         SufFormRefCode = "." & UseFilterMethodeSubFormName & ".Form"
   End Select

   MethodeCode = _
         "Private Function UseFilter(ByVal NewFilterString As String)" & vbNewLine & _
         vbNewLine & _
         LinePrefix & "   With Me" & SufFormRefCode & vbNewLine & _
         LinePrefix & "      .Filter = NewFilterString" & vbNewLine & _
         LinePrefix & "      .FilterOn = (Len(.Filter) > 0)" & vbNewLine & _
         LinePrefix & "   End With" & vbNewLine & _
         vbNewLine & _
         "End Function"

   GetUseFilterMethodeCode = MethodeCode

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
