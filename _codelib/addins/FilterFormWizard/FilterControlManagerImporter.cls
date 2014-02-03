VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FilterControlManagerImporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: FilterControlManagerImporter
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
'  <file>_codelib/addins/FilterFormWizard/FilterControlManagerImporter.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'  <use>_codelib/addins/shared/AppFileCodeModulTransfer.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_FilterControlLines As StringCollection

Private Sub Class_Initialize()
   Set m_FilterControlLines = New StringCollection
End Sub


Friend Property Get Self() As FilterControlManagerImporter
   Set Self = Me
End Property

Friend Sub WriteToForm(ByVal FormName As String, _
              Optional ByVal KindOfUseFilterMethode As Long = 0, Optional ByVal UseFilterMethodeSubFormName As String)
   
   Dim frm As Form

   CheckRequiredCodeModules
   
   If CurrentProject.AllForms(FormName).IsLoaded Then
      DoCmd.Close acForm, FormName
   End If

   DoCmd.OpenForm FormName, acDesign, , , , acWindowNormal

   AddFilterControlManagerCodeToFrom Application.Forms(FormName), KindOfUseFilterMethode, UseFilterMethodeSubFormName

   DoCmd.Save acForm, FormName

End Sub

Private Function CheckRequiredCodeModules() As Boolean
   With New AppFileCodeModulTransfer
      .TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager"
   End With
End Function

Private Sub AddFilterControlManagerCodeToFrom(ByVal frm As Form, _
                                              ByVal KindOfUseFilterMethode As Long, _
                                              ByVal UseFilterMethodeSubFormName As String)

#If EARLYBINDING Then
   Dim cm As CodeModule
#Else
   Dim cm As Object
#End If

   CheckCodeModuleBehindForm frm

   With New CodeModuleWriter

      .Load "Form_" & frm.Name

      .AddCode FilterControlManagerDeclarationCode & vbNewLine & _
               FilterControlManagerInitalisationCode(KindOfUseFilterMethode, UseFilterMethodeSubFormName)
   
      .AddCodeToEventProcedure "Load", "Form", FormLoadCode
      .AddCodeToEventProcedure "Unload", "Form", FormUnLoadCode
   
   End With

End Sub

Private Sub CheckCodeModuleBehindForm(ByVal frm As Form)

   If Not frm.HasModule Then
      frm.HasModule = True
      DoCmd.Save acForm, frm.Name
   End If

End Sub

Private Property Get FilterControlManagerDeclarationCode() As String
   FilterControlManagerDeclarationCode = _
      "Private WithEvents m_FilterControlManager As FilterControlManager" & vbNewLine
End Property

Private Property Get FilterControlManagerInitalisationCode( _
                        ByVal KindOfUseFilterMethode As Long, _
                        ByVal UseFilterMethodeSubFormName As String) As String
   
   Dim Code As String

   Code = _
      "Private Sub m_FilterControlManager_FilterStringChanged(ByVal NewFilterString As String)" & vbNewLine & _
      "   UseFilter NewFilterString" & vbNewLine & _
      "End Sub" & vbNewLine

   Code = Code & vbNewLine & _
      GetUseFilterMethodeCode(KindOfUseFilterMethode, UseFilterMethodeSubFormName) & vbNewLine

   Code = Code & vbNewLine & _
      "Private Sub InitFilterControlManager()" & vbNewLine & _
      "   Set m_FilterControlManager = New FilterControlManager" & vbNewLine & _
      "'   m_FilterControlManager.ConfigSqlFormat ""\#yyyy-mm-dd\#"", ""True"", ""*""" & vbNewLine & _
      "   InitFilterControls" & vbNewLine & _
      "   m_FilterControlManager.AutoFilterOn = True 'oder z. B. mittels Checkbox: = Nz(Me.cbAutoFilter.Value, False)" & vbNewLine & _
      "End Sub" & vbNewLine

   Code = Code & vbNewLine & _
      "Private Property Get FilterControlManager() As FilterControlManager" & vbNewLine & _
      "   If m_FilterControlManager Is Nothing Then InitFilterControlManager" & vbNewLine & _
      "   Set FilterControlManager = m_FilterControlManager" & vbNewLine & _
      "End Property" & vbNewLine

   Code = Code & vbNewLine & _
      "Private Sub InitFilterControls()" & vbNewLine & _
      vbNewLine & _
      "   With FilterControlManager.FilterControls" & vbNewLine & _
      vbNewLine & _
      GetFilterControlCodeLines & vbNewLine & _
      vbNewLine & _
      "   End With" & vbNewLine & _
      vbNewLine & _
      "End Sub" & vbNewLine

   FilterControlManagerInitalisationCode = Code

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
   FormLoadCode = "   InitFilterControlManager"
End Property

Private Property Get FormUnLoadCode() As String
   FormUnLoadCode = "   Set m_FilterControlManager = Nothing"
End Property

Friend Sub AddFilterControlDefinition(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
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
   
   Dim fcLine As String

   CheckControlRefCode Control

   fcLine = "Add """ & DataField & """, " & DataType & ", " & RelationalOperator & ", " & Control
   
   If Len(Control2) > 0 Then
      CheckControlRefCode Control2
      fcLine = fcLine & ", " & Control2
   End If

   GetFilterControlCodeLine = fcLine

End Function

Private Function GetMultiFilterControlCodeLine(ByVal DataField As String, ByVal DataType As String, ByVal RelationalOperator As String, _
                                         ByRef ControlArray() As String) As String

   Dim fcLine As String

   Dim i As Long
   For i = LBound(ControlArray) To UBound(ControlArray)
      CheckControlRefCode ControlArray(i)
   Next
   
   fcLine = "AddMultiControlCriteria """ & DataField & """, " & DataType & ", " & RelationalOperator & ", Null, " & Join(ControlArray, ", ")
   
   GetMultiFilterControlCodeLine = fcLine

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

Private Sub CleanControlArrayString(ByRef CtlString As String)
   Do While InStr(1, CtlString, ", ") > 0
      CtlString = Replace(CtlString, ", ", ",")
   Loop
End Sub
