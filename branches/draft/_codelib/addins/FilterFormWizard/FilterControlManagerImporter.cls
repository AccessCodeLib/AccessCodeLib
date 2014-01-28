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
' Code für FilterControlManager iumportieren
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/FilterControlManagerImporter.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Sub AddToForm(ByVal FormName As String)
   
   Dim frm As Form

   CheckRequiredCodeModules
   
   If CurrentProject.AllForms(FormName).IsLoaded Then
      DoCmd.Close acForm, FormName
   End If

   DoCmd.OpenForm FormName, acDesign, , , , acWindowNormal

   AddFilterControlManagerCodeToFrom Application.Forms(FormName)

   DoCmd.Save acForm, FormName

End Sub

Private Function CheckRequiredCodeModules() As Boolean
   TransferCodeModules "SqlTools", "StringCollection", "FilterStringBuilder", "FilterControlEventBridge", "FilterControl", "FilterControlCollection", "FilterControlManager"
End Function

Private Sub AddFilterControlManagerCodeToFrom(ByVal frm As Form)

#If EARLYBINDING Then
   Dim cm As CodeModule
#Else
   Dim cm As Object
#End If

   CheckCodeModuleBehindForm frm

   With New CodeModuleWriter

      .Load "Form_" & frm.Name

      .AddCode FilterControlManagerDeclarationCode & vbNewLine & _
                           FilterControlManagerInitalisationCode
   
      .AddCodeToEventProcedure "Load", "Form", FormLoadCode
      .AddCodeToEventProcedure "Unload", "Form", FormUnLoadCode
   
   End With

End Sub

Private Sub CheckCodeModuleBehindForm(ByVal frm As Form)

   If Not frm.HasModule Then
      frm.HasModule = True
   End If

End Sub

Private Property Get FilterControlManagerDeclarationCode() As String
   FilterControlManagerDeclarationCode = _
      "Private WithEvents m_FilterControlManager As FilterControlManager" & vbNewLine
End Property

Private Property Get FilterControlManagerInitalisationCode() As String
   
   Dim Code As String

   Code = _
      "Private Sub m_FilterControlManager_FilterStringChanged(ByVal NewFilterString As String)" & vbNewLine & _
      "   UseFilter NewFilterString" & vbNewLine & _
      "End Sub" & vbNewLine

   Code = Code & vbNewLine & _
      "Private Function UseFilter(ByVal NewFilterString As String)" & vbNewLine & _
      vbNewLine & _
      "'   With Me.sfrItems.Form" & vbNewLine & _
      "'      .Filter = NewFilterString" & vbNewLine & _
      "'      .FilterOn = (Len(.Filter) > 0)" & vbNewLine & _
      "'      Me.txtFilterString.Value = .Filter" & vbNewLine & _
      "'   End With" & vbNewLine & _
      vbNewLine & _
      "End Function" & vbNewLine

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
      "'      .Add ""DatenfeldName"", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.Steuerelement" & vbNewLine & _
      "'      .Add ..." & vbNewLine & _
      "'      .Add ..." & vbNewLine & _
      vbNewLine & _
      "   End With" & vbNewLine & _
      vbNewLine & _
      "End Sub" & vbNewLine

   FilterControlManagerInitalisationCode = Code

End Property

Private Property Get FormLoadCode() As String
   FormLoadCode = "   InitFilterControlManager"
End Property

Private Property Get FormUnLoadCode() As String
   FormUnLoadCode = "   Set m_FilterControlManager = Nothing"
End Property
