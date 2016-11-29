Attribute VB_Name = "modWizardCodeModulesData"
'---------------------------------------------------------------------------------------
' Modul: defGlobal_FilterFormWizard
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsfunktionen für FilterFormWizard
' </summary>
' <remarks>
' </remarks>
' \ingroup ACLibAddInFilterFormWizard
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/modWizardCodeModulesData.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Property Get SvnRev() As Long
   
   With CodeDb.OpenRecordset("select max(SvnRev) from usys_AppFiles")
      If Not .EOF Then
         SvnRev = Nz(.Fields(0).Value, 0)
      End If
      .Close
   End With
   
End Property
