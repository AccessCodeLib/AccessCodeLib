Attribute VB_Name = "modWizardCodeModulesData"
'---------------------------------------------------------------------------------------
' Package: _codelib.addins.FilterFormWizard.modWizardCodeModulesData
'---------------------------------------------------------------------------------------
'
' Helper functions for FilterFormWizard
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/FilterFormWizard/modWizardCodeModulesData.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit
Option Private Module

Public Property Get SvnRev() As String
   
   With CodeDb.OpenRecordset("select max(SvnRev) from usys_AppFiles")
      If Not .EOF Then
         SvnRev = Nz(.Fields(0).Value, 0)
      End If
      .Close
   End With
   
End Property

Public Property Get SvnRevMin() As String
   
   With CodeDb.OpenRecordset("select Min(SvnRev) from usys_AppFiles")
      If Not .EOF Then
         SvnRevMin = Nz(.Fields(0).Value, "0")
      End If
      .Close
   End With
   
End Property
