Attribute VB_Name = "modWizardCodeModulesData"
'---------------------------------------------------------------------------------------
' Package: _codelib.addins.shared.modWizardCodeModulesData
'---------------------------------------------------------------------------------------
'
' SCC file data in usys_AppFiles
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/modWizardCodeModulesData.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit
Option Private Module

Public Property Get SccRev() As String
   
   With CodeDb.OpenRecordset("select max(SccRev) from usys_AppFiles")
      If Not .EOF Then
         SccRev = Nz(.Fields(0).Value, 0)
      End If
      .Close
   End With
   
End Property

Public Property Get SccRevMin() As String
   
   With CodeDb.OpenRecordset("select Min(SccRev) from usys_AppFiles")
      If Not .EOF Then
         SccRevMin = Nz(.Fields(0).Value, "0")
      End If
      .Close
   End With
   
End Property
