VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AppFileCodeModulTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klasse: AppFileCodeModulTransfer
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Codemodule aus AppFile-Tabelle importieren
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/AppFileCodeModulTransfer.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler_AppFile.cls</use>
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Module u. Klassen übertragen
Public Sub TransferCodeModules(ParamArray sModulName() As Variant)
   
   Dim i As Long
   Dim arrSize As Long

   arrSize = UBound(sModulName)
   For i = 0 To arrSize
      CheckCodeModule sModulName(i), True
   Next

End Sub

'Module u. Klassen erneuern
Public Function ReplaceCodeModules(ParamArray sModulName() As Variant) As Boolean
   
   Dim i As Long
   Dim arrSize As Long
   Dim vbp As Object 'VBProject
 
   'VBProject der Anwendung:
   Set vbp = VBETools.CurrentVbProject

   'Module erneuern:
   If Not (vbp Is Nothing) Then
      arrSize = UBound(sModulName)
      For i = 0 To arrSize
      
         If CheckCodeModule(sModulName(i)) Then
            'Modul löschen
            vbp.VBComponents.Remove vbp.VBComponents(sModulName(i))
         End If
         
         'Module kopieren:
         CheckCodeModule sModulName(i), True
         
      Next
      ReplaceCodeModules = True
   End If
   
   Set vbp = Nothing

End Function


'Module u. Klassen auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModules(ParamArray sModulName() As Variant) As Boolean
   
   Dim i As Long
   Dim bolModulesExists As Boolean
   Dim arrSize As Long

   arrSize = UBound(sModulName)
   bolModulesExists = True
   For i = 0 To arrSize
      bolModulesExists = bolModulesExists And CheckCodeModule(sModulName(i), False)
   Next
   
   CheckCodeModules = bolModulesExists

End Function

'Modul oder Klasse auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModule(ByVal sModulName As String, _
                       Optional ByVal TransferMissingModule As Boolean = False) As Boolean
   
   Dim rst As DAO.Recordset
   Dim bolMissingModule As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & sModulName & "' and O.Type=-32761", dbOpenForwardOnly, dbReadOnly)
   bolMissingModule = rst.EOF
   rst.Close
   Set rst = Nothing
   
   If bolMissingModule And TransferMissingModule Then
      TransferCodeModul CurrentProject, acModule, sModulName
      bolMissingModule = False
   End If
   
   CheckCodeModule = Not bolMissingModule

End Function

Private Sub TransferCodeModul(ByRef targetProject As Access.CurrentProject, ByVal ObjType As AcObjectType, ByVal sModulName As String)

   Dim strFileName As String
   
   strFileName = FileTools.GetNewTempFileName
   CurrentApplication.CreateAppFile sModulName, strFileName
   targetProject.Application.LoadFromText ObjType, sModulName, strFileName
   Kill strFileName
   
End Sub
