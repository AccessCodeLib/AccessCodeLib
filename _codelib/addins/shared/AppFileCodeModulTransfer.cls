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
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

'Module u. Klassen übertragen
Public Sub TransferCodeModules(ParamArray ModulNames() As Variant)
   
   Dim i As Long
   Dim ArrSize As Long

   ArrSize = UBound(ModulNames)
   For i = 0 To ArrSize
      CheckCodeModule ModulNames(i), True
   Next

End Sub

'Module u. Klassen erneuern
Public Function ReplaceCodeModules(ParamArray ModulNames() As Variant) As Boolean
   
   Dim i As Long
   Dim ArrSize As Long
   Dim vbp As Object 'VBProject
 
   'VBProject der Anwendung:
   Set vbp = VbeTools.CurrentVbProject

   'Module erneuern:
   If Not (vbp Is Nothing) Then
      ArrSize = UBound(ModulNames)
      For i = 0 To ArrSize
      
         If CheckCodeModule(ModulNames(i)) Then
            'Modul löschen
            vbp.VBComponents.Remove vbp.VBComponents(ModulNames(i))
         End If
         
         'Module kopieren:
         CheckCodeModule ModulNames(i), True
         
      Next
      ReplaceCodeModules = True
   End If
   
   Set vbp = Nothing

End Function


'Module u. Klassen auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModules(ParamArray ModulNames() As Variant) As Boolean
   
   Dim i As Long
   Dim ModulesExists As Boolean
   Dim ArrSize As Long

   ArrSize = UBound(ModulNames)
   ModulesExists = True
   For i = 0 To ArrSize
      ModulesExists = ModulesExists And CheckCodeModule(ModulNames(i), False)
   Next
   
   CheckCodeModules = ModulesExists

End Function

'Modul oder Klasse auf Existenz prüfen. Es erfolgt keine inhaltliche Prüfung!
Public Function CheckCodeModule(ByVal ModulName As String, _
                       Optional ByVal TransferMissingModule As Boolean = False) As Boolean
   
   Dim rst As DAO.Recordset
   Dim IsMissing As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & ModulName & "' and O.Type=-32761", dbOpenForwardOnly, dbReadOnly)
   IsMissing = rst.EOF
   rst.Close
   Set rst = Nothing
   
   If IsMissing And TransferMissingModule Then
      TransferCodeModul CurrentProject, acModule, ModulName
      IsMissing = False
   End If
   
   CheckCodeModule = Not IsMissing

End Function

Private Sub TransferCodeModul(ByVal TargetProject As Access.CurrentProject, _
                              ByVal ObjType As AcObjectType, ByVal ModulName As String)

   Dim FileName As String
   
   FileName = FileTools.GetNewTempFileName
   CurrentApplication.CreateAppFile ModulName, FileName
   TargetProject.Application.LoadFromText ObjType, ModulName, FileName
   Kill FileName
   
End Sub
