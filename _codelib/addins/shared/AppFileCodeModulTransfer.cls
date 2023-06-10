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
' Class: _codelib.addins.shared.AppFileCodeModulTransfer
'---------------------------------------------------------------------------------------
'
' Import codemodules from AppFile table
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/AppFileCodeModulTransfer.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleWriter.cls</use>
'  <use>_codelib/addins/shared/VbeTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0

#If EARLYBINDING = 1 Then
Private m_CurrentVbProject As VBIDE.VBProject
#Else
Private m_CurrentVbProject As Object
#End If

Private m_CodeModuleWriter As CodeModuleWriter
Private m_UseVbComponentsImport As Boolean

'---------------------------------------------------------------------------------------
' Property: UseVbComponentsImport
'---------------------------------------------------------------------------------------
Public Property Get UseVbComponentsImport() As Boolean
   UseVbComponentsImport = m_UseVbComponentsImport
End Property

Public Property Let UseVbComponentsImport(ByVal NewValue As Boolean)
   m_UseVbComponentsImport = NewValue
End Property


'---------------------------------------------------------------------------------------
' Sub: TransferCodeModules
'---------------------------------------------------------------------------------------
'
' Transfer modules & classes
'
' Parameters:
'     CodeModulNames - (ParamArray) Name of CodeModule(s) to transfer
'
'---------------------------------------------------------------------------------------
Public Sub TransferCodeModules(ParamArray CodeModulNames() As Variant)
   
   Dim i As Long
   Dim ArrSize As Long

   ArrSize = UBound(CodeModulNames)
   For i = 0 To ArrSize
      CheckCodeModule CodeModulNames(i), True
   Next

End Sub

'---------------------------------------------------------------------------------------
' Function: ReplaceCodeModules
'---------------------------------------------------------------------------------------
'
' Renew modules & classes
'
' Parameters:
'     CodeModulNames - (ParamArray) Name of CodeModule(s) to renew
'
'---------------------------------------------------------------------------------------
Public Function ReplaceCodeModules(ParamArray CodeModulNames() As Variant) As Boolean
   
   Dim i As Long
   Dim ArrSize As Long
   Dim vbp As Object 'VBProject
 
   'VBProject of the application:
   Set vbp = VbeTools.CurrentVbProject

   'Renew modules:
   If Not (vbp Is Nothing) Then
      ArrSize = UBound(CodeModulNames)
      For i = 0 To ArrSize
      
         If CheckCodeModule(CodeModulNames(i)) And (UseVbComponentsImport = False) Then
            'Modul löschen
            vbp.VBComponents.Remove vbp.VBComponents(CodeModulNames(i))
         End If
         
         'Module aktualisieren
         TransferCodeModul CurrentProject, acModule, CodeModulNames(i)
         
      Next
      ReplaceCodeModules = True
   End If
   
   Set vbp = Nothing

End Function

'---------------------------------------------------------------------------------------
' Function: CheckCodeModules
'---------------------------------------------------------------------------------------
'
' Check modules and classes for existence.
'
' Parameters:
'     CodeModulNames - (ParamArray) Name of CodeModule(s) to check
'
' Remarks:
'     There is no content check.
'
'---------------------------------------------------------------------------------------
Public Function CheckCodeModules(ParamArray CodeModulNames() As Variant) As Boolean
   
   Dim i As Long
   Dim ModulesExists As Boolean
   Dim ArrSize As Long

   ArrSize = UBound(CodeModulNames)
   ModulesExists = True
   For i = 0 To ArrSize
      ModulesExists = ModulesExists And CheckCodeModule(CodeModulNames(i), False)
   Next
   
   CheckCodeModules = ModulesExists

End Function

'---------------------------------------------------------------------------------------
' Function: CheckCodeModule
'---------------------------------------------------------------------------------------
'
' Check codemodule for existence.
'
' Parameters:
'     CodeModulName           - Name of CodeModule to check
'     TransferMissingModule   - Insert CodeModule if missing
'
' Remarks:
'     There is no content check.
'
'---------------------------------------------------------------------------------------
Public Function CheckCodeModule(ByVal CodeModulName As String, _
                       Optional ByVal TransferMissingModule As Boolean = False) As Boolean
   
   Dim rst As DAO.Recordset
   Dim IsMissing As Boolean

   Set rst = CurrentDb.OpenRecordset("Select O.Name from MSysObjects O where O.Name = '" & CodeModulName & "' and O.Type=-32761", dbOpenForwardOnly, dbReadOnly)
   IsMissing = rst.EOF
   rst.Close
   Set rst = Nothing
   
   If IsMissing And TransferMissingModule Then
      TransferCodeModul CurrentProject, acModule, CodeModulName
      IsMissing = False
   End If
   
   CheckCodeModule = Not IsMissing

End Function

'---------------------------------------------------------------------------------------
' Sub: TransferCodeModul
'---------------------------------------------------------------------------------------
'
' Check codemodule for existence.
'
' Parameters:
'     TargetProject  - Access.CurrentProject of target application
'     ObjType        - AcObjectType of CodeModule to transfer (e. g. acModule or acForm)
'     CodeModulName  - Name of CodeModule to transfer
'
'---------------------------------------------------------------------------------------
Public Sub TransferCodeModul(ByVal TargetProject As Access.CurrentProject, _
                             ByVal ObjType As AcObjectType, ByVal CodeModulName As String)

   Dim FileName As String
   
   FileName = FileTools.GetNewTempFileName
   CurrentApplication.CreateAppFile CodeModulName, FileName
   If UseVbComponentsImport And ObjType = acModule Then
      If VbeTools.CodeModuleExists(CodeModulName) Then
         CurrentCodeModulWriter.ImportVbComponent 0, CodeModulName, FileName
         VbeTools.ActivateCurrentProject
         On Error Resume Next ' Bei mde-Add-In weigert sich Access manchmal die eigentliche Anwendung zu kompilieren
         TargetProject.Application.DoCmd.RunCommand acCmdCompileAndSaveAllModules
         If Err.Number <> 0 Then
            Err.Clear
         End If
         On Error GoTo 0
      Else
         ConvertToSaveAsTextFile FileName
         TargetProject.Application.LoadFromText ObjType, CodeModulName, FileName
      End If
   Else
      TargetProject.Application.LoadFromText ObjType, CodeModulName, FileName
   End If
   Kill FileName
   
   
End Sub

Private Sub ConvertToSaveAsTextFile(ByVal FilePath As String)
   
   Const SEARCHSTRING_ATTRIBUTNAME_BEGIN As String = "Attribute VB_Name = """

   Dim Pos As Long
   Dim CheckString As String
   Dim NewFile As String
   Dim FileNumber As Long

   FileNumber = FreeFile
   Open FilePath For Binary Access Read As FileNumber
   CheckString = String$(LOF(FileNumber), 0)
   Get FileNumber, , CheckString
   Close FileNumber
   
   Pos = InStr(1, CheckString, SEARCHSTRING_ATTRIBUTNAME_BEGIN)
   If Pos <= 1 Then
      Exit Sub
   End If
   
   CheckString = Mid(CheckString, Pos)
   
   NewFile = FilePath & "_SaT"
   FileNumber = FreeFile
   Open NewFile For Binary Access Write As FileNumber
   Put FileNumber, , CheckString
   Close FileNumber
   
   Kill FilePath
   Name NewFile As FilePath
   
End Sub

Private Property Get CurrentCodeModulWriter() As CodeModuleWriter
   If m_CodeModuleWriter Is Nothing Then
      Set m_CodeModuleWriter = New CodeModuleWriter
   End If
   Set CurrentCodeModulWriter = m_CodeModuleWriter
End Property
