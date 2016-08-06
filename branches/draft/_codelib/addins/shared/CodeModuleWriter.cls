VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: CodeModuleWriter
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Inhalt von Code-Modul auswerten
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleWriter.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const EARLYBINDING = 0

#If EARLYBINDING Then
   Private m_CodeModule As VBIDE.CodeModule
   Private m_CurrentVbProject As VBIDE.VBProject
#Else
   Private m_CodeModule As Object
   Private m_CurrentVbProject As Object
   
   Public Enum vbext_ProcKind
      vbext_pk_Proc = 0
      vbext_pk_Let = 1
      vbext_pk_Set = 2
      vbext_pk_Get = 3
   End Enum
   
   Public Enum vbext_ComponentType
      vbext_ct_StdModule = 1
      vbext_ct_ClassModule = 2
      vbext_ct_MSForm = 3
      vbext_ct_ActiveXDesigner = 11
      vbext_ct_Document = 100
   End Enum
#End If

'---------------------------------------------------------------------------------------
' Property: CodeModule
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Code-Modul, das geändert werden soll
' </summary>
' <returns>VBIDE.CodeModule</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Get CodeModule() As VBIDE.CodeModule
#Else
Public Property Get CodeModule() As Object
#End If

   Set CodeModule = m_CodeModule
   
End Property

'---------------------------------------------------------------------------------------
' Property: CodeModule
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Code-Modul, das geändert werden soll
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Set CodeModule(ByVal NewRef As VBIDE.CodeModule)
#Else
Public Property Set CodeModule(ByVal NewRef As Object)
#End If
   Set m_CodeModule = NewRef
End Property

Public Sub Load(ByVal ModuleName As String)
   Set m_CodeModule = CurrentVbProject.VBComponents.Item(ModuleName).CodeModule
End Sub

#If EARLYBINDING Then
Public Sub AddCode(ByVal Code As String)
#Else
Public Sub AddCode(ByVal Code As String)
#End If
   m_CodeModule.InsertLines m_CodeModule.CountOfDeclarationLines + 1, vbNewLine & Code
End Sub

#If EARLYBINDING Then
Public Sub AddCodeToEventProcedure(ByVal EventName As String, ByVal ObjectName As String, ByVal Code As String)
#Else
Public Sub AddCodeToEventProcedure(ByVal EventName As String, ByVal ObjectName As String, ByVal Code As String)
#End If
   Dim StartLine As Long
   
   StartLine = FindEventProcedureBodyLine(EventName, ObjectName)

   If StartLine < 0 Then
      StartLine = CreateEmptyEventProcedure(EventName, ObjectName)
   End If

   m_CodeModule.InsertLines StartLine + 1, vbNewLine & Code

End Sub

#If EARLYBINDING Then
Private Function FindEventProcedureBodyLine(ByVal EventName As String, ByVal ObjectName As String) As Long
#Else
Private Function FindEventProcedureBodyLine(ByVal EventName As String, ByVal ObjectName As String) As Long
#End If
   Dim StartLine As Long
   Dim EndLine As Long

   Dim ProcedureDeclarationCode As String

   ProcedureDeclarationCode = "Sub " & ObjectName & "_" & EventName & "("

   StartLine = m_CodeModule.CountOfDeclarationLines
   If m_CodeModule.Find(ProcedureDeclarationCode, StartLine, 0, EndLine, 500, False, False, False) = False Then
      StartLine = -1
   End If

   If StartLine > 0 Then
      FindEventProcedureBodyLine = m_CodeModule.ProcBodyLine(ObjectName & "_" & EventName, vbext_pk_Proc)
   Else
      FindEventProcedureBodyLine = -1
   End If

End Function

#If EARLYBINDING Then
Private Function CreateEmptyEventProcedure(ByVal EventName As String, ByVal ObjectName As String) As Long
#Else
Private Function CreateEmptyEventProcedure(ByVal EventName As String, ByVal ObjectName As String) As Long
#End If
   m_CodeModule.CreateEventProc EventName, ObjectName
   CreateEmptyEventProcedure = m_CodeModule.ProcBodyLine(ObjectName & "_" & EventName, vbext_pk_Proc)
End Function

#If EARLYBINDING Then
Friend Property Set CurrentVbProject(ByVal NewRef As VBIDE.VBProject)
#Else
Friend Property Set CurrentVbProject(ByVal NewRef As Object)
#End If
   Set m_CurrentVbProject = NewRef
End Property

#If EARLYBINDING Then
Friend Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Friend Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = UNCPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each Proj In VBE.VBProjects
            If Proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = Proj
               Exit For
            End If
         Next
      End If
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

End Property
