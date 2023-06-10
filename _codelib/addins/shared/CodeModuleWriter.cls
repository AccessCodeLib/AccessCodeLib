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
' Class: CodeModuleWriter
'---------------------------------------------------------------------------------------
'
' Create CodeModules
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleWriter.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/VbeTools.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

#Const VBIDE_EARLYBINDING = 0
#Const DEBUGMODE = 0

#If VBIDE_EARLYBINDING Then
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
'
' CodeModule to be changed
'
'---------------------------------------------------------------------------------------
#If VBIDE_EARLYBINDING Then
Public Property Get CodeModule() As VBIDE.CodeModule
#Else
Public Property Get CodeModule() As Object
#End If

   Set CodeModule = m_CodeModule
   
End Property

#If VBIDE_EARLYBINDING Then
Public Property Set CodeModule(ByVal NewRef As VBIDE.CodeModule)
#Else
Public Property Set CodeModule(ByVal NewRef As Object)
#End If
   Set m_CodeModule = NewRef
End Property

Public Sub Load(ByVal ModuleName As String)
   Set m_CodeModule = CurrentVbProject.VBComponents.Item(ModuleName).CodeModule
End Sub

Public Sub AddCode(ByVal Code As String)
   m_CodeModule.InsertLines m_CodeModule.CountOfDeclarationLines + 1, vbNewLine & Code
End Sub

Public Sub AddCodeToEventProcedure(ByVal EventName As String, ByVal ObjectName As String, ByVal Code As String)

   Dim StartLine As Long
   Dim CodeExists As Boolean
   Dim ProcEndLine As Long
   Dim CodeCheckLineStart As Long
   Dim CodeCheckLineEnd As Long
   
   StartLine = FindEventProcedureBodyLine(EventName, ObjectName, ProcEndLine)

   If StartLine < 0 Then
      StartLine = CreateEmptyEventProcedure(EventName, ObjectName)
   Else
      CodeCheckLineStart = StartLine
      CodeCheckLineEnd = ProcEndLine
      CodeExists = m_CodeModule.Find(Code, CodeCheckLineStart, 0, CodeCheckLineEnd, 0, True)
   End If
   
   If Not CodeExists Then
      m_CodeModule.InsertLines StartLine + 1, vbNewLine & Code
   End If
   
End Sub

Private Function FindEventProcedureBodyLine(ByVal EventName As String, ByVal ObjectName As String, _
                                   Optional ByRef ProcEndLine As Long) As Long

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
      ProcEndLine = FindEventProcedureBodyLine + m_CodeModule.ProcCountLines(ObjectName & "_" & EventName, vbext_pk_Proc) + 1
   Else
      FindEventProcedureBodyLine = -1
   End If

End Function

Private Function CreateEmptyEventProcedure(ByVal EventName As String, ByVal ObjectName As String) As Long
   m_CodeModule.CreateEventProc EventName, ObjectName
   CreateEmptyEventProcedure = m_CodeModule.ProcBodyLine(ObjectName & "_" & EventName, vbext_pk_Proc)
End Function

#If VBIDE_EARLYBINDING Then
Friend Property Set CurrentVbProject(ByVal NewRef As VBIDE.VBProject)
#Else
Friend Property Set CurrentVbProject(ByVal NewRef As Object)
#End If
   Set m_CurrentVbProject = NewRef
End Property

#If VBIDE_EARLYBINDING Then
Friend Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Friend Property Get CurrentVbProject() As Object
#End If

#If VBIDE_EARLYBINDING Then
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VbeTools.CurrentVbProject
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

End Property

Public Sub ImportVbComponent(ByVal ComponentType As vbext_ComponentType, ByVal CodeModulName As String, ByVal ImportFile As String)

#If VBIDE_EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
   Dim cm As VBIDE.CodeModule
#Else
   Dim VbcCol As Object
   Dim vbc As Object
   Dim cm As Object
#End If
   Dim CodeModuleExists As Boolean

   If ComponentType = 0 Then
      ComponentType = GetComponentTypeFromFile(ImportFile)
   End If

   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeModulName Then
         CodeModuleExists = True
         Exit For
      End If
   Next

   If CodeModuleExists Then   ' überprüfen, ob Typen übereinstimmen
      If vbc.Type <> ComponentType Then
         VbcCol.Remove vbc
         Set vbc = Nothing
         CodeModuleExists = False
      End If
   End If
  
   If Not CodeModuleExists Then
      Set vbc = VbcCol.Add(ComponentType)
      vbc.Name = CodeModulName
   End If
   
   Set cm = vbc.CodeModule
   cm.DeleteLines 1, cm.CountOfLines
   cm.AddFromFile ImportFile
   If ComponentType = vbext_ComponentType.vbext_ct_ClassModule Then
      'wegschneiden von:
      'VERSION 1.0 CLASS
      'BEGIN
      '  MultiUse = -1  'True
      'End
      DeleteClassHeaderLinesFromCodeModul cm
   End If
   
   'Leerzeilen am Ende säubern
   Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
      cm.DeleteLines cm.CountOfLines
   Loop

End Sub

Private Function GetComponentTypeFromFile(ByVal FilePath As String) As vbext_ComponentType

   Const SEARCHSTRING_ATTRIBUTNAME_BEGIN As String = "Attribute VB_Name = """

   Dim Pos As Long
   Dim CheckString As String
   Dim FileNumber As Long
   Dim CmType As vbext_ComponentType

   FileNumber = FreeFile
   Open FilePath For Binary Access Read As FileNumber
   CheckString = String$(LOF(FileNumber), 0)
   Get FileNumber, , CheckString
   Close FileNumber
   
   'Pos von VB_Name ermitteln (= Kennzeichen für Standardmodul oder Klasse)
   Pos = InStr(1, CheckString, SEARCHSTRING_ATTRIBUTNAME_BEGIN)
   
   'Typ bestimmen
   If Pos > 0 Then
      CmType = vbext_ComponentType.vbext_ct_StdModule
      If InStr(1, Left$(CheckString, Pos - 1), "CLASS", vbBinaryCompare) Then
         CmType = vbext_ComponentType.vbext_ct_ClassModule
      End If
   Else
      Err.Raise vbObjectError, "", "Codemodule file required (StdModule or ClassModule)"
   End If

   GetComponentTypeFromFile = CmType

End Function

#If VBIDE_EARLYBINDING = 1 Then
Private Sub DeleteClassHeaderLinesFromCodeModul(ByRef cm As CodeModule)
#Else
Private Sub DeleteClassHeaderLinesFromCodeModul(ByRef cm As Object)
#End If
'
'Header looks like this:
'--
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'--

   Dim Target As String
   Dim SearchStartLine As Long, SearchStartCol As Long, _
       SearchEndLine As Long, SearchEndCol As Long
   Dim HeaderStartLine As Long, HeaderEndLine As Long

   SearchStartLine = 1
   SearchStartCol = 1
   SearchEndLine = -1
   SearchEndCol = -1

   Target = "VERSION*CLASS"
   If Not cm.Find(Target, SearchStartLine, SearchStartCol, SearchEndLine, SearchEndCol, True, False, True) Then
      Exit Sub
   End If
   HeaderStartLine = SearchStartLine

   Target = "BEGIN"
   SearchStartLine = SearchEndLine
   SearchStartCol = SearchEndCol
   SearchEndLine = -1
   SearchEndCol = -1
   If Not cm.Find(Target, SearchStartLine, SearchStartCol, SearchEndLine, SearchEndCol, True, False, False) Then
      Exit Sub
   End If

   Target = "End"
   SearchStartLine = SearchEndLine
   SearchStartCol = SearchEndCol
   SearchEndLine = -1
   SearchEndCol = -1
   If Not cm.Find(Target, SearchStartLine, SearchStartCol, SearchEndLine, SearchEndCol, True, False, False) Then
      Exit Sub
   End If
   HeaderEndLine = SearchEndLine

#If DEBUGMODE = 1 Then
   Debug.Print "Header lines:", HeaderStartLine, HeaderEndLine
#End If

   cm.DeleteLines HeaderStartLine, HeaderEndLine

End Sub
