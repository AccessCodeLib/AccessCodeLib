VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeModuleReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: _codelib.addins.shared.CodeModuleReader
'---------------------------------------------------------------------------------------
'
' Evaluate content of CodeModule
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/CodeModuleReader.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleProcedure.cls</use>
'  <use>_codelib/addins/shared/CodeModuleHeaderItem.cls</use>
'  <use>file/FileTools.bas</use>
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

Private m_CodeModuleText As String
Private m_ProcCollection As Collection
Private m_HeaderItemCollection As Collection
Private m_RequiredModules As Collection
Private m_RequiredModulesChecked As Boolean

Private Const ERRNUMBER_CHECKDEPENDENCY_NOTRUN As Long = vbObjectError + 5001
Private Const ERRDESCRIPTION_CHECKDEPENDENCY_NOTRUN As String = "CheckDependency wurde noch nicht ausgeführt"


'---------------------------------------------------------------------------------------
' Property: CodeModule
'---------------------------------------------------------------------------------------
'
' Code module whose text is to be evaluated
'
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Get CodeModule() As VBIDE.CodeModule
#Else
Public Property Get CodeModule() As Object
#End If

   Set CodeModule = m_CodeModule
   
End Property

#If EARLYBINDING Then
Public Property Set CodeModule(ByRef NewRef As VBIDE.CodeModule)
#Else
Public Property Set CodeModule(ByRef NewRef As Object)
#End If

   Set m_CodeModule = NewRef

End Property

'---------------------------------------------------------------------------------------
' Function: Copy
'---------------------------------------------------------------------------------------
'
' Copy instance
'
' Returns:
'     CodeModuleReader
'
'---------------------------------------------------------------------------------------
Public Function Copy() As CodeModuleReader

   Dim NewReader As CodeModuleReader

   Set NewReader = New CodeModuleReader
   Set NewReader.CodeModule = Me.CodeModule
   Set Copy = NewReader

End Function

'---------------------------------------------------------------------------------------
' Property: Name
'---------------------------------------------------------------------------------------
'
' Name of VBComponent of the code module
'
' Remarks:
'     Name = CodeModule.Parent.Name
'
'---------------------------------------------------------------------------------------
Public Property Get Name() As String

   Name = m_CodeModule.Parent.Name

End Property

'---------------------------------------------------------------------------------------
' Property: RequiredModules
'---------------------------------------------------------------------------------------
'
' Collection of the required modules
'
' Returns:
'     Collection
'
'---------------------------------------------------------------------------------------
Public Property Get RequiredModules() As Collection

   If Not m_RequiredModulesChecked Then
      Err.Raise ERRNUMBER_CHECKDEPENDENCY_NOTRUN, "RequiredModules", ERRDESCRIPTION_CHECKDEPENDENCY_NOTRUN
   End If

   Set RequiredModules = m_RequiredModules
 
End Property

'---------------------------------------------------------------------------------------
' Function: CheckDependency
'---------------------------------------------------------------------------------------
'
' Abhängigkeiten prüfen
'
' Parameters:
'     VbpRef                  - VBProject, with Nothing active VBProject is used
'     ModulReaderCollection   - Collection of possible dependencies, with Nothing all code modules from VBProject are used
'
' Returns:
'     Anzahl der Abhängigkeiten
'
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Function CheckDependency(Optional ByVal VbpRef As VBIDE.VBProject = Nothing, _
                                Optional ByRef ModulReaderCollection As Collection = Nothing) As Long
#Else
Public Function CheckDependency(Optional ByVal VbpRef As Object = Nothing, _
                                Optional ByRef ModulReaderCollection As Collection = Nothing) As Long
#End If

   Dim CheckCollection As Collection
   Dim AddModule As Boolean

#If EARLYBINDING Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If

   Dim TempModuleReader As CodeModuleReader
   Dim TempProc As CodeModuleProcedure
   Dim TempHeaderItem As CodeModuleHeaderItem

   Set m_RequiredModules = Nothing
   Set m_RequiredModules = New Collection
   
   If ModulReaderCollection Is Nothing Then
   
      Set ModulReaderCollection = New Collection
      
      If VbpRef Is Nothing Then
         Set VbpRef = CurrentVbProject
      End If

      For Each vbc In VbpRef.VBComponents
         Set TempModuleReader = New CodeModuleReader
         Set TempModuleReader.CodeModule = vbc.CodeModule
         ModulReaderCollection.Add TempModuleReader, TempModuleReader.Name
      Next

   End If
   
   For Each TempModuleReader In ModulReaderCollection
      If Not (TempModuleReader.Name = Me.Name) Then
         AddModule = False
      
         'For a standard module, the procedures and the module constants/variables must be checked
         If TempModuleReader.CodeModule.Parent.Type = vbext_ComponentType.vbext_ct_StdModule Then
            Set CheckCollection = TempModuleReader.Procedures
            For Each TempProc In CheckCollection
               If Me.ProcedureUsed(TempProc) Then
                  AddModule = True
                  Exit For
               End If
            Next
            
            If Not AddModule Then ' Variablen, Konstanten, ... prüfen
               Set CheckCollection = TempModuleReader.HeaderItems
               For Each TempHeaderItem In CheckCollection
                  If Me.HeaderItemUsed(TempHeaderItem) Then
                     AddModule = True
                     Exit For
                  End If
               Next
            End If
            
         Else 'for classes it is enough to check the usage of the class
            If Me.ClassUsed(TempModuleReader.CodeModule.Parent.Name) Then
               AddModule = True
            End If
            
            If Not AddModule Then ' check only enums for classes
               Set CheckCollection = TempModuleReader.HeaderItems
               For Each TempHeaderItem In CheckCollection
                  If TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Enum Then
                     If Me.HeaderItemUsed(TempHeaderItem) Then
                        AddModule = True
                        Exit For
                     End If
                  End If
               Next
            End If
            
         End If
         
         If AddModule Then
            m_RequiredModules.Add TempModuleReader.Copy, TempModuleReader.Name
            'can there be a circular reference here if tempModuleReader reference is stored?
            ' => create a copy for security
         End If
         
      End If
   Next
   
   m_RequiredModulesChecked = True
   CheckDependency = m_RequiredModules.Count

End Function


'---------------------------------------------------------------------------------------
' Function: ClassUsed
'---------------------------------------------------------------------------------------
'
' Checks if class is used/needed in code module
'
' Parameters:
'     ClassName - Class name
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function ClassUsed(ByVal ClassName As String) As Boolean
   
   Dim SearchString(5) As String
   Dim i As Long
   Dim IsUsed As Boolean
   Dim MaxCnt As Long

   SearchString(0) = " As " & Trim$(ClassName) & vbNewLine
   SearchString(1) = " As " & Trim$(ClassName) & ","
   SearchString(2) = " As " & Trim$(ClassName) & " "
   
   MaxCnt = 2
   For i = 0 To MaxCnt
      SearchString(MaxCnt + 1 + i) = " New " & Mid(SearchString(i), 5)
   Next
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   For i = LBound(SearchString) To UBound(SearchString)
      If m_CodeModuleText Like "*" & SearchString(i) & "*" Then
         IsUsed = True
         Exit For
      End If
   Next
   
   ClassUsed = IsUsed

End Function

'---------------------------------------------------------------------------------------
' Function: ProcedureUsed
'---------------------------------------------------------------------------------------
'
' Checks if procedure is used/needed in code module
'
' Parameters:
'     CodeModuleProc - CodeModuleProcedure reference
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function ProcedureUsed(ByRef CodeModuleProc As CodeModuleProcedure) As Boolean
   
   Dim SearchStringArray() As String
   Dim i As Long
   Dim IsUsed As Boolean
   Dim Proc As CodeModuleProcedure
   Dim OnlyWithModulePrefix As Boolean

   'if internal procedure names overlap with the one to be checked,
   'a match is considered a hit only for the combination modulename.procedure
   For Each Proc In Me.Procedures(True, True)
      If Proc.Name = CodeModuleProc.Name Then
         OnlyWithModulePrefix = True
         Exit For
      End If
   Next

   SearchStringArray = CodeModuleProc.SearchStrings(OnlyWithModulePrefix)
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   'Text durchsuchen
   For i = LBound(SearchStringArray) To UBound(SearchStringArray)
#If USEREGEX = 1 Then
      With RegEx
         .Pattern = SearchStringArray(i)
         .Global = False
         IsUsed = .test(m_CodeModuleText)
         If IsUsed Then
            Exit For
         End If
      End With
#Else
      If m_CodeModuleText Like "*" & SearchStringArray(i) & "*" Then
         IsUsed = True
         Exit For
      End If
#End If
   Next
   
   ProcedureUsed = IsUsed

End Function

'---------------------------------------------------------------------------------------
' Function: HeaderItemUsed
'---------------------------------------------------------------------------------------
'
' Prüft, ob CodeModuleHeaderItem (Variable, Konstante, ...) im Code-Modul verwendet/benötigt wird
'
' Parameters:
'     HdrItm - CodeModuleHeaderItem reference
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function HeaderItemUsed(ByRef HdrItm As CodeModuleHeaderItem) As Boolean
   
   Dim SearchStringArray() As String
   Dim i As Long
   Dim IsUsed As Boolean
   
   '/**
   ' * @bug If a public variable overlaps with an internal (private variable in the header or in a proc) of the code module to be searched, the public variable is considered a hit
   '**/

   SearchStringArray = HdrItm.SearchStrings
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   For i = LBound(SearchStringArray) To UBound(SearchStringArray)
#If USEREGEX = 1 Then
      With RegEx
         .Pattern = SearchStringArray(i)
         .Global = False
         IsUsed = .test(m_CodeModuleText)
         If IsUsed Then Exit For
      End With
#Else
      If m_CodeModuleText Like "*" & SearchStringArray(i) & "*" Then
         IsUsed = True
         Exit For
      End If
#End If
   Next

   HeaderItemUsed = IsUsed

End Function

'---------------------------------------------------------------------------------------
' Property: PlainCodeText
'---------------------------------------------------------------------------------------
'
' Returns cleaned code of the code module
'
' Returns:
'     String
'
' Remarks:
'     Removes comment lines
'
'---------------------------------------------------------------------------------------
Public Property Get PlainCodeText() As String
   
   Dim strCodeText As String
   Dim strCodeLine As String
   Dim i As Long

   For i = 1 To m_CodeModule.CountOfLines
      strCodeText = strCodeText & Trim$(GetPlainCodeLine(m_CodeModule.Lines(i, 1))) & vbNewLine
   Next
   
   PlainCodeText = strCodeText

End Property

Private Function GetPlainCodeLine(ByVal CodeLine As String) As String

   Dim PosSingleQuote As Long
   Dim PosDoubleQuote As Long

   PosSingleQuote = InStr(1, CodeLine, "'")
   PosDoubleQuote = InStr(1, CodeLine, """")
   
   If PosSingleQuote = 0 Then
      GetPlainCodeLine = CodeLine
   Else
      If PosDoubleQuote = 0 Or PosSingleQuote < PosDoubleQuote Then
         GetPlainCodeLine = Left$(CodeLine, PosSingleQuote - 1)
      Else ' posSingleQuote > posDoubleQuote and posDoubleQuote > 0 Then
         'closing quotation mark:
         PosDoubleQuote = InStr(PosDoubleQuote + 1, CodeLine, """")
         If PosDoubleQuote > 0 Then
            GetPlainCodeLine = Left$(CodeLine, PosDoubleQuote) & _
                            GetPlainCodeLine(Mid$(CodeLine, PosDoubleQuote + 1))
         Else ' can not actually exist
            GetPlainCodeLine = CodeLine
         End If
      End If
   End If

End Function

'---------------------------------------------------------------------------------------
' Property: Procedures
'---------------------------------------------------------------------------------------
'
' Collection of all procedures (CodeModuleProcedure references)
'
' Parameters:
'     Refresh              - True = rebuild collection; False = if a collection has already been determined, it will be used and not checked again
'     IncludePrivateProc   - True = incl. procedures declared as private
'
' Returns
'     Collection - CodeModuleProcedure references
'
'---------------------------------------------------------------------------------------
Public Property Get Procedures(Optional ByVal Refresh As Boolean = False, _
                               Optional ByVal IncludePrivateProc As Boolean = False) As Collection

   Static IncludePrivate As Boolean
   
   If IncludePrivateProc <> IncludePrivate Then
      Refresh = True
      IncludePrivate = IncludePrivateProc
   End If
   
   If Refresh Then
      Set m_ProcCollection = Nothing
   ElseIf Not m_ProcCollection Is Nothing Then
      Set Procedures = m_ProcCollection
      Exit Property
   End If
   
   Set m_ProcCollection = New Collection

   Dim CurrentLine As Long
   Dim TempProc As CodeModuleProcedure
   Dim TempProcName As String
   Dim TempProcKind As vbext_ProcKind
   Dim AddProc As Boolean
   Dim TempProcLine As String

   CurrentLine = m_CodeModule.CountOfDeclarationLines + 1
   Do While CurrentLine <= m_CodeModule.CountOfLines
      TempProcKind = vbext_pk_Proc
      TempProcName = m_CodeModule.ProcOfLine(CurrentLine, TempProcKind)
      If Len(TempProcName) > 0 Then
         AddProc = True
         For Each TempProc In m_ProcCollection
            If TempProc.Name = TempProcName Then
               AddProc = False
               Exit For
            End If
         Next
         
         If AddProc Then 'Append only public elements
            TempProcLine = Trim(m_CodeModule.Lines(m_CodeModule.ProcBodyLine(TempProcName & "", TempProcKind), 1))
            If IncludePrivateProc Then
               AddProc = True
            ElseIf m_CodeModule.Parent.Type = vbext_ct_StdModule Then 'No Private
               AddProc = Not (Left(TempProcLine, 8) = "Private ")
            Else 'Public or Friend only
               AddProc = (Left(TempProcLine, 7) = "Public ") Or (Left(TempProcLine, 7) = "Friend ")
            End If
            
         End If
      
         If AddProc Then
            Set TempProc = New CodeModuleProcedure
            TempProc.Name = TempProcName
            TempProc.ProcKind = TempProcKind
            Set TempProc.ProcVbComponent = m_CodeModule.Parent
            On Error Resume Next
            m_ProcCollection.Add TempProc, TempProcName
            On Error GoTo 0
         End If
         
         CurrentLine = m_CodeModule.ProcStartLine(TempProcName & "", TempProcKind) + m_CodeModule.ProcCountLines(TempProcName, TempProcKind)
         ' Note: ProcCountLines is incl. comment above the procedure
         '       => ProcStartLine + ProcCountLines and not ProcBodyLine + ProcCountLines
         
      End If
      CurrentLine = CurrentLine + 1 'm_CodeModule.ProcCountLines(procName, procKind)
   Loop
   
   Set Procedures = m_ProcCollection

End Property

'---------------------------------------------------------------------------------------
' Property: HeaderItems
'---------------------------------------------------------------------------------------
'
' Entries of the declaration ready of the code module (variables, costants, type and enum definitions, ..)
'
' Parameters:
'     Refresh - True = rebuild collection; False = if a collection has already been determined, it is used and not checked again
'
' Returns:
'     Collection - CodeModuleHeaderItem references
'
' Remarks:
'     see also CodeModuleHeaderItemType
'
'---------------------------------------------------------------------------------------
Public Property Get HeaderItems(Optional ByVal Refresh As Boolean = False) As Collection

   Dim strCheckString As String
   Dim i As Long
   Dim strCheckLines As Variant
   Dim TempHeaderItem As CodeModuleHeaderItem
   Dim CheckHeaderItem As CodeModuleHeaderItem
   Dim lngPos As Long
   Dim bolAddHeaderItem As Boolean

   If Refresh Then
      Set m_HeaderItemCollection = Nothing
   ElseIf Not m_HeaderItemCollection Is Nothing Then
      Set HeaderItems = m_HeaderItemCollection
      Exit Property
   End If

   Set m_HeaderItemCollection = New Collection

   If m_CodeModule.CountOfDeclarationLines = 0 Then
      Set HeaderItems = m_HeaderItemCollection
      Exit Property
   End If
   
   strCheckString = m_CodeModule.Lines(1, m_CodeModule.CountOfDeclarationLines)
   strCheckLines = Split(strCheckString, vbNewLine)
   For i = LBound(strCheckLines) To UBound(strCheckLines)
      strCheckLines(i) = GetPlainCodeLine(strCheckLines(i))
   Next
   strCheckString = Join(strCheckLines, vbNewLine)
   
   'Remove line breaks with "_"
   strCheckString = Replace(strCheckString, "_" & vbNewLine, vbNewLine)
   Do While InStr(1, strCheckString, vbNewLine & vbNewLine) > 0
      strCheckString = Replace(strCheckString, vbNewLine & vbNewLine, vbNewLine)
   Loop

   strCheckLines = Split(strCheckString, vbNewLine)
   For i = LBound(strCheckLines) To UBound(strCheckLines)

      strCheckString = Trim$(strCheckLines(i))
      
      If strCheckString Like "public *" Or strCheckString Like "global *" Then
      
         bolAddHeaderItem = True
      
         lngPos = InStr(1, strCheckString, "Public ")
         If lngPos = 0 Then lngPos = InStr(1, strCheckString, "global ")
         strCheckString = Trim$(Mid$(strCheckString, lngPos + 7))
         Set TempHeaderItem = New CodeModuleHeaderItem
         
         '/**
         ' * @todo "Public x as long, y as long, ..." is not yet considered
         '**/
         
         If strCheckString Like "const *" Then
            TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Const
            strCheckString = Trim$(Mid$(strCheckString, 6))
         ElseIf strCheckString Like "enum *" Then
            TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Enum
            strCheckString = Trim$(Mid$(strCheckString, 5))
         ElseIf strCheckString Like "type *" Then
            TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Type
            strCheckString = Trim$(Mid$(strCheckString, 6))
         ElseIf strCheckString Like "declare *" Then
            TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Declare
            strCheckString = Trim$(Mid$(strCheckString, 8))
            If strCheckString Like "function *" Then
               strCheckString = Trim$(Mid$(strCheckString, 8))
            ElseIf strCheckString Like "sub *" Then
               strCheckString = Trim$(Mid$(strCheckString, 4))
            End If
         Else
            TempHeaderItem.ItemType = CodeModuleHeaderItemType.cmItemType_Var
         End If
         
         lngPos = InStr(1, strCheckString, " ")
         If lngPos = 0 Then lngPos = Len(strCheckString)
         TempHeaderItem.Name = Trim$(Left$(strCheckString, lngPos))
         
         Set TempHeaderItem.HeaderItemVBComponent = m_CodeModule.Parent
         
         For Each CheckHeaderItem In m_HeaderItemCollection
            If CheckHeaderItem.Name = TempHeaderItem.Name Then
               bolAddHeaderItem = False
               Exit For
            End If
         Next
         
         If bolAddHeaderItem Then
            m_HeaderItemCollection.Add TempHeaderItem, TempHeaderItem.Name
         End If
      
      End If
   Next
   
   Set HeaderItems = m_HeaderItemCollection

End Property


#If EARLYBINDING Then
Private Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Private Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim strCurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Check if the correct VbProject is selected (must be the one from CurrentDb):
      strCurrentDbName = UncPath(CurrentDb.Name)
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
