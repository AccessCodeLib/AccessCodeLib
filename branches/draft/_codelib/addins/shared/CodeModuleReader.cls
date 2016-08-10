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
' Class Module: CodeModuleReader
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
'  <file>_codelib/addins/shared/CodeModuleReader.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>_codelib/addins/shared/CodeModuleProcedure.cls</use>
'  <use>_codelib/addins/shared/CodeModuleHeaderItem.cls</use>
'  <use>file/modFiles.bas</use>
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
'/**
' <summary>
' Code-Modul, dessen Text ausgewertet werden soll
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
' Code-Modul, dessen Text ausgewertet werden soll
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
#If EARLYBINDING Then
Public Property Set CodeModule(ByRef NewRef As VBIDE.CodeModule)
#Else
Public Property Set CodeModule(ByRef NewRef As Object)
#End If

   Set m_CodeModule = NewRef

End Property

'---------------------------------------------------------------------------------------
' Function: Copy (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Instanz kopieren
' </summary>
' <returns>CodeModuleReader</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Copy() As CodeModuleReader

   Dim NewReader As CodeModuleReader

   Set NewReader = New CodeModuleReader
   Set NewReader.CodeModule = Me.CodeModule
   Set Copy = NewReader

End Function

'---------------------------------------------------------------------------------------
' Property: Name (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name von VBComponent des Code-Moduls
' </summary>
' <param name="Param"></param>
' <returns>String</returns>
' <remarks>
' Name = CodeModule.Parent.Name
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get Name() As String

   Name = m_CodeModule.Parent.Name

End Property

'---------------------------------------------------------------------------------------
' Property: RequiredModules (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Collection der erforderlichen Module
' </summary>
' <returns>Collection</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get RequiredModules() As Collection

   If Not m_RequiredModulesChecked Then
      Err.Raise ERRNUMBER_CHECKDEPENDENCY_NOTRUN, "RequiredModules", ERRDESCRIPTION_CHECKDEPENDENCY_NOTRUN
   End If

   Set RequiredModules = m_RequiredModules
 
End Property

'---------------------------------------------------------------------------------------
' Function: CheckDependency (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Abhängigkeiten prüfen
' </summary>
' <param name="vbp">VBProject, bei Nothing wird aktives VBProject verwendet</param>
' <param name="oModulReaderCollection">Collection der möglichen Abhängigkeiten, bei Nothing werden alle Code-Module aus VBProject verwendet,</param>
' <returns>Anzahl der Abhängigkeiten</returns>
' <remarks>
' </remarks>
'**/
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
      
         'Bei einem Standard-Modul müssen die Prozeduren und die Modul-Konstanten/Variablen geprüft werden
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
            
         Else 'bei Klassen reicht es aus, die Verwendung der Klasse zu prüfen
            If Me.ClassUsed(TempModuleReader.CodeModule.Parent.Name) Then
               AddModule = True
            End If
            
            If Not AddModule Then ' bei Klassen nur Enums prüfen
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
            'kann hier ein Zirkelbezug entstehen, wenn tempModuleReader-Referenz gespeichert wird?
            ' => zur Sicherheit Kopie übergeben
         End If
         
      End If
   Next
   
   m_RequiredModulesChecked = True
   CheckDependency = m_RequiredModules.Count

End Function


'---------------------------------------------------------------------------------------
' Function: ClassUsed (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft, ob Klasse im Code-Modul verwendet/benötigt wird
' </summary>
' <param name="ClassName">Name der Klasse</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
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
' Function: ProcedureUsed (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft, ob Prozedur im Code-Modul verwendet/benötigt wird
' </summary>
' <param name="cmp">CodeModuleProcedure-Referenz</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ProcedureUsed(ByRef CodeModuleProc As CodeModuleProcedure) As Boolean
   
   Dim SearchStringArray() As String
   Dim i As Long
   Dim IsUsed As Boolean
   Dim Proc As CodeModuleProcedure
   Dim OnlyWithModulePrefix As Boolean

   'bei Überlappung von internen Prozedurnamen mit der zu prüfenden,
   'gilt nur bei der Kombination modulname.prozedur eine Übereinstimmung als Treffer
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
' Function: HeaderItemUsed (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft, ob CodeModuleHeaderItem (Variable, Konstante, ...) im Code-Modul verwendet/benötigt wird
' </summary>
' <param name="hdrItm">CodeModuleHeaderItem-Referenz</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function HeaderItemUsed(ByRef HdrItm As CodeModuleHeaderItem) As Boolean
   
   Dim SearchStringArray() As String
   Dim i As Long
   Dim IsUsed As Boolean
   
   '/**
   ' * @bug Wenn sich eine Public-Variable mit einer interner (private Variable im Kopf oder in einer Proc)
   ' *      des zu durchsuchenden Codemoduls überlagert, wird die Public-Variable als Treffer angesehen
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
' Property: PlainCodeText (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt bereinigten Code des Code-Moduls zurück
' </summary>
' <returns>String</returns>
' <remarks>
' Entfernt Kommentarzeilen
' </remarks>
'**/
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
         'schließendes Anführungszeichen:
         PosDoubleQuote = InStr(PosDoubleQuote + 1, CodeLine, """")
         If PosDoubleQuote > 0 Then
            GetPlainCodeLine = Left$(CodeLine, PosDoubleQuote) & _
                            GetPlainCodeLine(Mid$(CodeLine, PosDoubleQuote + 1))
         Else ' kann es eigentlich nicht geben
            GetPlainCodeLine = CodeLine
         End If
      End If
   End If

End Function

'---------------------------------------------------------------------------------------
' Property: Procedures (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Collection aller Prozeduren (CodeModuleProcedure-Referenzen)
' </summary>
' <param name="bRefresh">True = Collection neu aufbauen; False = falls bereits eine Collection ermittelt wurde, wird diese verwendet und nicht noch einmal neu geprüft</param>
' <param name="includePrivateProc">True = inkl. als private deklarierten Prozeduren</param>
' <returns>Collection der CodeModuleProcedure-Referenzen</returns>
' <remarks>
' </remarks>
'**/
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
         
         If AddProc Then 'nur Public-Elemente anfügen
            TempProcLine = Trim(m_CodeModule.Lines(m_CodeModule.ProcBodyLine(TempProcName & "", TempProcKind), 1))
            If IncludePrivateProc Then
               AddProc = True
            ElseIf m_CodeModule.Parent.Type = vbext_ct_StdModule Then 'kein Private
               AddProc = Not (Left(TempProcLine, 8) = "Private ")
            Else 'nur Public oder Friend
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
         ' Anm.: ProcCountLines ist inkl. Kommentar oberhalb der Prozedur
         '       => ProcStartLine + ProcCountLines und nicht ProcBodyLine + ProcCountLines
         
      End If
      CurrentLine = CurrentLine + 1 'm_CodeModule.ProcCountLines(procName, procKind)
   Loop
   
   Set Procedures = m_ProcCollection

End Property

'---------------------------------------------------------------------------------------
' Property: HeaderItems (Josef Pötzl, 2010-03-23)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Einträge des Deklarationsbereits des Code-Moduls (Variablen, Kostanten, Type- und Enum-Definitionen, ..)
' </summary>
' <param name="bRefresh">True = Collection neu aufbauen; False = falls bereits eine Collection ermittelt wurde, wird diese verwendet und nicht noch einmal neu geprüft</param>
' <returns>Collection der CodeModuleHeaderItem-Referenzen</returns>
' <remarks>
' siehe auch CodeModuleHeaderItemType
' </remarks>
'**/
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
   
   'Zeilenumbrüche mit "_" entfernen
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
         ' * @todo "Public x as long, y as long, ..." wird noch nicht berücksichtigt
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
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
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
