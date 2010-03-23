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

Private Const ERRNUMBER_CheckDependency_notrun As Long = vbObjectError + 5001
Private Const ERRDESCRIPTION_CheckDependency_notrun As String = "CheckDependency wurde noch nicht ausgeführt"


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

On Error GoTo HandleErr

   Set CodeModule = m_CodeModule

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CodeModule", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Property Set CodeModule(cm As VBIDE.CodeModule)
#Else
Public Property Set CodeModule(cm As Object)
#End If

On Error GoTo HandleErr

   Set m_CodeModule = cm

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CodeModule", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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

   Dim newReader As CodeModuleReader
   
On Error GoTo HandleErr

   Set newReader = New CodeModuleReader
   Set newReader.CodeModule = Me.CodeModule
   Set Copy = newReader

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "Copy", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

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

On Error GoTo HandleErr

   Name = m_CodeModule.Parent.Name

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Name", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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

On Error GoTo HandleErr

   If Not m_RequiredModulesChecked Then
      Err.Raise ERRNUMBER_CheckDependency_notrun, "RequiredModules", ERRDESCRIPTION_CheckDependency_notrun
   End If

   Set RequiredModules = m_RequiredModules

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "RequiredModules", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Function CheckDependency(Optional ByVal vbp As VBIDE.VBProject = Nothing, _
                                Optional ByRef oModulReaderCollection As Collection = Nothing) As Long
#Else
Public Function CheckDependency(Optional ByVal vbp As Object = Nothing, _
                                Optional ByRef oModulReaderCollection As Collection = Nothing) As Long
#End If

   Dim checkCollection As Collection
   Dim bolAddModule As Boolean

#If EARLYBINDING Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If

   Dim tempModuleReader As CodeModuleReader
   Dim newModuleReader As CodeModuleReader
   Dim tempProc As CodeModuleProcedure
   Dim tempHeaderItem As CodeModuleHeaderItem
   
On Error GoTo HandleErr

   Set m_RequiredModules = Nothing
   Set m_RequiredModules = New Collection
   
   If oModulReaderCollection Is Nothing Then
   
      Set oModulReaderCollection = New Collection
      
      If vbp Is Nothing Then
         Set vbp = CurrentVbProject
      End If

      For Each vbc In vbp.VBComponents
         Set tempModuleReader = New CodeModuleReader
         Set tempModuleReader.CodeModule = vbc.CodeModule
         oModulReaderCollection.Add tempModuleReader, tempModuleReader.Name
      Next

   End If
   
   For Each tempModuleReader In oModulReaderCollection
      If Not (tempModuleReader.Name = Me.Name) Then
         bolAddModule = False
      
         'Bei einem Standard-Modul müssen die Prozeduren und die Modul-Konstanten/Variablen geprüft werden
         If tempModuleReader.CodeModule.Parent.Type = vbext_ComponentType.vbext_ct_StdModule Then
            Set checkCollection = tempModuleReader.Procedures
            For Each tempProc In checkCollection
               If Me.ProcedureUsed(tempProc) Then
                  bolAddModule = True
                  Exit For
               End If
            Next
            
            If Not bolAddModule Then ' Variablen, Konstanten, ... prüfen
               Set checkCollection = tempModuleReader.HeaderItems
               For Each tempHeaderItem In checkCollection
                  If Me.HeaderItemUsed(tempHeaderItem) Then
                     bolAddModule = True
                     Exit For
                  End If
               Next
            End If
            
         Else 'bei Klassen reicht es aus, die Verwendung der Klasse zu prüfen
            If Me.ClassUsed(tempModuleReader.CodeModule.Parent.Name) Then
               bolAddModule = True
            End If
            
            If Not bolAddModule Then ' bei Klassen nur Enums prüfen
               Set checkCollection = tempModuleReader.HeaderItems
               For Each tempHeaderItem In checkCollection
                  If tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Enum Then
                     If Me.HeaderItemUsed(tempHeaderItem) Then
                        bolAddModule = True
                        Exit For
                     End If
                  End If
               Next
            End If
            
         End If
         
         If bolAddModule Then
            m_RequiredModules.Add tempModuleReader.Copy, tempModuleReader.Name
            'kann hier ein Zirkelbezug entstehen, wenn tempModuleReader-Referenz gespeichert wird?
            ' => zur Sicherheit Kopie übergeben
         End If
         
      End If
   Next
   
   m_RequiredModulesChecked = True
   CheckDependency = m_RequiredModules.Count

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "CheckDependency", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

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
   
   Dim strSearchString(5) As String
   Dim i As Long
   Dim bolUsed As Boolean
   Dim lngMaxCnt As Long
   
On Error GoTo HandleErr

   strSearchString(0) = " As " & Trim$(ClassName) & vbNewLine
   strSearchString(1) = " As " & Trim$(ClassName) & ","
   strSearchString(2) = " As " & Trim$(ClassName) & " "
   
   lngMaxCnt = 2
   For i = 0 To lngMaxCnt
      strSearchString(lngMaxCnt + 1 + i) = " New " & Mid(strSearchString(i), 5)
   Next
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   For i = LBound(strSearchString) To UBound(strSearchString)
      If m_CodeModuleText Like "*" & strSearchString(i) & "*" Then
         bolUsed = True
         Exit For
      End If
   Next
   
   ClassUsed = bolUsed

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ClassUsed", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Function ProcedureUsed(ByRef cmp As CodeModuleProcedure) As Boolean
   
   Dim searchStringArray() As String
   Dim i As Long
   Dim bolUsed As Boolean
   Dim proc As CodeModuleProcedure
   Dim bolOnlyWithModulePrefix As Boolean
   
On Error GoTo HandleErr

   'bei Überlappung von internen Prozedurnamen mit der zu prüfenden,
   'gilt nur bei der Kombination modulname.prozedur eine Übereinstimmung als Treffer
   For Each proc In Me.Procedures(True)
      If proc.Name = cmp.Name Then
         bolOnlyWithModulePrefix = True
         Exit For
      End If
   Next

   searchStringArray = cmp.SearchStrings(bolOnlyWithModulePrefix)
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   'Text durchsuchen
   For i = LBound(searchStringArray) To UBound(searchStringArray)
#If USEREGEX = 1 Then
      With Regex
         .Pattern = searchStringArray(i)
         .Global = False
         bolUsed = .Test(m_CodeModuleText)
         If bolUsed Then
            Exit For
         End If
      End With
#Else
      If m_CodeModuleText Like "*" & searchStringArray(i) & "*" Then
         bolUsed = True
         Exit For
      End If
#End If
   Next
   
   ProcedureUsed = bolUsed

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ProcedureUsed", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Function HeaderItemUsed(ByRef hdrItm As CodeModuleHeaderItem) As Boolean
   
   Dim searchStringArray() As String
   Dim i As Long
   Dim bolUsed As Boolean
   
   '/**
   ' * @bug Wenn sich eine Public-Variable mit einer interner (private Variable im Kopf oder in einer Proc)
   ' *      des zu durchsuchenden Codemoduls überlagert, wird die Public-Variable als Treffer angesehen
   '**/
   
On Error GoTo HandleErr

   searchStringArray = hdrItm.SearchStrings
   
   If StrPtr(m_CodeModuleText) = 0 Then
      m_CodeModuleText = PlainCodeText
   End If
   
   For i = LBound(searchStringArray) To UBound(searchStringArray)
#If USEREGEX = 1 Then
      With Regex
         .Pattern = searchStringArray(i)
         .Global = False
         bolUsed = .Test(m_CodeModuleText)
         If bolUsed Then Exit For
      End With
#Else
      If m_CodeModuleText Like "*" & searchStringArray(i) & "*" Then
         bolUsed = True
         Exit For
      End If
#End If
   Next

   HeaderItemUsed = bolUsed

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "HeaderItemUsed", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr

   For i = 1 To m_CodeModule.CountOfLines
      strCodeText = strCodeText & Trim$(getPlainCodeLine(m_CodeModule.Lines(i, 1))) & vbNewLine
   Next
   
   PlainCodeText = strCodeText

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "PlainCodeText", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Function getPlainCodeLine(ByVal strCodeLine As String) As String

   Dim posSingleQuote As Long
   Dim posDoubleQuote As Long
   
On Error GoTo HandleErr

   posSingleQuote = InStr(1, strCodeLine, "'")
   posDoubleQuote = InStr(1, strCodeLine, """")
   
   If posSingleQuote = 0 Then
      getPlainCodeLine = strCodeLine
   Else
      If posDoubleQuote = 0 Or posSingleQuote < posDoubleQuote Then
         getPlainCodeLine = Left$(strCodeLine, posSingleQuote - 1)
      Else ' posSingleQuote > posDoubleQuote and posDoubleQuote > 0 Then
         'schließendes Anführungszeichen:
         posDoubleQuote = InStr(posDoubleQuote + 1, strCodeLine, """")
         If posDoubleQuote > 0 Then
            getPlainCodeLine = Left$(strCodeLine, posDoubleQuote) & _
                            getPlainCodeLine(Mid$(strCodeLine, posDoubleQuote + 1))
         Else ' kann es eigentlich nicht geben
            getPlainCodeLine = strCodeLine
         End If
      End If
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "getPlainCodeLine", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Property Get Procedures(Optional ByVal bRefresh As Boolean = False, _
                               Optional ByVal includePrivateProc As Boolean = False) As Collection
   
On Error GoTo HandleErr

   Static bolIncludePrivateProc As Boolean
   
   If includePrivateProc <> bolIncludePrivateProc Then
      bRefresh = True
      bolIncludePrivateProc = includePrivateProc
   End If
   
   If bRefresh Then
      Set m_ProcCollection = Nothing
   ElseIf Not m_ProcCollection Is Nothing Then
      Set Procedures = m_ProcCollection
      Exit Property
   End If
   
   Set m_ProcCollection = New Collection

   Dim currentLine As Long
   Dim tempProc As CodeModuleProcedure
   Dim tempProcName As String
   Dim tempProcKind As vbext_ProcKind
   Dim bolAddProc As Boolean
   Dim tempProcLine As String

   currentLine = m_CodeModule.CountOfDeclarationLines + 1
   Do While currentLine <= m_CodeModule.CountOfLines
      tempProcKind = vbext_pk_Proc
      tempProcName = m_CodeModule.ProcOfLine(currentLine, tempProcKind)
      If Len(tempProcName) > 0 Then
         bolAddProc = True
         For Each tempProc In m_ProcCollection
            If tempProc.Name = tempProcName Then
               bolAddProc = False
               Exit For
            End If
         Next
         
         If bolAddProc Then 'nur Public-Elemente anfügen
            tempProcLine = Trim(m_CodeModule.Lines(m_CodeModule.ProcBodyLine(tempProcName & "", tempProcKind), 1))
            If includePrivateProc Then
               bolAddProc = True
            ElseIf m_CodeModule.Parent.Type = vbext_ct_StdModule Then 'kein Private
               bolAddProc = Not (Left(tempProcLine, 8) = "Private ")
            Else 'nur Public oder Friend
               bolAddProc = (Left(tempProcLine, 7) = "Public ") Or (Left(tempProcLine, 7) = "Friend ")
            End If
            
         End If
      
         If bolAddProc Then
            Set tempProc = New CodeModuleProcedure
            tempProc.Name = tempProcName
            tempProc.procKind = tempProcKind
            Set tempProc.ProcVBComponent = m_CodeModule.Parent
            On Error Resume Next
            m_ProcCollection.Add tempProc, tempProcName
            On Error GoTo 0
         End If
         
         currentLine = m_CodeModule.ProcStartLine(tempProcName & "", tempProcKind) + m_CodeModule.ProcCountLines(tempProcName, tempProcKind)
         ' Anm.: ProcCountLines ist inkl. Kommentar oberhalb der Prozedur
         '       => ProcStartLine + ProcCountLines und nicht ProcBodyLine + ProcCountLines
         
      End If
      currentLine = currentLine + 1 'm_CodeModule.ProcCountLines(procName, procKind)
   Loop
   
   Set Procedures = m_ProcCollection

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "Procedures", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
Public Property Get HeaderItems(Optional ByVal bRefresh As Boolean = False) As Collection

   Dim strCheckString As String
   Dim i As Long
   Dim strCheckLines As Variant
   Dim tempHeaderItem As CodeModuleHeaderItem
   Dim checkHeaderItem As CodeModuleHeaderItem
   Dim lngPos As Long
   Dim bolAddHeaderItem As Boolean

On Error GoTo HandleErr

   If bRefresh Then
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
      strCheckLines(i) = getPlainCodeLine(strCheckLines(i))
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
         Set tempHeaderItem = New CodeModuleHeaderItem
         
         '/**
         ' * @todo "Public x as long, y as long, ..." wird noch nicht berücksichtigt
         '**/
         
         If strCheckString Like "const *" Then
            tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Const
            strCheckString = Trim$(Mid$(strCheckString, 6))
         ElseIf strCheckString Like "enum *" Then
            tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Enum
            strCheckString = Trim$(Mid$(strCheckString, 5))
         ElseIf strCheckString Like "type *" Then
            tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Type
            strCheckString = Trim$(Mid$(strCheckString, 6))
         ElseIf strCheckString Like "declare *" Then
            tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Declare
            strCheckString = Trim$(Mid$(strCheckString, 8))
            If strCheckString Like "function *" Then
               strCheckString = Trim$(Mid$(strCheckString, 8))
            ElseIf strCheckString Like "sub *" Then
               strCheckString = Trim$(Mid$(strCheckString, 4))
            End If
         Else
            tempHeaderItem.ItemType = CodeModuleHeaderItemType.CodeModuleHeaderItemType_Var
         End If
         
         lngPos = InStr(1, strCheckString, " ")
         If lngPos = 0 Then lngPos = Len(strCheckString)
         tempHeaderItem.Name = Trim$(Left$(strCheckString, lngPos))
         
         Set tempHeaderItem.HeaderItemVBComponent = m_CodeModule.Parent
         
         For Each checkHeaderItem In m_HeaderItemCollection
            If checkHeaderItem.Name = tempHeaderItem.Name Then
               bolAddHeaderItem = False
               Exit For
            End If
         Next
         
         If bolAddHeaderItem Then
            m_HeaderItemCollection.Add tempHeaderItem, tempHeaderItem.Name
         End If
      
      End If
   Next
   
   Set HeaderItems = m_HeaderItemCollection

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "HeaderItems", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property


#If EARLYBINDING Then
Private Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Private Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim proj As VBProject
#Else
   Dim proj As Object
#End If
   Dim strCurrentDbName As String

On Error GoTo HandleErr

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = UNCPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> strCurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each proj In VBE.VBProjects
            If proj.FileName = strCurrentDbName Then
               Set m_CurrentVbProject = proj
               Exit For
            End If
         Next
      End If
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "CurrentVbProject", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property
