VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibFileManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Import/Export der Access-Objekte in/aus lokaler Code-Bibliothek"
'---------------------------------------------------------------------------------------
' Klassenmodul: ACLibFileManager (Josef Pötzl)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Import/Export der Access-Objekte in/aus lokaler Code-Bibliothek
' </summary>
' <remarks>
' Das Modul "_config_Application" wird unabhängig vom Import-Modus nicht mehr überschrieben,
' sobald es in der Anwendung vorhanden ist.
' </remarks>
'\ingroup ACLibAddInImportWizard
'
' @todo Abhängigkeit zu Verweisen berücksichtigen und fehlende Verweise einstellen
' @todo getCodeLibInfoFromFile: Auslesen der CodeLib-Informationen aus Ac07-Formluaren/Berichten,
'       die mit SaveAsText erstellt wurden (unicode-Text)
'
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/ImportWizard/ACLibFileManager.cls</file>
'  <description>Kuemmert sich um den Import und Export der Codemodule aus dem und in das Repository</description>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>file/modFiles.bas</use>
'  <use>text/StringTools.bas</use>
'  <use>utils/AccessProgressBar.cls</use>
'  <use>_codelib/addins/ImportWizard/defGlobal_ACLibImportWizard.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
' !!! in m_ImportFileCollection wird ein Array zum Abspeichern verwendet
'     (0): File-Object
'     (1): Import-Modus
'     Grund: Type konnte nicht verwendet werden und eine extra Klasse ist (noch) zu umständlich
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

Private Const m_EXTENSION_KEY As String = "ACLibFileManager"

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_ACLibConfig As String = "ACLibConfiguration"
Private Const EXTENSION_PROPNAME_LocalRepositoryRoot As String = "LocalRepositoryRoot"
Private Const EXTENSION_PROPNAME_PrivateRepositoryRoot As String = "PrivateRepositoryRoot"

#Const DEBUGMODE = 0
#Const EARLYBINDING = 0

#If EARLYBINDING = 1 Then
Private m_CurrentVbProject As VBIDE.VBProject
#Else
Private m_CurrentVbProject As Object
#End If

Private Const SEARCHSTRING_CODELIB_BEGIN As String = "<codelib>"
Private Const SEARCHSTRING_CODELIB_END   As String = "</codelib>"
Private Const SEARCHSTRING_FILE_BEGIN As String = "<file>"
Private Const SEARCHSTRING_FILE_END   As String = "</file>"
Private Const SEARCHSTRING_LICENSE_BEGIN As String = "<license>"
Private Const SEARCHSTRING_LICENSE_END   As String = "</license>"
Private Const SEARCHSTRING_DESCRIPTION_BEGIN As String = "<description>"
Private Const SEARCHSTRING_DESCRIPTION_END   As String = "</description>"
Private Const SEARCHSTRING_USE_BEGIN As String = "<use>"
Private Const SEARCHSTRING_USE_END   As String = "</use>"
Private Const SEARCHSTRING_REF_BEGIN As String = "<ref>"
Private Const SEARCHSTRING_REF_END   As String = "</ref>"
Private Const SEARCHSTRING_REF_NAME_BEGIN As String = "<name>"
Private Const SEARCHSTRING_REF_NAME_END   As String = "</name>"
Private Const SEARCHSTRING_REF_MAJOR_BEGIN As String = "<major>"
Private Const SEARCHSTRING_REF_MAJOR_END   As String = "</major>"
Private Const SEARCHSTRING_REF_MINOR_BEGIN As String = "<minor>"
Private Const SEARCHSTRING_REF_MINOR_END   As String = "</minor>"
Private Const SEARCHSTRING_REF_GUID_BEGIN As String = "<guid>"
Private Const SEARCHSTRING_REF_GUID_END   As String = "</guid>"
Private Const SEARCHSTRING_TEST_BEGIN As String = "<test>"
Private Const SEARCHSTRING_TEST_END   As String = "</test>"

Private Const SEARCHSTRING_EXECUTE_BEGIN As String = "<execute>"
Private Const SEARCHSTRING_EXECUTE_END As String = "</execute>"

Private Const SEARCHSTRING_REPLACE_BEGIN As String = "<replace>"
Private Const SEARCHSTRING_REPLACE_END   As String = "</replace>"
Private Const SEARCHSTRING_ATTRIBUTNAME_BEGIN As String = "Attribute VB_Name = """
Private Const SEARCHSTRING_ATTRIBUTNAME_END   As String = """"
Private Const SEARCHSTRING_FORMIDENTIFER As String = "BEGIN FORM"
Private Const SEARCHSTRING_REPORTIDENTIFER As String = "Begin Report"

Private Const MODULNAME_CONFIG_APPLICATION As String = "_config_Application"

Private Const REPOSTITORY_ROOT_CODE_ApplicationRoot As String = "%AppFolder%"
Private Const REPOSTITORY_ROOT_CODE_PrivateRoot As String = "%PrivateRoot%"

Private m_ImportFileCollection As Collection
Private m_ReplacedFilesCollection As Collection
Private m_cli As CodeLibInfo


#If EARLYBINDING = 1 Then
Private m_fso As FileSystemObject
#Else
Private m_fso As Object
'Private Const ForReading As Long = 1
#End If

'
Private m_LocalRepositoryRootDirectory As String   ' Lokales Root-Verzeichnis für Repository-Export in die CodeLib
Private m_PrivateRepositoryRootDirectory As String ' Lokales Root-Verzeichnis für privaten Repository-Export

' Events
Public Event PropertyMissingLocalRepositoryRootDirectory(ByRef NewValue As String)
Public Event ImportRepositoryFile(ByVal ObjectName As String, ByVal RepositoryFile As String, _
                                  ByVal ElementType As CodeLibElementType, ByRef Dependency As Variant, _
                                  ByRef ImportFile As Object, ByRef Cancel As Integer)


'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As Object) ' as ApplicationHandler
   Set m_ApplicationHandler = objRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

End Sub


' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                              ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub



'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ACLibFileManager
'---------------------------------------------------------------------------------------

Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         
   End Select
   
End Sub

Private Sub Dispose()
On Error Resume Next
   Set m_ImportFileCollection = Nothing
   Set m_fso = Nothing
   Set m_CurrentVbProject = Nothing
End Sub

Public Sub ExportAll()
   ExportAllModules
   ExportAllForms
   ExportAllReports
End Sub

Public Sub ExportAllModules()

   Dim ao As AccessObject

   activateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllModules
      exportVbComponent ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Module-Export abgeschlossen"
#End If

End Sub

Public Sub ExportAllForms()

   Dim ao As AccessObject

   If checkDisallowExportFormReport() Then
      Exit Sub
   End If

   activateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllForms
      exportAccessObject AcObjectType.acForm, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Formular-Export abgeschlossen"
#End If

End Sub

Private Function checkDisallowExportFormReport() As Boolean

   If Val(CurrentDb.Version) >= 12 Then
      MsgBox "Formular und Berichte ab dem Access2007-Format dürfen nicht exportiert werden." & _
             "Siehe Bug #16"
      checkDisallowExportFormReport = True
   End If

End Function

Public Sub ExportAllReports()

   Dim ao As AccessObject

   activateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllReports
      exportAccessObject AcObjectType.acReport, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Bericht-Export abgeschlossen"
#End If

End Sub

Public Sub ImportAllFilesFromRepository(Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                                        Optional ByVal bImportTestFiles As Boolean = False)

   Dim folderCol As Collection
#If EARLYBINDING Then
   Dim checkFolder As Folder
   Dim tempFolder As Folder
   Dim TempFile As File
#Else
   Dim checkFolder As Object
   Dim tempFolder As Object
   Dim TempFile As Object
#End If
   
   Set folderCol = New Collection
   Set m_ImportFileCollection = New Collection
   
   Set tempFolder = fso.GetFolder(LocalRepositoryRootDirectory)
   folderCol.Add tempFolder, tempFolder.Path
   
   Do While folderCol.Count > 0
      
      Set checkFolder = folderCol(1)
      
      For Each TempFile In checkFolder.Files
         If Not ignoreFile(TempFile) Then
            addMissingFile TempFile, ImportMode
         End If
      Next
      
      For Each tempFolder In checkFolder.SubFolders
         If Not ignoreFolder(tempFolder) Then
            folderCol.Add tempFolder, tempFolder.Path
         End If
      Next
   
      folderCol.Remove 1
   
   Loop
   Set TempFile = Nothing
   Set folderCol = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

End Sub

Public Sub ImportRepositoryFile(ByVal RepositoryPath As String, _
                       Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                       Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim pathString As String

   pathString = GetRepositoryFullPath(RepositoryPath)
   
   Dim TempFile As Object
   Set TempFile = fso.GetFile(pathString)
   addMissingFile TempFile, ImportMode
   Set TempFile = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

End Sub

Public Sub ImportRepositoryFiles(ByRef RepositoryPathArray() As String, _
                        Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                        Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim ArraySize As Long
   Dim i As Long
   Dim pathString As String
   Dim TempFile As Object

   ArraySize = UBound(RepositoryPathArray)
   For i = LBound(RepositoryPathArray) To ArraySize - 1
      pathString = GetRepositoryFullPath(RepositoryPathArray(i))
      Set TempFile = fso.GetFile(pathString)
      addMissingFile TempFile, ImportMode
   Next
   Set TempFile = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

End Sub

Private Sub importFilesFromImportCollection( _
                                   Optional ByRef CompileAfterImport As Boolean = True, _
                                   Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim TempFile As Object
   Dim fileImportMode As CodeLibImportMode
   Dim i As Long
   Dim maxCnt As Long
   
   Dim colItem As Variant
   
   ' eine VB-Modul aus CurrentVbProject aktivieren, sonst wirkt RunCommand acCmdCompileAndSaveAllModules
   ' möglicherweise im falschen VBProject
   ' zur Sicherheit mit On error resume next, weil es nicht unbedingt notwendig ist und den Ablauf nicht stört
   On Error Resume Next
   activateCurrentProject
   On Error GoTo 0
   If CurrentProject.AllModules.Count > 0 Then
      DoCmd.RunCommand acCmdCompileAndSaveAllModules
   End If
    
    maxCnt = m_ImportFileCollection.Count
    
    '/*
    ' * TODO: Bestimmung der Anzahl alle Dateien (inkl. Abhängigkeiten) vor dem Import
    ' * TODO: Umbau AccessProgressBar auf einmaligen Durchlauf (Init nur einmal)
    ' */
   
    i = 1
    Do While i <= maxCnt
        colItem = m_ImportFileCollection(i)
        Set TempFile = colItem(0)
        fileImportMode = colItem(1)
        
        AccessProgressBar.Init "Importiere " & TempFile & "...", 2, 1
        AccessProgressBar.PerformStep
        
        ImportFile TempFile, fileImportMode, bImportTestFiles
        
        AccessProgressBar.PerformStep
        maxCnt = m_ImportFileCollection.Count
        i = i + 1
    Loop



   If CompileAfterImport Then
      activateCurrentProject
      CurrentProject.Application.DoCmd.RunCommand acCmdCompileAndSaveAllModules
   End If
   
   'Executes ausführen
   If (0 / 1) + (Not Not m_cli.ExecuteList) Then
      AccessProgressBar.Init "Executes ausführen...", UBound(m_cli.ExecuteList) + 1, 1
      For i = 0 To UBound(m_cli.ExecuteList)
        AccessProgressBar.PerformStep
        If StringTools.Contains(m_cli.ExecuteList(i), REPOSTITORY_ROOT_CODE_PrivateRoot) Then
            Eval VBA.Strings.Replace(m_cli.ExecuteList(i), REPOSTITORY_ROOT_CODE_PrivateRoot, LocalRepositoryRootDirectory())
        Else
            Eval (m_cli.ExecuteList(i))
        End If
      Next
      If AccessProgressBar.IsInitialized Then AccessProgressBar.Clear
   End If
   
   'Modul-Variablen aufräumen, da Import-Vorgang an dieser Stelle abgeschlossen ist
   Set m_ImportFileCollection = Nothing
   Set m_ReplacedFilesCollection = Nothing
   Set m_fso = Nothing
   
#If DEBUGMODE Then
   Debug.Print "Import abgeschlossen"
#End If

End Sub


Private Function ignoreFolder(ByRef oTestFolder As Object) As Boolean
'/*
' * @todo Ausschlussliste für Verzeichnisse flexibler gestalten
'**/

   Select Case oTestFolder.Name
      Case ".svn", "_codelib"
         ignoreFolder = True
      Case Else
         '
   End Select

End Function

Private Function ignoreFile(ByRef oTestFile As Object) As Boolean
'/*
' * @todo Ausschlussliste für Dateien flexibler gestalten
'**/

   Select Case oTestFile.Name
      Case "_info.txt", "_config_Application.bas"
         ignoreFile = True
      Case Else
         '
   End Select

End Function

#If EARLYBINDING Then
Public Property Get CurrentVbProject() As VBIDE.VBProject
#Else
Public Property Get CurrentVbProject() As Object
#End If

#If EARLYBINDING Then
   Dim proj As VBProject
#Else
   Dim proj As Object
#End If
   Dim strCurrentDbName As String

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

End Property

Private Sub activateCurrentProject()

   Dim bolActivate As Boolean
#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If

   If CurrentVbProject.VBComponents.Count > 0 Then
      'zuerst nur Code-Module verwenden, da sonst Formulare/Berichte im Entwurf geöffnet werden
      For Each vbc In CurrentVbProject.VBComponents
         If vbc.Type <> vbext_ct_Document Then
            vbc.Activate
            bolActivate = True
            Exit For
         End If
      Next
      'falls nur vbext_ct_Document vorhanden sind, das erste verwenden
      If Not bolActivate Then
         CurrentVbProject.VBComponents(1).Activate
      End If
   End If

End Sub

#If EARLYBINDING Then
Private Property Get fso() As FileSystemObject
#Else
Private Property Get fso() As Object
#End If

   If fso Is Nothing Then
      'Set fso = New FileSystemObject
      Set m_fso = CreateObject("Scripting.FileSystemObject")
   End If
   Set fso = m_fso

End Property

Private Property Get CurrentFileCollection() As Collection

   If m_ImportFileCollection Is Nothing Then
      Set m_ImportFileCollection = New Collection
   End If
   Set CurrentFileCollection = m_ImportFileCollection

End Property

Private Property Get CurrentReplacedFilesCollection() As Collection

   If m_ReplacedFilesCollection Is Nothing Then
      Set m_ReplacedFilesCollection = New Collection
   End If
   Set CurrentReplacedFilesCollection = m_ReplacedFilesCollection

End Property

Private Sub addReplacedFilePath(ByVal repFilePath As String)

   Dim TempFile As Object
   Dim TempFilePath As Variant
   Dim i As Long
   Dim maxCount As Long
   Dim col As Collection

   maxCount = CurrentFileCollection.Count
   For i = 1 To maxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set TempFile = m_ImportFileCollection(i)(0)
      If TempFile.Path = repFilePath Then
         m_ImportFileCollection.Remove TempFile.Path
         Exit For
      End If
   Next
   Set TempFile = Nothing
   
   Set col = CurrentReplacedFilesCollection
   For Each TempFilePath In col
      If TempFilePath = repFilePath Then
         Exit Sub
      End If
   Next
   
   CurrentReplacedFilesCollection.Add CVar(repFilePath), repFilePath

End Sub

Private Function IsReplacedFile(ByRef RepFile As Object) As Boolean

   Dim TempFile As Variant

   For Each TempFile In CurrentReplacedFilesCollection
      If TempFile = RepFile.Path Then
         IsReplacedFile = True
         Exit Function
      End If
   Next

End Function

Private Property Get LocalRepositoryRootDirectory() As String

   If Len(m_LocalRepositoryRootDirectory) = 0 Then
      RaiseEvent PropertyMissingLocalRepositoryRootDirectory(m_LocalRepositoryRootDirectory)
      If Len(m_LocalRepositoryRootDirectory) = 0 Then
         m_LocalRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLibConfig, EXTENSION_PROPNAME_LocalRepositoryRoot, vbNullString)
      End If
   End If
   LocalRepositoryRootDirectory = m_LocalRepositoryRootDirectory
 
End Property

Private Property Let LocalRepositoryRootDirectory(ByVal LocalPath As String)

   If Right$(LocalPath, 1) <> "\" Then
      LocalPath = LocalPath & "\"
   End If
   m_LocalRepositoryRootDirectory = LocalPath
 
End Property

Private Property Get PrivateRepositoryRootDirectory() As String

   If Len(m_PrivateRepositoryRootDirectory) = 0 Then
      RaiseEvent PropertyMissingLocalRepositoryRootDirectory(m_PrivateRepositoryRootDirectory)
      If Len(m_PrivateRepositoryRootDirectory) = 0 Then
         m_PrivateRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLibConfig, EXTENSION_PROPNAME_PrivateRepositoryRoot, vbNullString)
      End If
   End If
   PrivateRepositoryRootDirectory = m_PrivateRepositoryRootDirectory

End Property

Private Property Let PrivateRepositoryRootDirectory(ByVal PrivatePath As String)

   If Right$(PrivatePath, 1) <> "\" Then
      PrivatePath = PrivatePath & "\"
   End If
   m_PrivateRepositoryRootDirectory = PrivatePath

End Property

Public Function GetRepositoryFullPath(ByVal ReleativPath As String) As String

   Dim repPath As String

   ReleativPath = Replace(ReleativPath, "/", "\")
   
   If Left(ReleativPath, Len(REPOSTITORY_ROOT_CODE_ApplicationRoot)) = REPOSTITORY_ROOT_CODE_ApplicationRoot Then
      repPath = CurrentProject.Path & "\"
      ReleativPath = Mid$(ReleativPath, Len(REPOSTITORY_ROOT_CODE_ApplicationRoot) + 1)
   
   ElseIf Left(ReleativPath, Len(REPOSTITORY_ROOT_CODE_PrivateRoot)) = REPOSTITORY_ROOT_CODE_PrivateRoot Then
      repPath = PrivateRepositoryRootDirectory
      If Len(repPath) = 0 Then
         Err.Raise vbObjectError, "getRepositoryFullPath", "Wert für privates Root-Verzeichnise fehlt (PrivateRepositoryRootDirectory = '')."
         Exit Function
      End If
      ReleativPath = Mid$(ReleativPath, Len(REPOSTITORY_ROOT_CODE_PrivateRoot) + 1)
      
   Else
      repPath = LocalRepositoryRootDirectory
      If Len(repPath) = 0 Then
         Err.Raise vbObjectError, "getRepositoryFullPath", "Wert für lokales Root-Verzeichnise fehlt (LocalRepositoryRootDirectory = '')."
         Exit Function
      End If
   End If
   
   Do While Left$(ReleativPath, 1) = "\"
      ReleativPath = Mid$(ReleativPath, 2)
   Loop

   GetRepositoryFullPath = repPath & ReleativPath

End Function

Private Sub ImportFile(ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode, _
              Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim i As Long
   Dim TempFile As Object
   Dim FilePath As String
   Dim cancelImport As Integer
   
   Dim ReplacedFile As String

   getCodeLibInfoFromFile m_cli, ImportFile, (ImportMode <> clim_ImportSelectedOnly)
   
   If Len(m_cli.RepositoryFileReplacement) > 0 Then
      addReplacedFilePath GetRepositoryFullPath(m_cli.RepositoryFileReplacement)
   End If
   
   'Abbruch ermöglichen
   RaiseEvent ImportRepositoryFile(m_cli.Name, m_cli.RepositoryFile, m_cli.Type, m_cli.Dependency, ImportFile, cancelImport)
   If cancelImport Then
      Exit Sub
   End If
   
   'Lizenz (LICENSE-Block)
   If Len(m_cli.LicenseFile) > 0 Then
      FilePath = GetRepositoryFullPath(m_cli.LicenseFile)
      Set TempFile = fso.GetFile(FilePath)
      addMissingFile TempFile, clim_ImportMissingItems
   End If

   'Erforderliche Module (USE-Block)
   If ImportMode <> clim_ImportSelectedOnly Then
      If (0 / 1) + (Not Not m_cli.Dependency) Then
      For i = 0 To UBound(m_cli.Dependency)
         FilePath = GetRepositoryFullPath(m_cli.Dependency(i))
         If fso.FileExists(FilePath) Then
            Set TempFile = fso.GetFile(FilePath)
            addMissingFile TempFile, ImportMode
         End If
      Next
      End If
   End If
   
   'zuerst Verweise
   If (0 / 1) + (Not Not m_cli.References) Then
      For i = 0 To UBound(m_cli.References)
         addMissingReference m_cli.References(i)
      Next
   End If
   
   'dann Codemodul
   Select Case m_cli.Type
      Case CodeLibElementType.clet_ClassModule, CodeLibElementType.clet_StdModule
         importVbComponent m_cli, ImportFile, ImportMode
      Case CodeLibElementType.clet_Form
         importAccessObject acForm, m_cli, ImportFile, ImportMode
      Case CodeLibElementType.clet_Report
         importAccessObject acReport, m_cli, ImportFile, ImportMode
      Case Else
         ' eventuell Fehler auslösen?
   End Select
      
   'eventuell die Tests
   If bImportTestFiles Then
   If (0 / 1) + (Not Not m_cli.TestFiles) Then
      For i = 0 To UBound(m_cli.TestFiles)
         FilePath = GetRepositoryFullPath(m_cli.TestFiles(i))
         If fso.FileExists(FilePath) Then
            Set TempFile = fso.GetFile(FilePath)
            addMissingFile TempFile, ImportMode
         End If
      Next
   End If
   End If
   
#If DEBUGMODE Then
   Debug.Print ImportFile.Path; " --> "; m_cli.Name
#End If
  
End Sub

Private Sub addMissingReference(ByRef NewCodeLibRef As CodeLibInfoReference)
   
   Dim refs As Access.References
   Dim ref As Access.Reference
   Dim bolReplace As Boolean

   Set refs = Access.References
   For Each ref In refs
      If ref.Name = NewCodeLibRef.Name Then
         If ref.Major < NewCodeLibRef.Major Then
            bolReplace = True
         ElseIf ref.Major = NewCodeLibRef.Major And ref.Minor < NewCodeLibRef.Minor Then
            bolReplace = True
         Else 'Nicht ersetzen, Abbruch
            Exit Sub
         End If
         Exit For
      End If
   Next
   
   If bolReplace Then
      refs.Remove ref
   End If
   
   refs.AddFromGuid NewCodeLibRef.GUID, NewCodeLibRef.Major, NewCodeLibRef.Minor
 
End Sub

Private Sub addMissingFile(ByVal UsedFile As Object, ByVal ImportMode As CodeLibImportMode)
   
   Dim varfilColItem As Variant
   Dim TempFile As Object
   Dim i As Long
   Dim maxCount As Long

   If IsReplacedFile(UsedFile) Then Exit Sub

   maxCount = CurrentFileCollection.Count
   For i = 1 To maxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set TempFile = m_ImportFileCollection(i)(0)
      If TempFile.Path = UsedFile.Path Then
         Exit Sub
      End If
   Next

   varfilColItem = Array(UsedFile, ImportMode)
   m_ImportFileCollection.Add varfilColItem, UsedFile.Path
 
End Sub


Private Sub exportAccessObject(ByRef aoType As AcObjectType, ByRef aoName As String)

   Select Case aoType
      Case AcObjectType.acModule
         exportVbComponent aoName
      Case AcObjectType.acForm, AcObjectType.acReport
         exportAccessObjectCli aoType, aoName
      Case Else
         '
   End Select

End Sub

Private Sub exportAccessObjectCli(ByRef aoType As AcObjectType, ByRef aoName As String)

'/**
' * @todo Was ist mit Formularen und Berichten ohne Modul?
'**/
   
   Dim cli As CodeLibInfo

   getCodeLibInfoFromAccessObject cli, aoType, aoName
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      createMissingFolder getParentFolderPath(cli.LocalFile)
      Application.SaveAsText aoType, aoName, cli.LocalFile
      cleanAccessBugExport cli.LocalFile
   End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert

#If DEBUGMODE Then
   Debug.Print aoName; " --> ";
   If Len(cli.RepositoryFile) > 0 Then
      Debug.Print cli.LocalFile
   Else
      Debug.Print "nicht exportiert"
   End If
#End If
 
End Sub

Private Sub cleanAccessBugExport(ByVal sFile As String)
   
   Dim isAcc2010 As Boolean
   
   On Error Resume Next
   If Application.Version >= "12.0" Then
      isAcc2010 = True
   End If
   On Error GoTo 0
   If Not isAcc2010 Then
      Exit Sub
   End If
   
   Dim Pos As Long
   Dim checkString As String
   Dim TempString As String
   Dim i As Long
   Dim fileNumber As Long

   fileNumber = FreeFile
   Open sFile For Binary Access Read As fileNumber
   checkString = String$(LOF(fileNumber), 0)
   Get fileNumber, , checkString
   Close fileNumber
   
   TempString = Replace(checkString, "NoSaveCTIWhenDisabled =1", vbNullString)
   If StrComp(TempString, checkString, vbBinaryCompare) = 0 Then
      Exit Sub
   End If
   
   'tempString = Replace(checkString, vbnewline & vbn, vbNullString)
   
   Kill sFile
   fileNumber = FreeFile
   Open sFile For Binary Access Write As fileNumber
      Put fileNumber, , TempString
   Close fileNumber
   
   
End Sub

Private Sub exportVbComponent(ByRef vbcName As String)

#If EARLYBINDING Then
   Dim vbc As VBIDE.VBComponent
   Dim cm As VBIDE.CodeModule
#Else
   Dim vbc As Object
   Dim cm As Object
#End If

   Dim cli As CodeLibInfo
  
   Set vbc = CurrentVbProject.VBComponents(vbcName)
   
   'Leerzeilen am Ende säubern
   Set cm = vbc.CodeModule
   Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
      cm.DeleteLines cm.CountOfLines
   Loop
   
   getCodeLibInfoFromVbComponent cli, vbc
   
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      createMissingFolder getParentFolderPath(cli.LocalFile)
      vbc.Export cli.LocalFile
   End If
   
#If DEBUGMODE Then
   Debug.Print vbcName; " --> "; cli.LocalFile
#End If

End Sub

Private Sub importVbComponent(ByRef cli As CodeLibInfo, ByRef ImportFile As Object, ByRef fileImportMode As CodeLibImportMode)
   
#If EARLYBINDING = 1 Then
   Dim vbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
   Dim cm As VBIDE.CodeModule
#Else
   Dim vbcCol As Object
   Dim vbc As Object
   Dim cm As Object
#End If
   Dim codeModuleExists As Boolean

   If Len(Trim$(cli.Name)) = 0 Then
      Err.Raise vbObjectError, "ADF_ImportExport.importVbComponent", "Es wurde versucht ein Module ohne Namen zu erstellen."
      Exit Sub
   End If
   
   Set vbcCol = CurrentVbProject.VBComponents
   For Each vbc In vbcCol
      If vbc.Name = cli.Name Then
         If vbc.Type <> cli.Type Then
            ' trotz clim_ImportMissingItems importieren
         ElseIf fileImportMode = clim_ImportMissingItems Then 'Abbruch, damit nicht überschrieben wird
            Exit Sub
         ElseIf vbc.Name = MODULNAME_CONFIG_APPLICATION Then ' _config_Application niemals überschreiben, wenn es schon existiert
            Exit Sub
         End If
         codeModuleExists = True
         Exit For
      End If
   Next

   If codeModuleExists Then   ' überprüfen, ob Typen übereinstimmen
      If vbc.Type <> cli.Type Then
         vbcCol.Remove vbc
         Set vbc = Nothing
         codeModuleExists = False
      End If
   End If

   If codeModuleExists Then ' Inhalt über CodeModule ändern, damit MS-Add-In für Quellcodeverwaltung keine Schwierigkeiten macht
      Set cm = vbc.CodeModule
      cm.DeleteLines 1, cm.CountOfLines
      cm.AddFromFile ImportFile.Path
      If cli.Type = CodeLibElementType.clet_ClassModule Then
         'wegschneiden von:
         'VERSION 1.0 CLASS
         'BEGIN
         '  MultiUse = -1  'True
         'End
         deleteClassHeaderLinesFromCodeModul cm
      End If
      'Leerzeilen am Ende säubern
      Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
         cm.DeleteLines cm.CountOfLines
      Loop
      
   Else
      vbcCol.import ImportFile.Path
   End If

End Sub

#If EARLYBINDING = 1 Then
Private Sub deleteClassHeaderLinesFromCodeModul(ByRef cm As CodeModule)
#Else
Private Sub deleteClassHeaderLinesFromCodeModul(ByRef cm As Object)
#End If
'
'Header sieht so aus:
'--
'VERSION 1.0 CLASS
'BEGIN
'  MultiUse = -1  'True
'End
'--

   Dim target As String
   Dim searchStartLine As Long, searchStartCol As Long, _
       searchEndLine As Long, searchEndCol As Long
   Dim headerStartLine As Long, headerEndLine As Long

   searchStartLine = 1
   searchStartCol = 1
   searchEndLine = -1
   searchEndCol = -1
   
   target = "VERSION*CLASS"
   If Not cm.Find(target, searchStartLine, searchStartCol, searchEndLine, searchEndCol, True, False, True) Then
      Exit Sub
   End If
   headerStartLine = searchStartLine
   
   target = "BEGIN"
   searchStartLine = searchEndLine
   searchStartCol = searchEndCol
   searchEndLine = -1
   searchEndCol = -1
   If Not cm.Find(target, searchStartLine, searchStartCol, searchEndLine, searchEndCol, True, False, False) Then
      Exit Sub
   End If
   
   target = "End"
   searchStartLine = searchEndLine
   searchStartCol = searchEndCol
   searchEndLine = -1
   searchEndCol = -1
   If Not cm.Find(target, searchStartLine, searchStartCol, searchEndLine, searchEndCol, True, False, False) Then
      Exit Sub
   End If
   headerEndLine = searchEndLine

#If DEBUGMODE = 1 Then
   Debug.Print "Header lines:", headerStartLine, headerEndLine
#End If

   cm.DeleteLines headerStartLine, headerEndLine

End Sub

Private Sub importAccessObject(ByRef acojbType As AcObjectType, ByRef cli As CodeLibInfo, _
                               ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode)
   
   Dim col As Object ' AllForms oder AllReports
   Dim acObj As AccessObject

   If Len(Trim$(cli.Name)) = 0 Then
      Err.Raise vbObjectError, "ADF_ImportExport.importVbComponent", "Es wurde versucht ein Module ohne Namen zu erstellen."
      Exit Sub
   End If
   
   Select Case acojbType
      Case AcObjectType.acForm
         Set col = CurrentProject.AllForms
      Case AcObjectType.acReport
         Set col = CurrentProject.AllReports
      Case Else
'/**
' * @todo Import von Makros ermöglichen
'**/
         Exit Sub
   
   End Select
   
   'Sonderfall: ACLibImportWizardForm .. weil dieser Formularname vom Import-Wizard verwendet wird
                                       ' und Access der Meinung ist, das Formular ist bereits geöffnet
   If cli.Name = "ACLibImportWizardForm" Then
      'immer umbennenen, da sonst bei ungeöffneten Formular das Formular des Add-Ins gelöscht wird!
      'If CodeProject.AllForms(cli.Name).IsLoaded Then
         cli.Name = cli.Name & "_NEU"
      'End If
   End If

   For Each acObj In col
      If acObj.Name = cli.Name Then
         If ImportMode = clim_ImportMissingItems Then 'Abbruch, damit nicht überschrieben wird
            Exit Sub
         End If
         If acObj.IsLoaded Then
            DoCmd.Close acojbType, cli.Name, acSaveNo
         End If
         Set acObj = Nothing
         DoCmd.DeleteObject acojbType, cli.Name
         Exit For
      End If
   Next
   Set col = Nothing
   
   Application.LoadFromText acojbType, cli.Name, ImportFile.Path
   DoCmd.Close acojbType, cli.Name

   'Sonderfall: ACLibImportWizardForm
   If cli.Name = "ACLibImportWizardForm_NEU" Then
      MsgBox "Bitte das Formular ""ACLibImportWizardForm_NEU"" manuell in ""ACLibImportWizardForm"" umbennenen." & vbNewLine & _
             "Es konnte nicht mit dem Originalnamen erstellt werden, da der Import-Assistent ein gleichnamiges Formular geöffnet hat."
   End If

End Sub

#If EARLYBINDING = 1 Then
Public Function GetCodeLibInfoFromCodeModule(ByVal sCodeModulName As String, Optional ByVal oVBProject As VBProject) As CodeLibInfo
#Else
Public Function GetCodeLibInfoFromCodeModule(ByVal sCodeModulName As String, Optional ByVal oVBProject As Object) As CodeLibInfo
#End If

   Dim cli As CodeLibInfo
#If EARLYBINDING = 1 Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If
   
   If oVBProject Is Nothing Then
      Set oVBProject = CurrentVbProject
   End If
   
   Set vbc = oVBProject.VBComponents(sCodeModulName)
   
   getCodeLibInfoFromVbComponent cli, vbc
   
   GetCodeLibInfoFromCodeModule = cli
   
End Function

Private Sub getCodeLibInfoFromVbComponent(ByRef cli As CodeLibInfo, ByRef vbc As Object)
        
   Dim checkString As String
   Dim countOfCodeLines As Long

   cli.Name = vbc.Name
   cli.Type = vbc.Type
   
   countOfCodeLines = vbc.CodeModule.CountOfDeclarationLines
   If countOfCodeLines < 3 Then
      'Eventuell nur Prozeduren vorhanden und nach den Options eine Leerzeile
      '=> zur Sicherheit den Code bis zur ersten Prozedur verwenden
      countOfCodeLines = getFirstProcLine(vbc.CodeModule) - 1
      If countOfCodeLines < 3 Then ' ist auch keine Prozedur vorhanden? => gesamten Code verwenden
         countOfCodeLines = vbc.CodeModule.CountOfLines
      End If
   End If
   If countOfCodeLines = 0 Then 'kein Code vorhanden, der ausgewertet werden könnte
      Exit Sub
   End If
   checkString = vbc.CodeModule.Lines(1, countOfCodeLines)
   checkString = findSubString(checkString, SEARCHSTRING_CODELIB_BEGIN, SEARCHSTRING_CODELIB_END)
   
   If Len(checkString) > 0 Then
      getCodeLibInfoRepositoryFile cli, checkString
      getCodeLibInfoRepositoryFileReplacement cli, checkString
      getCodeLibInfoLicenseFile cli, checkString
      getCodeLibInfoDependency cli, checkString
      getCodeLibInfoReferences cli, checkString
   End If

End Sub

Private Sub getCodeLibInfoFromAccessObject(ByRef cli As CodeLibInfo, _
               ByRef aoType As AcObjectType, ByRef aoName As String)

#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If
   Dim vbcName As String
   Dim tempForm As Access.Form
   Dim tempReport As Access.Report
   
   Select Case aoType
      Case AcObjectType.acModule
         vbcName = aoName
   
      Case AcObjectType.acForm, AcObjectType.acReport
         vbcName = "Form_" & Replace(aoName, " ", "_")
         If Not CodeModuleExits(vbcName) Then
            vbcName = vbNullString
         End If
         
      Case AcObjectType.acReport
         vbcName = "Report_" & Replace(aoName, " ", "_")
         If Not CodeModuleExits(vbcName) Then
            vbcName = vbNullString
         End If

      Case Else ' ... ???
      
   End Select
   
   If Len(vbcName) > 0 Then
      Set vbc = CurrentVbProject.VBComponents(vbcName)
      getCodeLibInfoFromVbComponent cli, vbc
      If Len(cli.RepositoryFile) > 0 Then
         cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert
   End If
   
#If DEBUGMODE Then
   Debug.Print aoName; " --> "; cli.LocalFile
#End If

End Sub

Private Function CodeModuleExits(ByVal CodeModuleName As String) As Boolean
   
#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If

   For Each vbc In CurrentVbProject.VBComponents
      Debug.Print vbc.Name, CodeModuleName
      If vbc.Name = CodeModuleName Then
         CodeModuleExits = True
         Exit Function
      End If
   Next
   
End Function

Public Function GetCodeLibInfoFromFilePath(ByVal FilePath As String, Optional ByVal FindDependency As Boolean = True) As CodeLibInfo
   Dim cli As CodeLibInfo
   getCodeLibInfoFromFile cli, fso.GetFile(FilePath), FindDependency
   GetCodeLibInfoFromFilePath = cli
End Function


Private Sub getCodeLibInfoFromFile(ByRef cli As CodeLibInfo, ByVal InputFile As Object, _
                          Optional ByVal FindDependency As Boolean = True, _
                          Optional ByVal FindTests As Boolean = True)
   
   Dim Pos As Long
   Dim checkString As String
   Dim TempString As String
   Dim i As Long
   Dim fileNumber As Long

   fileNumber = FreeFile
   Open InputFile.Path For Binary Access Read As fileNumber
   checkString = String$(LOF(fileNumber), 0)
   Get fileNumber, , checkString
   Close fileNumber
   
   'Namen ermitteln
   cli.Name = findSubString(checkString, SEARCHSTRING_ATTRIBUTNAME_BEGIN, SEARCHSTRING_ATTRIBUTNAME_END, Pos)
   
   If Len(cli.Name) = 0 Then ' aus Dateinamen
      cli.Name = Left$(InputFile.Name, InStrRev(InputFile.Name, ".") - 1)
   End If
   
   'Typ bestimmen
   If Pos > 0 Then ' da SEARCHSTRING_ATTRIBUTNAME_BEGIN gefunden wurde,
                   ' ist das definitiv ein Module oder eine Klasse
      cli.Type = clet_StdModule
      If InStr(1, Left$(checkString, Pos - 1), "CLASS", vbBinaryCompare) Then
         cli.Type = clet_ClassModule
      End If
   Else ' nach "BEGIN FORM" bzw "BEGIN REPORT" suchen (steht immer in der 4. Zeile)
      
      'SaveAsText bei Ac2007-Formularen ergibt unicode =>
      'checkString = StrConv(checkString, vbFromUnicode)

'/**
' * @todo Auslesen der Daten aus Ac07-Formluaren (accdb-Format), die mit SaveAsText erstellt wurden (unicode!)
' *       Wie findet man heraus, ob der String als Unicode ankommt?
' *       Reicht die Erkennung über die ersten 2 Zeichen [chr(255) und chr(254)] aus?
'**/

      Pos = 1
      For i = 1 To 3
         Pos = InStr(Pos, checkString, vbNewLine, vbBinaryCompare) + 1
         If Pos <= 0 Then Exit For
      Next
      If Pos > 1 Then
         TempString = findSubString(Mid$(checkString, Pos - 1), vbNewLine, vbNewLine)
      End If
      
      If TempString = SEARCHSTRING_FORMIDENTIFER Then
         cli.Type = clet_Form
      ElseIf TempString = SEARCHSTRING_REPORTIDENTIFER Then
         cli.Type = clet_Report
   
      Else 'aus Dateiendung
      
         Select Case Mid$(InputFile.Name, InStrRev(InputFile.Name, ".") + 1)
            Case "frm", "acf"
               cli.Type = clet_Form
            Case "rep", "acr"
               cli.Type = clet_Report
            Case "bas", "acm"
               cli.Type = clet_StdModule
            Case "cls"
               cli.Type = clet_ClassModule
            Case Else
               '
         End Select
   
      End If

   End If
   
   'CODELIB-Block auslesen
   checkString = findSubString(checkString, SEARCHSTRING_CODELIB_BEGIN, SEARCHSTRING_CODELIB_END)
   If Len(checkString) > 0 Then
      getCodeLibInfoRepositoryFile cli, checkString
      getCodeLibInfoRepositoryFileReplacement cli, checkString
      getCodeLibInfoLicenseFile cli, checkString
      If FindDependency Then getCodeLibInfoDependency cli, checkString
      getCodeLibInfoReferences cli, checkString
      If FindTests Then getCodeLibInfoTests cli, checkString
      getCodeLibInfoDescription cli, checkString
      getCodeLibInfoExecuteList cli, checkString
   End If

End Sub

Private Sub getCodeLibInfoRepositoryFile(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   cli.RepositoryFile = Replace(findSubString(sSourceString, SEARCHSTRING_FILE_BEGIN, SEARCHSTRING_FILE_END), "\", "/")
End Sub

Private Sub getCodeLibInfoRepositoryFileReplacement(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   cli.RepositoryFileReplacement = Replace(findSubString(sSourceString, SEARCHSTRING_REPLACE_BEGIN, SEARCHSTRING_REPLACE_END), "\", "/")
End Sub

Private Sub getCodeLibInfoLicenseFile(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   cli.LicenseFile = Replace(findSubString(sSourceString, SEARCHSTRING_LICENSE_BEGIN, SEARCHSTRING_LICENSE_END), "\", "/")
End Sub

Private Sub getCodeLibInfoDescription(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   cli.Description = Replace(findSubString(sSourceString, SEARCHSTRING_DESCRIPTION_BEGIN, SEARCHSTRING_DESCRIPTION_END), "\", "/")
End Sub

Private Sub getCodeLibInfoExecuteList(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
    Dim lst() As String

    If getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_EXECUTE_BEGIN, SEARCHSTRING_EXECUTE_END, lst) > 0 Then
        cli.ExecuteList = lst
    End If
End Sub

Private Sub getCodeLibInfoDependency(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
'Private Const SEARCHSTRING_USE_BEGIN As String = "<use>"
'Private Const SEARCHSTRING_USE_END   As String = "</use>"
   
   Dim dep() As String

   If getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_USE_BEGIN, SEARCHSTRING_USE_END, dep) > 0 Then
      cli.Dependency = dep
   End If
  
End Sub

Private Sub getCodeLibInfoTests(ByRef cli As CodeLibInfo, ByRef sSourceString As String)

   Dim tests() As String

   If getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_TEST_BEGIN, SEARCHSTRING_TEST_END, tests) > 0 Then
      cli.TestFiles = tests
   End If
 
End Sub

Private Function getCodelibInfoTagStrings(ByRef sSourceString As String, _
                                    ByRef sStartTag As String, ByRef sEndTag As String, _
                                    ByRef TagStrings() As String) As Long
   
   Dim startPos As Long
   Dim endPos As Long
   
   Dim tagString As String
   Dim tagIdx As Long

   tagIdx = -1
   
   startPos = InStr(1, sSourceString, sStartTag)
   Do While startPos > 0
      startPos = startPos + Len(sStartTag)
      endPos = InStr(startPos, sSourceString, sEndTag)
      If endPos = 0 Then
         Exit Do
      End If
      
      tagString = Trim$(Mid$(sSourceString, startPos, endPos - startPos))
      If Len(tagString) > 0 Then
         tagIdx = tagIdx + 1
         ReDim Preserve TagStrings(tagIdx)
         TagStrings(tagIdx) = tagString
      End If
   
      endPos = endPos + Len(sEndTag)
      startPos = InStr(endPos, sSourceString, sStartTag)
   
   Loop

   getCodelibInfoTagStrings = tagIdx + 1
 
End Function

Private Sub getCodeLibInfoReferences(ByRef cli As CodeLibInfo, ByRef sSourceString As String)

   Dim refString As String
   
   Dim refBlocks() As String
   Dim refBlockMaxIdx As Long
   Dim ref() As CodeLibInfoReference
   Dim i As Long

   refBlockMaxIdx = -1 + getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_REF_BEGIN, SEARCHSTRING_REF_END, refBlocks)
   
   If refBlockMaxIdx >= 0 Then
      ReDim ref(refBlockMaxIdx)
      For i = 0 To refBlockMaxIdx
         refString = refBlocks(i)
         ref(i).Name = findSubString(refString, SEARCHSTRING_REF_NAME_BEGIN, SEARCHSTRING_REF_NAME_END)
         ref(i).Major = Val(findSubString(refString, SEARCHSTRING_REF_MAJOR_BEGIN, SEARCHSTRING_REF_MAJOR_END))
         ref(i).Minor = Val(findSubString(refString, SEARCHSTRING_REF_MINOR_BEGIN, SEARCHSTRING_REF_MINOR_END))
         ref(i).GUID = findSubString(refString, SEARCHSTRING_REF_GUID_BEGIN, SEARCHSTRING_REF_GUID_END)
      Next
      
      cli.References = ref
      
   End If
 
End Sub


Private Function findSubString(ByRef SourceString As String, _
                               ByRef StartTag As String, ByRef EndTag As String, _
                               Optional ByRef LocatedStartPos As Long, Optional ByRef LocatedEndPos As Long) As String

   Dim startPos As Long, endPos As Long

   startPos = InStr(1, SourceString, StartTag, vbTextCompare)
   If startPos <= 0 Then
      Exit Function
   End If
   
   startPos = startPos + Len(StartTag)
   endPos = InStr(startPos, SourceString, EndTag, vbTextCompare)
   If endPos <= 0 Then
      Exit Function
   End If
   
   LocatedStartPos = startPos
   LocatedEndPos = endPos
   findSubString = Trim$(Mid$(SourceString, startPos, endPos - startPos))
  
End Function

Private Sub createMissingFolder(ByRef FullPath As String)
   
   Dim TempPath As String

   If Len(Dir$(FullPath, vbDirectory)) = 0 Then
      TempPath = getParentFolderPath(FullPath)
      createMissingFolder TempPath
      MkDir FullPath
   End If
  
End Sub

Private Function getParentFolderPath(ByRef FullPath As String) As String

   getParentFolderPath = Left$(FullPath, InStrRev(FullPath, "\", , vbTextCompare) - 1)
 
End Function

'---------------------------------------------------------------------------------------
' Sub: RefreshAll (Josef Pötzl, 2009-12-16)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle in der Anwendung vorhandenen Codelib-Elemente neu importieren
' </summary>
' <param name="Param"></param>
' <returns></returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RefreshAll(Optional ByVal ImportMode As CodeLibImportMode = _
                               CodeLibImportMode.clim_ImportMissingItems, _
                      Optional ByVal ShowMessageAfterRefresh As Boolean = False)
   

   refreshAllAccessObjects AcObjectType.acDefault, ImportMode, ShowMessageAfterRefresh

End Sub

'---------------------------------------------------------------------------------------
' Sub: RefreshAllModules (Josef Pötzl, 2009-12-16)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle in der Anwendung vorhandenen Code-Module mit Codelib-Blcok neu importieren
' (nur Module und Klassen, keine Formulare oder Berichte)
' </summary>
' <param name="ImportMode">Import-Modus (Stanard: clim_ImportAllUsedItems)</param>
' <param name="ShowMessageAfterRefresh">Messagebox nach Abschluss der Aktualiserung anzeigen</param>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RefreshAllModules(Optional ByVal ImportMode As CodeLibImportMode = _
                                      CodeLibImportMode.clim_ImportMissingItems, _
                             Optional ByVal ShowMessageAfterRefresh As Boolean = False)

   refreshAllAccessObjects AcObjectType.acModule, ImportMode, ShowMessageAfterRefresh
 
End Sub

Private Sub refreshAllAccessObjects(ByVal aoType As AcObjectType, _
                                    ByVal ImportMode As CodeLibImportMode, _
                                    ByVal ShowMessageAfterRefresh As Boolean)
   
'/**
' * @todo Regelung finden, falls FileCollection noch Einträge enthält
'**/

   If CurrentFileCollection.Count > 0 Then
      'leeren
      Set m_ImportFileCollection = New Collection
   End If
   
   Select Case aoType
      Case AcObjectType.acDefault
         addAccessObjectCollectionToRefreshFileCollection acModule, ImportMode
         addAccessObjectCollectionToRefreshFileCollection acForm, ImportMode
         addAccessObjectCollectionToRefreshFileCollection acReport, ImportMode
      Case Else
         addAccessObjectCollectionToRefreshFileCollection aoType, ImportMode
   End Select
   
   importFilesFromImportCollection
   
   If ShowMessageAfterRefresh Then
      MsgBox "Aktualisierung ist abgeschlossen"
   End If

End Sub

Private Sub addAccessObjectCollectionToRefreshFileCollection( _
                        ByRef aoType As AcObjectType, _
                        ByRef ImportMode As CodeLibImportMode)

   Dim aoCol As Object ' AllModules, ...
   Dim ao As AccessObject

   Select Case aoType
      Case AcObjectType.acModule
         Set aoCol = CurrentProject.AllModules
      Case AcObjectType.acForm
         Set aoCol = CurrentProject.AllForms
      Case AcObjectType.acReport
         Set aoCol = CurrentProject.AllReports
      Case Else 'Mit Fehler Abbrechen
         Err.Raise vbObjectError, "refreshAccessObjects", "Nur Module, Formular und Berichte können aktualisiert werden"
         Exit Sub
   End Select

   For Each ao In aoCol
      addAccessObjectToRefreshFileCollection ao, ImportMode
   Next

End Sub

Private Sub addAccessObjectToRefreshFileCollection(ByRef ao As AccessObject, ByRef ImportMode As CodeLibImportMode)
   
   Dim cli As CodeLibInfo
   Dim TempFile As Object

   getCodeLibInfoFromAccessObject cli, ao.Type, ao.Name
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      Set TempFile = fso.GetFile(cli.LocalFile)
      addMissingFile TempFile, ImportMode
   End If

End Sub

#If EARLYBINDING = 1 Then
Private Function getFirstProcLine(cm As VBIDE.CodeModule) As Long
#Else
Private Function getFirstProcLine(cm As Object) As Long
#End If

   Dim lngLine As Long
   Dim procKind As vbext_ProcKind
   Dim procName As String

   For lngLine = cm.CountOfDeclarationLines + 1 To cm.CountOfLines
      procName = cm.ProcOfLine(lngLine, procKind)
      If Len(procName) > 0 Then
         getFirstProcLine = cm.ProcBodyLine(procName, procKind)
         Exit For
      End If
   Next

End Function

#If EARLYBINDING = 1 Then
Public Sub InsertCodeLibInfo(ByRef cm As VBIDE.CodeModule, cli As CodeLibInfo)
#Else
Public Sub InsertCodeLibInfo(ByRef cm As Object, cli As CodeLibInfo)
#End If

   Dim tempCli As CodeLibInfo
   Dim clTagArray() As String
   Dim ref() As CodeLibInfoReference
   Dim lngDepCnt As Long
   Dim i As Long
   Dim strCodeLibInfo As String
   Dim ExecuteList() As String

   strCodeLibInfo = _
      "'<codelib>" & vbNewLine & _
      "'  <file>" & cli.RepositoryFile & "</file>" & vbNewLine
      
   '<replace>-Zeile
   If Len(cli.RepositoryFileReplacement) > 0 Then
      strCodeLibInfo = strCodeLibInfo & _
      "'  <replace>" & cli.RepositoryFileReplacement & "</replace>" & vbNewLine
   End If

   '<use>-Blöcke
   clTagArray = cli.Dependency
   If Not ((0 / 1) + (Not Not clTagArray) = 0) Then
      For i = LBound(clTagArray) To UBound(clTagArray)
         strCodeLibInfo = strCodeLibInfo & "'  <use>" & clTagArray(i) & "</use>" & vbNewLine
      Next
   End If
   
   '<ref>-Blöcke
   ref = cli.References
   If Not ((0 / 1) + (Not Not ref) = 0) Then
      For i = LBound(ref) To UBound(ref)
         strCodeLibInfo = strCodeLibInfo & "'  <ref>" & _
         "<name>" & ref(i).Name & "</name>" & _
         "<major>" & ref(i).Major & "</major>" & _
         "<minor>" & ref(i).Minor & "</minor>" & _
         "<guid>" & ref(i).GUID & "</guid>" & _
         "</ref>" & vbNewLine
      Next
   End If
   
   '<test>-Blöcke
   clTagArray = cli.TestFiles
   If Not ((0 / 1) + (Not Not clTagArray) = 0) Then
      For i = LBound(clTagArray) To UBound(clTagArray)
         strCodeLibInfo = strCodeLibInfo & "'  <test>" & clTagArray(i) & "</test>" & vbNewLine
      Next
   End If
   
   '<execute>-Blöcke
   ExecuteList = cli.ExecuteList
   If Not ((0 / 1) + (Not Not ExecuteList) = 0) Then
      For i = LBound(ExecuteList) To UBound(ExecuteList)
         strCodeLibInfo = strCodeLibInfo & "'  <execute>" & clTagArray(i) & "</execute>" & vbNewLine
      Next
   End If
   
   
   strCodeLibInfo = strCodeLibInfo & _
      "'</codelib>" & vbNewLine
   
#If DEBUGMODE Then
   Debug.Print strCodeLibInfo
#End If

End Sub
