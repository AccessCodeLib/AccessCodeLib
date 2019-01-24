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

Private Const EXTENSION_KEY As String = "ACLibFileManager"

' Integrierte Erweiterungen
Private Const EXTENSION_KEY_ACLIBCONFIG As String = "ACLibConfiguration"
Private Const EXTENSION_PROPNAME_LOCALREPOSITORYROOT As String = "LocalRepositoryRoot"
Private Const EXTENSION_PROPNAME_PRIVATEREPOSITORYROOT As String = "PrivateRepositoryRoot"

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

Private Const REPOSTITORY_ROOT_CODE_APPLICATIONROOT As String = "%AppFolder%"
Private Const REPOSTITORY_ROOT_CODE_PRIVATEROOT As String = "%PrivateRoot%"

Private m_ImportFileCollection As Collection
Private m_ReplacedFilesCollection As Collection
Private m_CLI As CodeLibInfo


#If EARLYBINDING = 1 Then
Private m_FSO As FileSystemObject
#Else
Private m_FSO As Object
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

Public Property Set ApplicationHandlerRef(ByRef ObjRef As Object) ' as ApplicationHandler
   Set m_ApplicationHandler = ObjRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = EXTENSION_KEY
End Property

Private Sub Class_Terminate()
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal ExtensionKeyToCheck As String, ByRef Exists As Boolean)
   If ExtensionKeyToCheck = EXTENSION_KEY Then Exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal ExtensionKeyToCheck As String, ByRef ExtensionReference As Object)
   If ExtensionKeyToCheck = EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal ExtensionKeyToCheck As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   If ExtensionKeyToCheck = EXTENSION_KEY Then
      GetExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
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

Private Sub GetExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         ResumeMessage = "Property '" & PropertyName & "' is not supported-"
         
   End Select
   
End Sub

Private Sub Dispose()
On Error Resume Next
   Set m_ImportFileCollection = Nothing
   Set m_FSO = Nothing
   Set m_CurrentVbProject = Nothing
End Sub

Public Sub ExportAll()
   ExportAllModules
   ExportAllForms
   ExportAllReports
End Sub

Public Sub ExportAllModules(Optional ExportComponentsWithoutCodeLibInfo As Boolean = False)

   Dim ao As AccessObject

   ActivateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllModules
      ExportVbComponent ao.Name, ExportComponentsWithoutCodeLibInfo
   Next
   
#If DEBUGMODE Then
   Debug.Print "Module-Export abgeschlossen"
#End If

End Sub

Public Sub ExportAllForms()

   Dim ao As AccessObject

   If CheckDisallowExportFormReport() Then
      Exit Sub
   End If

   ActivateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllForms
      ExportAccessObject AcObjectType.acForm, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Formular-Export abgeschlossen"
#End If

End Sub

Private Function CheckDisallowExportFormReport() As Boolean

   If Val(CurrentDb.Version) >= 12 Then
      MsgBox "Formular und Berichte ab dem Access2007-Format dürfen nicht exportiert werden." & _
             "Siehe Bug #16"
      CheckDisallowExportFormReport = True
   End If

End Function

Public Sub ExportAllReports()

   Dim ao As AccessObject

   ActivateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllReports
      ExportAccessObject AcObjectType.acReport, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Bericht-Export abgeschlossen"
#End If

End Sub

Public Sub ImportAllFilesFromRepository(Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                                        Optional ByVal ImportTestFiles As Boolean = False)

   Dim FolderCol As Collection
#If EARLYBINDING Then
   Dim CheckFolder As Folder
   Dim TempFolder As Folder
   Dim TempFile As File
#Else
   Dim CheckFolder As Object
   Dim TempFolder As Object
   Dim TempFile As Object
#End If
   
   Set FolderCol = New Collection
   Set m_ImportFileCollection = New Collection
   
   Set TempFolder = fso.GetFolder(LocalRepositoryRootDirectory)
   FolderCol.Add TempFolder, TempFolder.Path
   
   Do While FolderCol.Count > 0
      
      Set CheckFolder = FolderCol(1)
      
      For Each TempFile In CheckFolder.Files
         If Not IgnoreFile(TempFile) Then
            AddMissingFile TempFile, ImportMode
         End If
      Next
      
      For Each TempFolder In CheckFolder.SubFolders
         If Not IgnoreFolder(TempFolder) Then
            FolderCol.Add TempFolder, TempFolder.Path
         End If
      Next
   
      FolderCol.Remove 1
   
   Loop
   Set TempFile = Nothing
   Set FolderCol = Nothing
   
   ImportFilesFromImportCollection True, ImportTestFiles

End Sub

Public Sub ImportRepositoryFile(ByVal RepositoryPath As String, _
                       Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                       Optional ByVal ImportTestFiles As Boolean = False)
   
   Dim PathString As String

   PathString = GetRepositoryFullPath(RepositoryPath)
   
   Dim TempFile As Object
   Set TempFile = fso.GetFile(PathString)
   AddMissingFile TempFile, ImportMode
   Set TempFile = Nothing
   
   ImportFilesFromImportCollection True, ImportTestFiles

End Sub

Public Sub ImportRepositoryFiles(ByRef RepositoryPathArray() As String, _
                        Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                        Optional ByVal ImportTestFiles As Boolean = False)
   
   Dim ArraySize As Long
   Dim i As Long
   Dim PathString As String
   Dim TempFile As Object

   ArraySize = UBound(RepositoryPathArray)
   For i = LBound(RepositoryPathArray) To ArraySize - 1
      PathString = GetRepositoryFullPath(RepositoryPathArray(i))
      Set TempFile = fso.GetFile(PathString)
      AddMissingFile TempFile, ImportMode
   Next
   Set TempFile = Nothing
   
   ImportFilesFromImportCollection True, ImportTestFiles

End Sub

Private Sub ImportFilesFromImportCollection( _
                                   Optional ByRef CompileAfterImport As Boolean = True, _
                                   Optional ByVal ImportTestFiles As Boolean = False)
   
   Dim TempFile As Object
   Dim FileImportMode As CodeLibImportMode
   Dim i As Long
   Dim MaxCnt As Long
   
   Dim ColItem As Variant
   
   ' eine VB-Modul aus CurrentVbProject aktivieren, sonst wirkt RunCommand acCmdCompileAndSaveAllModules
   ' möglicherweise im falschen VBProject
   ' zur Sicherheit mit On error resume next, weil es nicht unbedingt notwendig ist und den Ablauf nicht stört
   On Error Resume Next
   ActivateCurrentProject
   On Error GoTo 0
   If CurrentProject.AllModules.Count > 0 Then
      DoCmd.RunCommand acCmdCompileAndSaveAllModules
   End If
    
    MaxCnt = m_ImportFileCollection.Count
    
    '/*
    ' * TODO: Bestimmung der Anzahl alle Dateien (inkl. Abhängigkeiten) vor dem Import
    ' * TODO: Umbau AccessProgressBar auf einmaligen Durchlauf (Init nur einmal)
    ' */
   
    i = 1
    Do While i <= MaxCnt
        ColItem = m_ImportFileCollection(i)
        Set TempFile = ColItem(0)
        FileImportMode = ColItem(1)
        
        AccessProgressBar.Init "Importiere " & TempFile & "...", 2, 1
        AccessProgressBar.PerformStep
        
        ImportFile TempFile, FileImportMode, ImportTestFiles
        
        AccessProgressBar.PerformStep
        MaxCnt = m_ImportFileCollection.Count
        i = i + 1
    Loop



   If CompileAfterImport Then
      ActivateCurrentProject
      CurrentProject.Application.DoCmd.RunCommand acCmdCompileAndSaveAllModules
   End If
   
   'Executes ausführen
   If (0 / 1) + (Not Not m_CLI.ExecuteList) Then
      AccessProgressBar.Init "Executes ausführen...", UBound(m_CLI.ExecuteList) + 1, 1
      For i = 0 To UBound(m_CLI.ExecuteList)
        AccessProgressBar.PerformStep
        If StringTools.Contains(m_CLI.ExecuteList(i), REPOSTITORY_ROOT_CODE_PRIVATEROOT) Then
            Eval VBA.Strings.Replace(m_CLI.ExecuteList(i), REPOSTITORY_ROOT_CODE_PRIVATEROOT, LocalRepositoryRootDirectory())
        Else
            Eval (m_CLI.ExecuteList(i))
        End If
      Next
      If AccessProgressBar.IsInitialized Then AccessProgressBar.Clear
   End If
   
   'Modul-Variablen aufräumen, da Import-Vorgang an dieser Stelle abgeschlossen ist
   Set m_ImportFileCollection = Nothing
   Set m_ReplacedFilesCollection = Nothing
   Set m_FSO = Nothing
   
#If DEBUGMODE Then
   Debug.Print "Import abgeschlossen"
#End If

End Sub


Private Function IgnoreFolder(ByRef TestFolder As Object) As Boolean
'/*
' * @todo Ausschlussliste für Verzeichnisse flexibler gestalten
'**/

   Select Case TestFolder.Name
      Case ".svn", "_codelib"
         IgnoreFolder = True
      Case Else
         '
   End Select

End Function

Private Function IgnoreFile(ByRef TestFile As Object) As Boolean
'/*
' * @todo Ausschlussliste für Dateien flexibler gestalten
'**/

   Select Case TestFile.Name
      Case "_info.txt", "_config_Application.bas"
         IgnoreFile = True
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
   Dim Proj As VBProject
#Else
   Dim Proj As Object
#End If
   Dim CurrentDbName As String

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Prüfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      CurrentDbName = UncPath(CurrentDb.Name)
      If m_CurrentVbProject.FileName <> CurrentDbName Then
         Set m_CurrentVbProject = Nothing
         For Each Proj In VBE.VBProjects
            If Proj.FileName = CurrentDbName Then
               Set m_CurrentVbProject = Proj
               Exit For
            End If
         Next
      End If
   End If
   
   Set CurrentVbProject = m_CurrentVbProject

End Property

Private Sub ActivateCurrentProject()

   Dim Activated As Boolean
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
            Activated = True
            Exit For
         End If
      Next
      'falls nur vbext_ct_Document vorhanden sind, das erste verwenden
      If Not Activated Then
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
      Set m_FSO = CreateObject("Scripting.FileSystemObject")
   End If
   Set fso = m_FSO

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

Private Sub AddReplacedFilePath(ByVal RepFilePath As String)

   Dim TempFile As Object
   Dim TempFilePath As Variant
   Dim i As Long
   Dim MaxCount As Long
   Dim col As Collection

   MaxCount = CurrentFileCollection.Count
   For i = 1 To MaxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set TempFile = m_ImportFileCollection(i)(0)
      If TempFile.Path = RepFilePath Then
         m_ImportFileCollection.Remove TempFile.Path
         Exit For
      End If
   Next
   Set TempFile = Nothing
   
   Set col = CurrentReplacedFilesCollection
   For Each TempFilePath In col
      If TempFilePath = RepFilePath Then
         Exit Sub
      End If
   Next
   
   CurrentReplacedFilesCollection.Add CVar(RepFilePath), RepFilePath

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
         m_LocalRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLIBCONFIG, EXTENSION_PROPNAME_LOCALREPOSITORYROOT, vbNullString)
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
         m_PrivateRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLIBCONFIG, EXTENSION_PROPNAME_PRIVATEREPOSITORYROOT, vbNullString)
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

   Dim RepPath As String

   ReleativPath = Replace(ReleativPath, "/", "\")
   
   If Left(ReleativPath, Len(REPOSTITORY_ROOT_CODE_APPLICATIONROOT)) = REPOSTITORY_ROOT_CODE_APPLICATIONROOT Then
      RepPath = CurrentProject.Path & "\"
      ReleativPath = Mid$(ReleativPath, Len(REPOSTITORY_ROOT_CODE_APPLICATIONROOT) + 1)
   
   ElseIf Left(ReleativPath, Len(REPOSTITORY_ROOT_CODE_PRIVATEROOT)) = REPOSTITORY_ROOT_CODE_PRIVATEROOT Then
      RepPath = PrivateRepositoryRootDirectory
      If Len(RepPath) = 0 Then
         Err.Raise vbObjectError, "getRepositoryFullPath", "Wert für privates Root-Verzeichnise fehlt (PrivateRepositoryRootDirectory = '')."
         Exit Function
      End If
      ReleativPath = Mid$(ReleativPath, Len(REPOSTITORY_ROOT_CODE_PRIVATEROOT) + 1)
      
   Else
      RepPath = LocalRepositoryRootDirectory
      If Len(RepPath) = 0 Then
         Err.Raise vbObjectError, "getRepositoryFullPath", "Wert für lokales Root-Verzeichnise fehlt (LocalRepositoryRootDirectory = '')."
         Exit Function
      End If
   End If
   
   Do While Left$(ReleativPath, 1) = "\"
      ReleativPath = Mid$(ReleativPath, 2)
   Loop

   GetRepositoryFullPath = RepPath & ReleativPath

End Function

Private Sub ImportFile(ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode, _
              Optional ByVal ImportTestFiles As Boolean = False)
   
   Dim i As Long
   Dim TempFile As Object
   Dim FilePath As String
   Dim CancelImport As Integer
   
   GetCodeLibInfoFromFile m_CLI, ImportFile, (ImportMode <> clim_ImportSelectedOnly)
   
   If Len(m_CLI.RepositoryFileReplacement) > 0 Then
      AddReplacedFilePath GetRepositoryFullPath(m_CLI.RepositoryFileReplacement)
   End If
   
   'Abbruch ermöglichen
   RaiseEvent ImportRepositoryFile(m_CLI.Name, m_CLI.RepositoryFile, m_CLI.Type, m_CLI.Dependency, ImportFile, CancelImport)
   If CancelImport Then
      Exit Sub
   End If
   
   'Lizenz (LICENSE-Block)
   If Len(m_CLI.LicenseFile) > 0 Then
      FilePath = GetRepositoryFullPath(m_CLI.LicenseFile)
      Set TempFile = fso.GetFile(FilePath)
      AddMissingFile TempFile, clim_ImportMissingItems
   End If

   'Erforderliche Module (USE-Block)
   If ImportMode <> clim_ImportSelectedOnly Then
      If (0 / 1) + (Not Not m_CLI.Dependency) Then
      For i = 0 To UBound(m_CLI.Dependency)
         FilePath = GetRepositoryFullPath(m_CLI.Dependency(i))
         If fso.FileExists(FilePath) Then
            Set TempFile = fso.GetFile(FilePath)
            AddMissingFile TempFile, ImportMode
         End If
      Next
      End If
   End If
   
   'zuerst Verweise
   If (0 / 1) + (Not Not m_CLI.References) Then
      For i = 0 To UBound(m_CLI.References)
         AddMissingReference m_CLI.References(i)
      Next
   End If
   
   'dann Codemodul
   Select Case m_CLI.Type
      Case CodeLibElementType.clet_ClassModule, CodeLibElementType.clet_StdModule
         ImportVbComponent m_CLI, ImportFile, ImportMode
      Case CodeLibElementType.clet_Form
         ImportAccessObject acForm, m_CLI, ImportFile, ImportMode
      Case CodeLibElementType.clet_Report
         ImportAccessObject acReport, m_CLI, ImportFile, ImportMode
      Case Else
         ' eventuell Fehler auslösen?
   End Select
      
   'eventuell die Tests
   If ImportTestFiles Then
   If (0 / 1) + (Not Not m_CLI.TestFiles) Then
      For i = 0 To UBound(m_CLI.TestFiles)
         FilePath = GetRepositoryFullPath(m_CLI.TestFiles(i))
         If fso.FileExists(FilePath) Then
            Set TempFile = fso.GetFile(FilePath)
            AddMissingFile TempFile, ImportMode
         End If
      Next
   End If
   End If
   
#If DEBUGMODE Then
   Debug.Print ImportFile.Path; " --> "; m_CLI.Name
#End If
  
End Sub

Private Sub AddMissingReference(ByRef NewCodeLibRef As CodeLibInfoReference)
   
   Dim Refs As Access.References
   Dim ref As Access.Reference
   Dim ReplaceRef As Boolean

   Set Refs = Access.References
   For Each ref In Refs
      If ref.Name = NewCodeLibRef.Name Then
         If ref.Major < NewCodeLibRef.Major Then
            ReplaceRef = True
         ElseIf ref.Major = NewCodeLibRef.Major And ref.Minor < NewCodeLibRef.Minor Then
            ReplaceRef = True
         Else 'Nicht ersetzen, Abbruch
            Exit Sub
         End If
         Exit For
      End If
   Next
   
   If ReplaceRef Then
      Refs.Remove ref
   End If
   
   Refs.AddFromGuid NewCodeLibRef.GUID, NewCodeLibRef.Major, NewCodeLibRef.Minor
 
End Sub

Private Sub AddMissingFile(ByVal UsedFile As Object, ByVal ImportMode As CodeLibImportMode)
   
   Dim FileColItem As Variant
   Dim TempFile As Object
   Dim i As Long
   Dim MaxCount As Long

   If IsReplacedFile(UsedFile) Then Exit Sub

   MaxCount = CurrentFileCollection.Count
   For i = 1 To MaxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set TempFile = m_ImportFileCollection(i)(0)
      If TempFile.Path = UsedFile.Path Then
         Exit Sub
      End If
   Next

   FileColItem = Array(UsedFile, ImportMode)
   m_ImportFileCollection.Add FileColItem, UsedFile.Path
 
End Sub


Private Sub ExportAccessObject(ByRef AoType As AcObjectType, ByRef AoName As String)

   Select Case AoType
      Case AcObjectType.acModule
         ExportVbComponent AoName
      Case AcObjectType.acForm, AcObjectType.acReport
         ExportAccessObjectCli AoType, AoName
      Case Else
         '
   End Select

End Sub

Private Sub ExportAccessObjectCli(ByRef AoType As AcObjectType, ByRef AoName As String)

'/**
' * @todo Was ist mit Formularen und Berichten ohne Modul?
'**/
   
   Dim cli As CodeLibInfo

   GetCodeLibInfoFromAccessObject cli, AoType, AoName
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      CreateMissingFolder GetParentFolderPath(cli.LocalFile)
      Application.SaveAsText AoType, AoName, cli.LocalFile
      CleanAccessBugExport cli.LocalFile
   End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert

#If DEBUGMODE Then
   Debug.Print AoName; " --> ";
   If Len(cli.RepositoryFile) > 0 Then
      Debug.Print cli.LocalFile
   Else
      Debug.Print "nicht exportiert"
   End If
#End If
 
End Sub

Private Sub CleanAccessBugExport(ByVal FilePath As String)
   
   Dim IsAcc2010 As Boolean
   
   On Error Resume Next
   If Application.Version >= "12.0" Then
      IsAcc2010 = True
   End If
   On Error GoTo 0
   If Not IsAcc2010 Then
      Exit Sub
   End If
   
   Dim CheckString As String
   Dim TempString As String
   Dim FileNumber As Long

   FileNumber = FreeFile
   Open FilePath For Binary Access Read As FileNumber
   CheckString = String$(LOF(FileNumber), 0)
   Get FileNumber, , CheckString
   Close FileNumber
   
   TempString = Replace(CheckString, "NoSaveCTIWhenDisabled =1", vbNullString)
   If StrComp(TempString, CheckString, vbBinaryCompare) = 0 Then
      Exit Sub
   End If
   
   Kill FilePath
   FileNumber = FreeFile
   Open FilePath For Binary Access Write As FileNumber
      Put FileNumber, , TempString
   Close FileNumber
   
   
End Sub

Private Sub ExportVbComponent(ByRef VbcName As String, Optional ExportComponentsWithoutCodeLibInfo As Boolean = False)

#If EARLYBINDING Then
   Dim vbc As VBIDE.VBComponent
   Dim cm As VBIDE.CodeModule
#Else
   Dim vbc As Object
   Dim cm As Object
#End If

   Dim cli As CodeLibInfo
  
   Set vbc = CurrentVbProject.VBComponents(VbcName)
   
   'Leerzeilen am Ende säubern
   Set cm = vbc.CodeModule
   Do While Len(Trim$(cm.Lines(cm.CountOfLines, 1))) = 0 And cm.CountOfLines > 1
      cm.DeleteLines cm.CountOfLines
   Loop
   
   GetCodeLibInfoFromVbComponent cli, vbc
   
   If ExportComponentsWithoutCodeLibInfo Then
   If Len(cli.RepositoryFile) = 0 Then
      cli.RepositoryFile = REPOSTITORY_ROOT_CODE_APPLICATIONROOT & "/CodeExport/" & GenerateRepositoryFileName(cli.Name, cli.Type)
   End If
   End If
   
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      CreateMissingFolder GetParentFolderPath(cli.LocalFile)
      vbc.Export cli.LocalFile
   End If
   
#If DEBUGMODE Then
   Debug.Print VbcName; " --> "; cli.LocalFile
#End If

End Sub

Private Function GenerateRepositoryFileName(ByVal CodeModuleName As String, ByVal CodeModuleType As CodeLibElementType) As String
   
   Dim FileExt As String
   
   Select Case CodeModuleType
      Case CodeLibElementType.clet_ClassModule
         FileExt = ".cls"
      Case CodeLibElementType.clet_StdModule
         FileExt = ".bas"
      Case CodeLibElementType.clet_Form
         FileExt = ".frm"
      Case CodeLibElementType.clet_Report
         FileExt = ".rpt"
      Case Else
         Err.Raise vbObjectError, "ACLibFileManager.GenerateRepositoryFileName", "CodeModuleType " & CodeModuleType & " not supported"
   End Select
   
   GenerateRepositoryFileName = CodeModuleName & FileExt
   
End Function

Private Sub ImportVbComponent(ByRef CodeLibInf As CodeLibInfo, ByRef ImportFile As Object, ByRef FileImportMode As CodeLibImportMode)
   
#If EARLYBINDING = 1 Then
   Dim VbcCol As VBIDE.VBComponents
   Dim vbc As VBIDE.VBComponent
   Dim cm As VBIDE.CodeModule
#Else
   Dim VbcCol As Object
   Dim vbc As Object
   Dim cm As Object
#End If
   Dim CodeModuleExists As Boolean

   If Len(Trim$(CodeLibInf.Name)) = 0 Then
      Err.Raise vbObjectError, "ADF_ImportExport.importVbComponent", "Es wurde versucht ein Module ohne Namen zu erstellen."
      Exit Sub
   End If
   
   Set VbcCol = CurrentVbProject.VBComponents
   For Each vbc In VbcCol
      If vbc.Name = CodeLibInf.Name Then
         If vbc.Type <> CodeLibInf.Type Then
            ' trotz clim_ImportMissingItems importieren
         ElseIf FileImportMode = clim_ImportMissingItems Then 'Abbruch, damit nicht überschrieben wird
            Exit Sub
         ElseIf vbc.Name = MODULNAME_CONFIG_APPLICATION Then ' _config_Application niemals überschreiben, wenn es schon existiert
            Exit Sub
         End If
         CodeModuleExists = True
         Exit For
      End If
   Next

   If CodeModuleExists Then   ' überprüfen, ob Typen übereinstimmen
      If vbc.Type <> CodeLibInf.Type Then
         VbcCol.Remove vbc
         Set vbc = Nothing
         CodeModuleExists = False
      End If
   End If

   If CodeModuleExists Then ' Inhalt über CodeModule ändern, damit MS-Add-In für Quellcodeverwaltung keine Schwierigkeiten macht
      Set cm = vbc.CodeModule
      cm.DeleteLines 1, cm.CountOfLines
      cm.AddFromFile ImportFile.Path
      If CodeLibInf.Type = CodeLibElementType.clet_ClassModule Then
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
      
   Else
      VbcCol.import ImportFile.Path
   End If

End Sub

#If EARLYBINDING = 1 Then
Private Sub DeleteClassHeaderLinesFromCodeModul(ByRef cm As CodeModule)
#Else
Private Sub DeleteClassHeaderLinesFromCodeModul(ByRef cm As Object)
#End If
'
'Header sieht so aus:
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

Private Sub ImportAccessObject(ByRef AcOjbType As AcObjectType, ByRef CodeLibInf As CodeLibInfo, _
                               ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode)
   
   Dim col As Object ' AllForms oder AllReports
   Dim AcObj As AccessObject

   If Len(Trim$(CodeLibInf.Name)) = 0 Then
      Err.Raise vbObjectError, "ADF_ImportExport.importVbComponent", "Es wurde versucht ein Module ohne Namen zu erstellen."
      Exit Sub
   End If
   
   Select Case AcOjbType
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
   If CodeLibInf.Name = "ACLibImportWizardForm" Then
      'immer umbennenen, da sonst bei ungeöffneten Formular das Formular des Add-Ins gelöscht wird!
      CodeLibInf.Name = CodeLibInf.Name & "_NEU"
   End If

   For Each AcObj In col
      If AcObj.Name = CodeLibInf.Name Then
         If ImportMode = clim_ImportMissingItems Then 'Abbruch, damit nicht überschrieben wird
            Exit Sub
         End If
         If AcObj.IsLoaded Then
            DoCmd.Close AcOjbType, CodeLibInf.Name, acSaveNo
         End If
         Set AcObj = Nothing
         DoCmd.DeleteObject AcOjbType, CodeLibInf.Name
         Exit For
      End If
   Next
   Set col = Nothing
   
   Application.LoadFromText AcOjbType, CodeLibInf.Name, ImportFile.Path
   DoCmd.Close AcOjbType, CodeLibInf.Name

   'Sonderfall: ACLibImportWizardForm
   If CodeLibInf.Name = "ACLibImportWizardForm_NEU" Then
      MsgBox "Bitte das Formular ""ACLibImportWizardForm_NEU"" manuell in ""ACLibImportWizardForm"" umbennenen." & vbNewLine & _
             "Es konnte nicht mit dem Originalnamen erstellt werden, da der Import-Assistent ein gleichnamiges Formular geöffnet hat."
   End If

End Sub

#If EARLYBINDING = 1 Then
Public Function GetCodeLibInfoFromCodeModule(ByVal CodeModulName As String, Optional ByVal VbProjectRef As VBProject) As CodeLibInfo
#Else
Public Function GetCodeLibInfoFromCodeModule(ByVal CodeModulName As String, Optional ByVal VbProjectRef As Object) As CodeLibInfo
#End If

   Dim cli As CodeLibInfo
#If EARLYBINDING = 1 Then
   Dim vbc As VBIDE.VBComponent
#Else
   Dim vbc As Object
#End If
   
   If VbProjectRef Is Nothing Then
      Set VbProjectRef = CurrentVbProject
   End If
   
   Set vbc = VbProjectRef.VBComponents(CodeModulName)
   
   GetCodeLibInfoFromVbComponent cli, vbc
   
   GetCodeLibInfoFromCodeModule = cli
   
End Function

Private Sub GetCodeLibInfoFromVbComponent(ByRef CodeLibInf As CodeLibInfo, ByRef VbComp As Object)
        
   Dim CheckString As String
   Dim CountOfCodeLines As Long

   CodeLibInf.Name = VbComp.Name
   CodeLibInf.Type = VbComp.Type
   
   CountOfCodeLines = VbComp.CodeModule.CountOfDeclarationLines
   If CountOfCodeLines < 3 Then
      'Eventuell nur Prozeduren vorhanden und nach den Options eine Leerzeile
      '=> zur Sicherheit den Code bis zur ersten Prozedur verwenden
      CountOfCodeLines = GetFirstProcLine(VbComp.CodeModule) - 1
      If CountOfCodeLines < 3 Then ' ist auch keine Prozedur vorhanden? => gesamten Code verwenden
         CountOfCodeLines = VbComp.CodeModule.CountOfLines
      End If
   End If
   If CountOfCodeLines = 0 Then 'kein Code vorhanden, der ausgewertet werden könnte
      Exit Sub
   End If
   CheckString = VbComp.CodeModule.Lines(1, CountOfCodeLines)
   CheckString = FindSubString(CheckString, SEARCHSTRING_CODELIB_BEGIN, SEARCHSTRING_CODELIB_END)
   
   If Len(CheckString) > 0 Then
      GetCodeLibInfoRepositoryFile CodeLibInf, CheckString
      GetCodeLibInfoRepositoryFileReplacement CodeLibInf, CheckString
      GetCodeLibInfoLicenseFile CodeLibInf, CheckString
      GetCodeLibInfoDependency CodeLibInf, CheckString
      GetCodeLibInfoReferences CodeLibInf, CheckString
   End If

End Sub

Private Sub GetCodeLibInfoFromAccessObject(ByRef CodeLibInf As CodeLibInfo, _
               ByRef AoType As AcObjectType, ByRef AoName As String)

#If EARLYBINDING Then
   Dim vbc As VBComponent
#Else
   Dim vbc As Object
#End If
   Dim VbcName As String
   
   Select Case AoType
      Case AcObjectType.acModule
         VbcName = AoName
   
      Case AcObjectType.acForm, AcObjectType.acReport
         VbcName = "Form_" & Replace(AoName, " ", "_")
         If Not CodeModuleExits(VbcName) Then
            VbcName = vbNullString
         End If
         
      Case AcObjectType.acReport
         VbcName = "Report_" & Replace(AoName, " ", "_")
         If Not CodeModuleExits(VbcName) Then
            VbcName = vbNullString
         End If

      Case Else ' ... ???
      
   End Select
   
   If Len(VbcName) > 0 Then
      Set vbc = CurrentVbProject.VBComponents(VbcName)
      GetCodeLibInfoFromVbComponent CodeLibInf, vbc
      If Len(CodeLibInf.RepositoryFile) > 0 Then
         CodeLibInf.LocalFile = GetRepositoryFullPath(CodeLibInf.RepositoryFile)
      End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert
   End If
   
#If DEBUGMODE Then
   Debug.Print AoName; " --> "; CodeLibInf.LocalFile
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
   GetCodeLibInfoFromFile cli, fso.GetFile(FilePath), FindDependency
   GetCodeLibInfoFromFilePath = cli
End Function


Private Sub GetCodeLibInfoFromFile(ByRef CodeLibInf As CodeLibInfo, ByVal InputFile As Object, _
                          Optional ByVal FindDependency As Boolean = True, _
                          Optional ByVal FindTests As Boolean = True)
   
   Dim Pos As Long
   Dim CheckString As String
   Dim TempString As String
   Dim i As Long
   Dim FileNumber As Long

   FileNumber = FreeFile
   Open InputFile.Path For Binary Access Read As FileNumber
   CheckString = String$(LOF(FileNumber), 0)
   Get FileNumber, , CheckString
   Close FileNumber
   
   'Namen ermitteln
   CodeLibInf.Name = FindSubString(CheckString, SEARCHSTRING_ATTRIBUTNAME_BEGIN, SEARCHSTRING_ATTRIBUTNAME_END, Pos)
   
   If Len(CodeLibInf.Name) = 0 Then ' aus Dateinamen
      CodeLibInf.Name = Left$(InputFile.Name, InStrRev(InputFile.Name, ".") - 1)
   End If
   
   'Typ bestimmen
   If Pos > 0 Then ' da SEARCHSTRING_ATTRIBUTNAME_BEGIN gefunden wurde,
                   ' ist das definitiv ein Module oder eine Klasse
      CodeLibInf.Type = clet_StdModule
      If InStr(1, Left$(CheckString, Pos - 1), "CLASS", vbBinaryCompare) Then
         CodeLibInf.Type = clet_ClassModule
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
         Pos = InStr(Pos, CheckString, vbNewLine, vbBinaryCompare) + 1
         If Pos <= 0 Then Exit For
      Next
      If Pos > 1 Then
         TempString = FindSubString(Mid$(CheckString, Pos - 1), vbNewLine, vbNewLine)
      End If
      
      If TempString = SEARCHSTRING_FORMIDENTIFER Then
         CodeLibInf.Type = clet_Form
      ElseIf TempString = SEARCHSTRING_REPORTIDENTIFER Then
         CodeLibInf.Type = clet_Report
   
      Else 'aus Dateiendung
      
         Select Case Mid$(InputFile.Name, InStrRev(InputFile.Name, ".") + 1)
            Case "frm", "acf"
               CodeLibInf.Type = clet_Form
            Case "rep", "acr"
               CodeLibInf.Type = clet_Report
            Case "bas", "acm"
               CodeLibInf.Type = clet_StdModule
            Case "cls"
               CodeLibInf.Type = clet_ClassModule
            Case Else
               '
         End Select
   
      End If

   End If
   
   'CODELIB-Block auslesen
   CheckString = FindSubString(CheckString, SEARCHSTRING_CODELIB_BEGIN, SEARCHSTRING_CODELIB_END)
   If Len(CheckString) > 0 Then
      GetCodeLibInfoRepositoryFile CodeLibInf, CheckString
      GetCodeLibInfoRepositoryFileReplacement CodeLibInf, CheckString
      GetCodeLibInfoLicenseFile CodeLibInf, CheckString
      If FindDependency Then GetCodeLibInfoDependency CodeLibInf, CheckString
      GetCodeLibInfoReferences CodeLibInf, CheckString
      If FindTests Then GetCodeLibInfoTests CodeLibInf, CheckString
      GetCodeLibInfoDescription CodeLibInf, CheckString
      GetCodeLibInfoExecuteList CodeLibInf, CheckString
   End If

End Sub

Private Sub GetCodeLibInfoRepositoryFile(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
   CodeLibInf.RepositoryFile = Replace(FindSubString(SourceString, SEARCHSTRING_FILE_BEGIN, SEARCHSTRING_FILE_END), "\", "/")
End Sub

Private Sub GetCodeLibInfoRepositoryFileReplacement(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
   CodeLibInf.RepositoryFileReplacement = Replace(FindSubString(SourceString, SEARCHSTRING_REPLACE_BEGIN, SEARCHSTRING_REPLACE_END), "\", "/")
End Sub

Private Sub GetCodeLibInfoLicenseFile(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
   CodeLibInf.LicenseFile = Replace(FindSubString(SourceString, SEARCHSTRING_LICENSE_BEGIN, SEARCHSTRING_LICENSE_END), "\", "/")
End Sub

Private Sub GetCodeLibInfoDescription(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
   CodeLibInf.Description = Replace(FindSubString(SourceString, SEARCHSTRING_DESCRIPTION_BEGIN, SEARCHSTRING_DESCRIPTION_END), "\", "/")
End Sub

Private Sub GetCodeLibInfoExecuteList(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
    Dim lst() As String

    If GetCodelibInfoTagStrings(SourceString, SEARCHSTRING_EXECUTE_BEGIN, SEARCHSTRING_EXECUTE_END, lst) > 0 Then
        CodeLibInf.ExecuteList = lst
    End If
End Sub

Private Sub GetCodeLibInfoDependency(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)
'Private Const SEARCHSTRING_USE_BEGIN As String = "<use>"
'Private Const SEARCHSTRING_USE_END   As String = "</use>"
   
   Dim dep() As String

   If GetCodelibInfoTagStrings(SourceString, SEARCHSTRING_USE_BEGIN, SEARCHSTRING_USE_END, dep) > 0 Then
      CodeLibInf.Dependency = dep
   End If
  
End Sub

Private Sub GetCodeLibInfoTests(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)

   Dim Tests() As String

   If GetCodelibInfoTagStrings(SourceString, SEARCHSTRING_TEST_BEGIN, SEARCHSTRING_TEST_END, Tests) > 0 Then
      CodeLibInf.TestFiles = Tests
   End If
 
End Sub

Private Function GetCodelibInfoTagStrings(ByRef SourceString As String, _
                                    ByRef StartTag As String, ByRef EndTag As String, _
                                    ByRef TagStrings() As String) As Long
   
   Dim StartPos As Long
   Dim EndPos As Long
   
   Dim TagString As String
   Dim TagIdx As Long

   TagIdx = -1
   
   StartPos = InStr(1, SourceString, StartTag)
   Do While StartPos > 0
      StartPos = StartPos + Len(StartTag)
      EndPos = InStr(StartPos, SourceString, EndTag)
      If EndPos = 0 Then
         Exit Do
      End If
      
      TagString = Trim$(Mid$(SourceString, StartPos, EndPos - StartPos))
      If Len(TagString) > 0 Then
         TagIdx = TagIdx + 1
         ReDim Preserve TagStrings(TagIdx)
         TagStrings(TagIdx) = TagString
      End If
   
      EndPos = EndPos + Len(EndTag)
      StartPos = InStr(EndPos, SourceString, StartTag)
   
   Loop

   GetCodelibInfoTagStrings = TagIdx + 1
 
End Function

Private Sub GetCodeLibInfoReferences(ByRef CodeLibInf As CodeLibInfo, ByRef SourceString As String)

   Dim RefString As String
   
   Dim RefBlocks() As String
   Dim RefBlockMaxIdx As Long
   Dim ref() As CodeLibInfoReference
   Dim i As Long

   RefBlockMaxIdx = -1 + GetCodelibInfoTagStrings(SourceString, SEARCHSTRING_REF_BEGIN, SEARCHSTRING_REF_END, RefBlocks)
   
   If RefBlockMaxIdx >= 0 Then
      ReDim ref(RefBlockMaxIdx)
      For i = 0 To RefBlockMaxIdx
         RefString = RefBlocks(i)
         ref(i).Name = FindSubString(RefString, SEARCHSTRING_REF_NAME_BEGIN, SEARCHSTRING_REF_NAME_END)
         ref(i).Major = Val(FindSubString(RefString, SEARCHSTRING_REF_MAJOR_BEGIN, SEARCHSTRING_REF_MAJOR_END))
         ref(i).Minor = Val(FindSubString(RefString, SEARCHSTRING_REF_MINOR_BEGIN, SEARCHSTRING_REF_MINOR_END))
         ref(i).GUID = FindSubString(RefString, SEARCHSTRING_REF_GUID_BEGIN, SEARCHSTRING_REF_GUID_END)
      Next
      
      CodeLibInf.References = ref
      
   End If
 
End Sub


Private Function FindSubString(ByRef SourceString As String, _
                               ByRef StartTag As String, ByRef EndTag As String, _
                               Optional ByRef LocatedStartPos As Long, Optional ByRef LocatedEndPos As Long) As String

   Dim StartPos As Long, EndPos As Long

   StartPos = InStr(1, SourceString, StartTag, vbTextCompare)
   If StartPos <= 0 Then
      Exit Function
   End If
   
   StartPos = StartPos + Len(StartTag)
   EndPos = InStr(StartPos, SourceString, EndTag, vbTextCompare)
   If EndPos <= 0 Then
      Exit Function
   End If
   
   LocatedStartPos = StartPos
   LocatedEndPos = EndPos
   FindSubString = Trim$(Mid$(SourceString, StartPos, EndPos - StartPos))
  
End Function

Private Sub CreateMissingFolder(ByRef FullPath As String)
   
   Dim TempPath As String

   If Len(Dir$(FullPath, vbDirectory)) = 0 Then
      TempPath = GetParentFolderPath(FullPath)
      CreateMissingFolder TempPath
      MkDir FullPath
   End If
  
End Sub

Private Function GetParentFolderPath(ByRef FullPath As String) As String

   GetParentFolderPath = Left$(FullPath, InStrRev(FullPath, "\", , vbTextCompare) - 1)
 
End Function

'---------------------------------------------------------------------------------------
' Sub: RefreshAll
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle in der Anwendung vorhandenen Codelib-Elemente neu importieren
' </summary>
' <param name="ImportMode">Import-Modus (Standard: clim_ImportAllUsedItems)</param>
' <param name="ShowMessageAfterRefresh">Messagebox nach Abschluss der Aktualiserung anzeigen</param>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RefreshAll(Optional ByVal ImportMode As CodeLibImportMode = _
                               CodeLibImportMode.clim_ImportAllUsedItems, _
                      Optional ByVal ShowMessageAfterRefresh As Boolean = False)
   

   RefreshAllAccessObjects AcObjectType.acDefault, ImportMode, ShowMessageAfterRefresh

End Sub

'---------------------------------------------------------------------------------------
' Sub: RefreshAllModules
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Alle in der Anwendung vorhandenen Code-Module mit Codelib-Blcok neu importieren
' (nur Module und Klassen, keine Formulare oder Berichte)
' </summary>
' <param name="ImportMode">Import-Modus (Standard: clim_ImportAllUsedItems)</param>
' <param name="ShowMessageAfterRefresh">Messagebox nach Abschluss der Aktualiserung anzeigen</param>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub RefreshAllModules(Optional ByVal ImportMode As CodeLibImportMode = _
                                      CodeLibImportMode.clim_ImportAllUsedItems, _
                             Optional ByVal ShowMessageAfterRefresh As Boolean = False)

   RefreshAllAccessObjects AcObjectType.acModule, ImportMode, ShowMessageAfterRefresh
 
End Sub

Private Sub RefreshAllAccessObjects(ByVal AoType As AcObjectType, _
                                    ByVal ImportMode As CodeLibImportMode, _
                                    ByVal ShowMessageAfterRefresh As Boolean)
   
'/**
' * @todo Regelung finden, falls FileCollection noch Einträge enthält
'**/

   If CurrentFileCollection.Count > 0 Then
      'leeren
      Set m_ImportFileCollection = New Collection
   End If
   
   Select Case AoType
      Case AcObjectType.acDefault
         AddAccessObjectCollectionToRefreshFileCollection acModule, ImportMode
         AddAccessObjectCollectionToRefreshFileCollection acForm, ImportMode
         AddAccessObjectCollectionToRefreshFileCollection acReport, ImportMode
      Case Else
         AddAccessObjectCollectionToRefreshFileCollection AoType, ImportMode
   End Select
   
   ImportFilesFromImportCollection
   
   If ShowMessageAfterRefresh Then
      MsgBox "Aktualisierung ist abgeschlossen"
   End If

End Sub

Private Sub AddAccessObjectCollectionToRefreshFileCollection( _
                        ByRef AoType As AcObjectType, _
                        ByRef ImportMode As CodeLibImportMode)

   Dim AoCol As Object ' AllModules, ...
   Dim ao As AccessObject

   Select Case AoType
      Case AcObjectType.acModule
         Set AoCol = CurrentProject.AllModules
      Case AcObjectType.acForm
         Set AoCol = CurrentProject.AllForms
      Case AcObjectType.acReport
         Set AoCol = CurrentProject.AllReports
      Case Else 'Mit Fehler Abbrechen
         Err.Raise vbObjectError, "refreshAccessObjects", "Nur Module, Formular und Berichte können aktualisiert werden"
         Exit Sub
   End Select

   For Each ao In AoCol
      AddAccessObjectToRefreshFileCollection ao, ImportMode
   Next

End Sub

Private Sub AddAccessObjectToRefreshFileCollection(ByRef AcObj As AccessObject, ByRef ImportMode As CodeLibImportMode)
   
   Dim cli As CodeLibInfo
   Dim TempFile As Object

   GetCodeLibInfoFromAccessObject cli, AcObj.Type, AcObj.Name
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = GetRepositoryFullPath(cli.RepositoryFile)
      Set TempFile = fso.GetFile(cli.LocalFile)
      AddMissingFile TempFile, ImportMode
   End If

End Sub

#If EARLYBINDING = 1 Then
Private Function GetFirstProcLine(ByRef CodMod As VBIDE.CodeModule) As Long
#Else
Private Function GetFirstProcLine(ByRef CodMod As Object) As Long
#End If

   Dim LineNum As Long
   Dim ProcKind As vbext_ProcKind
   Dim ProcName As String

   For LineNum = CodMod.CountOfDeclarationLines + 1 To CodMod.CountOfLines
      ProcName = CodMod.ProcOfLine(LineNum, ProcKind)
      If Len(ProcName) > 0 Then
         GetFirstProcLine = CodMod.ProcBodyLine(ProcName, ProcKind)
         Exit For
      End If
   Next

End Function

#If EARLYBINDING = 1 Then
Public Sub InsertCodeLibInfo(ByRef CodMod As VBIDE.CodeModule, ByRef CodeLibInf As CodeLibInfo)
#Else
Public Sub InsertCodeLibInfo(ByRef CodMod As Object, ByRef CodeLibInf As CodeLibInfo)
#End If

   Dim ClTagArray() As String
   Dim ref() As CodeLibInfoReference
   Dim i As Long
   Dim CodeLibInfoText As String
   Dim ExecuteList() As String

   CodeLibInfoText = _
      "'<codelib>" & vbNewLine & _
      "'  <file>" & CodeLibInf.RepositoryFile & "</file>" & vbNewLine
      
   '<replace>-Zeile
   If Len(CodeLibInf.RepositoryFileReplacement) > 0 Then
      CodeLibInfoText = CodeLibInfoText & _
      "'  <replace>" & CodeLibInf.RepositoryFileReplacement & "</replace>" & vbNewLine
   End If

   '<use>-Blöcke
   ClTagArray = CodeLibInf.Dependency
   If Not ((0 / 1) + (Not Not ClTagArray) = 0) Then
      For i = LBound(ClTagArray) To UBound(ClTagArray)
         CodeLibInfoText = CodeLibInfoText & "'  <use>" & ClTagArray(i) & "</use>" & vbNewLine
      Next
   End If
   
   '<ref>-Blöcke
   ref = CodeLibInf.References
   If Not ((0 / 1) + (Not Not ref) = 0) Then
      For i = LBound(ref) To UBound(ref)
         CodeLibInfoText = CodeLibInfoText & "'  <ref>" & _
         "<name>" & ref(i).Name & "</name>" & _
         "<major>" & ref(i).Major & "</major>" & _
         "<minor>" & ref(i).Minor & "</minor>" & _
         "<guid>" & ref(i).GUID & "</guid>" & _
         "</ref>" & vbNewLine
      Next
   End If
   
   '<test>-Blöcke
   ClTagArray = CodeLibInf.TestFiles
   If Not ((0 / 1) + (Not Not ClTagArray) = 0) Then
      For i = LBound(ClTagArray) To UBound(ClTagArray)
         CodeLibInfoText = CodeLibInfoText & "'  <test>" & ClTagArray(i) & "</test>" & vbNewLine
      Next
   End If
   
   '<execute>-Blöcke
   ExecuteList = CodeLibInf.ExecuteList
   If Not ((0 / 1) + (Not Not ExecuteList) = 0) Then
      For i = LBound(ExecuteList) To UBound(ExecuteList)
         CodeLibInfoText = CodeLibInfoText & "'  <execute>" & ClTagArray(i) & "</execute>" & vbNewLine
      Next
   End If
   
   CodeLibInfoText = CodeLibInfoText & _
      "'</codelib>" & vbNewLine
   
#If DEBUGMODE Then
   Debug.Print CodeLibInfoText
#End If

End Sub
