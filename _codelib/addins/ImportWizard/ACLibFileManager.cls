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
' Klassenmodul: ACLibFileManager (Josef Pötzl, 2009-12-11)
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
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <use>base/ApplicationHandler.cls</use>
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
Option Compare Database
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
Public Event PropertyMissingLocalRepositoryRootDirectory(ByRef newValue As String)
Public Event ImportRepositoryFile(ByVal ObjectName As String, ByVal RepositoryFile As String, _
                                  ByVal ElementType As CodeLibElementType, ByRef Dependency As Variant, _
                                  ByRef ImportFile As Object, ByRef Cancel As Integer)


'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As Object) ' as ApplicationHandler
On Error GoTo HandleErr

   Set m_ApplicationHandler = objRef

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ApplicationHandlerRef", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Property

Public Property Get ExtensionKey() As String
On Error Resume Next
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
On Error Resume Next
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
On Error GoTo HandleErr

   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_ApplicationHandler_ExtensionLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)
On Error GoTo HandleErr

   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "m_ApplicationHandler_ExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
End Sub


' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                              ByRef ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
On Error Resume Next
   Set m_ApplicationHandler = Nothing
End Sub



'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ACLibFileManager
'---------------------------------------------------------------------------------------

Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, _
                                       ByRef ResumeMode As ApplicationHandlerResumeModes, _
                                       ByRef ResumeMessage As Variant)
   
On Error GoTo HandleErr

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName

      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         
   End Select

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getExtensionPropertyLookup", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub Dispose()
On Error Resume Next
   Set m_ImportFileCollection = Nothing
   Set m_fso = Nothing
   Set m_CurrentVbProject = Nothing
End Sub

Public Sub ExportAll()

On Error GoTo HandleErr

   ExportAllModules
   ExportAllForms
   ExportAllReports

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ExportAll", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Sub ExportAllModules()

   Dim ao As AccessObject
   
On Error GoTo HandleErr

   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllModules
      exportVbComponent ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Module-Export abgeschlossen"
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ExportAllModules", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Sub ExportAllForms()

   Dim ao As AccessObject
   
On Error GoTo HandleErr

   If checkDisallowExportFormReport() Then
      Exit Sub
   End If

   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllForms
      exportAccessObject AcObjectType.acForm, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Formular-Export abgeschlossen"
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ExportAllForms", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function checkDisallowExportFormReport() As Boolean
   
On Error GoTo HandleErr

   If Val(CurrentDb.Version) >= 12 Then
      MsgBox "Formular und Berichte ab dem Access2007-Format dürfen nicht exportiert werden." & _
             "Siehe Bug #16"
      checkDisallowExportFormReport = True
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "checkDisallowExportFormReport", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Public Sub ExportAllReports()

   Dim ao As AccessObject
   
On Error GoTo HandleErr

   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   For Each ao In CurrentProject.AllReports
      exportAccessObject AcObjectType.acReport, ao.Name
   Next
   
#If DEBUGMODE Then
   Debug.Print "Bericht-Export abgeschlossen"
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ExportAllReports", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Sub ImportAllFilesFromRepository(Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                                        Optional ByVal bImportTestFiles As Boolean = False)

   Dim folderCol As Collection
   
On Error GoTo HandleErr

#If EARLYBINDING Then
   Dim checkFolder As Folder
   Dim tempFolder As Folder
   Dim tempFile As file
#Else
   Dim checkFolder As Object
   Dim tempFolder As Object
   Dim tempFile As Object
#End If
   
   Set folderCol = New Collection
   Set m_ImportFileCollection = New Collection
   
   Set tempFolder = fso.GetFolder(LocalRepositoryRootDirectory)
   folderCol.Add tempFolder, tempFolder.Path
   
   Do While folderCol.Count > 0
      
      Set checkFolder = folderCol(1)
      
      For Each tempFile In checkFolder.Files
         If Not ignoreFile(tempFile) Then
            addMissingFile tempFile, ImportMode
         End If
      Next
      
      For Each tempFolder In checkFolder.SubFolders
         If Not ignoreFolder(tempFolder) Then
            folderCol.Add tempFolder, tempFolder.Path
         End If
      Next
   
      folderCol.Remove 1
   
   Loop
   Set tempFile = Nothing
   Set folderCol = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ImportAllFiles", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Public Sub ImportRepositoryFile(ByVal RepositoryPath As String, _
                       Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                       Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim pathString As String
   
On Error GoTo HandleErr

   pathString = getRepositoryFullPath(RepositoryPath)
   
   Dim tempFile As Object
   Set tempFile = fso.GetFile(pathString)
   addMissingFile tempFile, ImportMode
   Set tempFile = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ImportRepositoryFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Sub ImportRepositoryFiles(ByRef RepositoryPathArray() As String, _
                        Optional ByVal ImportMode As CodeLibImportMode = clim_ImportMissingItems, _
                        Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim arraySize As Long
   Dim i As Long
   Dim pathString As String
   Dim tempFile As Object
   
On Error GoTo HandleErr

   arraySize = UBound(RepositoryPathArray)
   For i = LBound(RepositoryPathArray) To arraySize
      pathString = getRepositoryFullPath(RepositoryPathArray(i))
      Set tempFile = fso.GetFile(pathString)
      addMissingFile tempFile, ImportMode
   Next
   Set tempFile = Nothing
   
   importFilesFromImportCollection True, bImportTestFiles

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ImportRepositoryFiles", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub importFilesFromImportCollection( _
                                   Optional ByRef CompileAfterImport As Boolean = True, _
                                   Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim tempFile As Object
   Dim fileImportMode As CodeLibImportMode
   Dim i As Long
   Dim maxCnt As Long
   
   Dim colItem As Variant
   
   ' eine VB-Modul aus CurrentVbProject aktivieren, sonst wirkt RunCommand acCmdCompileAndSaveAllModules
   ' möglicherweise im falschen VBProject
   ' zur Sicherheit mit On error resume next, weil es nicht unbedingt notwendig ist und den Ablauf nicht stört
On Error GoTo HandleErr

   activateCurrentProject
   DoCmd.RunCommand acCmdCompileAndSaveAllModules
   
   maxCnt = m_ImportFileCollection.Count
   i = 1
   Do While i <= maxCnt
      'Set tempFile = m_ImportFileCollection(i)
      colItem = m_ImportFileCollection(i)
      Set tempFile = colItem(0)
      fileImportMode = colItem(1)
      ImportFile tempFile, fileImportMode, bImportTestFiles
      i = i + 1
      maxCnt = m_ImportFileCollection.Count
   Loop

   If CompileAfterImport Then
      activateCurrentProject
      CurrentProject.Application.DoCmd.RunCommand acCmdCompileAndSaveAllModules
   End If
   
   'Modul-Variablen aufräumen, da Import-Vorgang an dieser Stelle abgeschlossen ist
   Set m_ImportFileCollection = Nothing
   Set m_ReplacedFilesCollection = Nothing
   Set m_fso = Nothing
   
#If DEBUGMODE Then
   Debug.Print "Import abgeschlossen"
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.importFilesFromImportCollection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


Private Function ignoreFolder(ByRef oTestFolder As Object) As Boolean
'/*
' * @todo Ausschlussliste für Verzeichnisse flexibler gestalten
'**/
On Error GoTo HandleErr

   Select Case oTestFolder.Name
      Case ".svn", "_codelib"
         ignoreFolder = True
      Case Else
         '
   End Select

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ignoreFolder", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Function ignoreFile(ByRef oTestFile As Object) As Boolean
'/*
' * @todo Ausschlussliste für Dateien flexibler gestalten
'**/
On Error GoTo HandleErr

   Select Case oTestFile.Name
      Case "_info.txt", "_config_Application.bas"
         ignoreFile = True
      Case Else
         '
   End Select

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ignoreFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
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

On Error GoTo HandleErr

   If m_CurrentVbProject Is Nothing Then
      Set m_CurrentVbProject = VBE.ActiveVBProject
      'Pürfen, ob das richtige VbProject gewählt wurde (muss das von CurrentDb sein)
      strCurrentDbName = CurrentDb.Name
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
   Select Case HandleError(Err.Number, "ACLibFileManager.CurrentVbProject", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Sub activateCurrentProject()

On Error GoTo HandleErr

   On Error Resume Next
   If CurrentVbProject.VBComponents.Count > 0 Then
      CurrentVbProject.VBComponents(1).Activate
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.activateCurrentProject", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

#If EARLYBINDING Then
Private Property Get fso() As FileSystemObject
#Else
Private Property Get fso() As Object
#End If

   On Error GoTo HandleErr

   If fso Is Nothing Then
      'Set fso = New FileSystemObject
      Set m_fso = CreateObject("Scripting.FileSystemObject")
   End If
   Set fso = m_fso

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.fso", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Property Get CurrentFileCollection() As Collection

On Error GoTo HandleErr

   If m_ImportFileCollection Is Nothing Then
      Set m_ImportFileCollection = New Collection
   End If
   Set CurrentFileCollection = m_ImportFileCollection

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.CurrentFileCollection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Property Get CurrentReplacedFilesCollection() As Collection

On Error GoTo HandleErr

   If m_ReplacedFilesCollection Is Nothing Then
      Set m_ReplacedFilesCollection = New Collection
   End If
   Set CurrentReplacedFilesCollection = m_ReplacedFilesCollection

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.CurrentReplacedFilesCollection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

Private Sub addReplacedFilePath(ByVal repFilePath As String)

   Dim tempFile As Object
   Dim i As Long
   Dim maxCount As Long
   
   Dim col As Collection
   
On Error GoTo HandleErr

   maxCount = CurrentFileCollection.Count
   For i = 1 To maxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set tempFile = m_ImportFileCollection(i)(0)
      If tempFile.Path = repFilePath Then
         m_ImportFileCollection.Remove tempFile.Path
         Exit For
      End If
   Next
   Set tempFile = Nothing
   
   Set col = CurrentReplacedFilesCollection
   For Each tempFile In col
      If tempFile.Path = repFilePath Then
         Exit Sub
      End If
   Next
   Set tempFile = Nothing

   Set tempFile = fso.GetFile(repFilePath)
   CurrentReplacedFilesCollection.Add tempFile, tempFile.Path

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.addReplacmentFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function IsReplacedFile(ByRef RepFile As Object) As Boolean

   Dim tempFile As Object
   
On Error GoTo HandleErr

   For Each tempFile In CurrentReplacedFilesCollection
      If tempFile.Path = RepFile.Path Then
         IsReplacedFile = True
         Exit Function
      End If
   Next

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.IsReplacedFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Property Get LocalRepositoryRootDirectory() As String

On Error GoTo HandleErr

   If Len(m_LocalRepositoryRootDirectory) = 0 Then
      RaiseEvent PropertyMissingLocalRepositoryRootDirectory(m_LocalRepositoryRootDirectory)
      If Len(m_LocalRepositoryRootDirectory) = 0 Then
         m_LocalRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLibConfig, EXTENSION_PROPNAME_LocalRepositoryRoot, vbNullString)
      End If
   End If
   LocalRepositoryRootDirectory = m_LocalRepositoryRootDirectory

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.LocalRepositoryRootDirectory", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Property Let LocalRepositoryRootDirectory(ByVal LocalPath As String)

On Error GoTo HandleErr

   If Right$(LocalPath, 1) <> "\" Then
      LocalPath = LocalPath & "\"
   End If
   m_LocalRepositoryRootDirectory = LocalPath

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.LocalRepositoryRootDirectory", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Property Get PrivateRepositoryRootDirectory() As String

On Error GoTo HandleErr

   If Len(m_PrivateRepositoryRootDirectory) = 0 Then
      RaiseEvent PropertyMissingLocalRepositoryRootDirectory(m_PrivateRepositoryRootDirectory)
      If Len(m_PrivateRepositoryRootDirectory) = 0 Then
         m_PrivateRepositoryRootDirectory = CurrentApplication.GetExtensionProperty(EXTENSION_KEY_ACLibConfig, EXTENSION_PROPNAME_PrivateRepositoryRoot, vbNullString)
      End If
   End If
   PrivateRepositoryRootDirectory = m_PrivateRepositoryRootDirectory

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.LocalRepositoryRootDirectory", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Property Let PrivateRepositoryRootDirectory(ByVal PrivatePath As String)

On Error GoTo HandleErr

   If Right$(PrivatePath, 1) <> "\" Then
      PrivatePath = PrivatePath & "\"
   End If
   m_PrivateRepositoryRootDirectory = PrivatePath

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.LocalRepositoryRootDirectory", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Function getRepositoryFullPath(ByVal ReleativPath As String) As String

   Dim repPath As String
   
On Error GoTo HandleErr

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

   getRepositoryFullPath = repPath & ReleativPath

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getRepositoryFullPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub ImportFile(ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode, _
              Optional ByVal bImportTestFiles As Boolean = False)
   
   Dim cli As CodeLibInfo
   Dim i As Long
   Dim tempFile As Object
   Dim filePath As String
   Dim cancelImport As Integer
   
On Error GoTo HandleErr

   getCodeLibInfoFromFile cli, ImportFile, (ImportMode <> clim_ImportSelectedOnly)
   
   If Len(cli.RepositoryFileReplacement) > 0 Then
      addReplacedFilePath getRepositoryFullPath(cli.RepositoryFileReplacement)
   End If
   
   'Abbruch ermöglichen
   RaiseEvent ImportRepositoryFile(cli.Name, cli.RepositoryFile, cli.Type, cli.Dependency, ImportFile, cancelImport)
   If cancelImport Then
      Exit Sub
   End If
   
   'Lizenz (LICENSE-Block)
   If Len(cli.LicenseFile) > 0 Then
      filePath = getRepositoryFullPath(cli.LicenseFile)
      Set tempFile = fso.GetFile(filePath)
      addMissingFile tempFile, clim_ImportMissingItems
   End If
   
   'Erforderliche Module (USE-Block)
   If ImportMode <> clim_ImportSelectedOnly Then
      If (0 / 1) + (Not Not cli.Dependency) Then
      For i = 0 To UBound(cli.Dependency)
         filePath = getRepositoryFullPath(cli.Dependency(i))
         If fso.FileExists(filePath) Then
            Set tempFile = fso.GetFile(filePath)
            addMissingFile tempFile, ImportMode
         End If
      Next
      End If
   End If
   
   'zuerst Verweise
   If (0 / 1) + (Not Not cli.References) Then
      For i = 0 To UBound(cli.References)
         addMissingReference cli.References(i)
      Next
   End If
   
   'dann Codemodul
   Select Case cli.Type
      Case CodeLibElementType.clet_ClassModule, CodeLibElementType.clet_StdModule
         importVbComponent cli, ImportFile, ImportMode
      Case CodeLibElementType.clet_Form
         importAccessObject acForm, cli, ImportFile, ImportMode
      Case CodeLibElementType.clet_Report
         importAccessObject acReport, cli, ImportFile, ImportMode
      Case Else
         ' eventuell Fehler auslösen?
   End Select
   
   'eventuell die Tests
   If bImportTestFiles Then
   If (0 / 1) + (Not Not cli.TestFiles) Then
      For i = 0 To UBound(cli.TestFiles)
         filePath = getRepositoryFullPath(cli.TestFiles(i))
         If fso.FileExists(filePath) Then
            Set tempFile = fso.GetFile(filePath)
            addMissingFile tempFile, ImportMode
         End If
      Next
   End If
   End If
   
#If DEBUGMODE Then
   Debug.Print ImportFile.Path; " --> "; cli.Name
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ImportFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub addMissingReference(ByRef NewCodeLibRef As CodeLibInfoReference)
   
   Dim refs As Access.References
   Dim ref As Access.Reference
   Dim bolReplace As Boolean
   
On Error GoTo HandleErr

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

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "addMissingReference", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub addMissingFile(ByVal UsedFile As Object, ByVal ImportMode As CodeLibImportMode)
   
   Dim varfilColItem As Variant
   Dim tempFile As Object
   Dim i As Long
   Dim maxCount As Long
   
On Error GoTo HandleErr

   If IsReplacedFile(UsedFile) Then Exit Sub

   maxCount = CurrentFileCollection.Count
   For i = 1 To maxCount
      'tempFile = varfileColItem(0) =>
      'varfilColItem = m_ImportFileCollection(i) =>
      Set tempFile = m_ImportFileCollection(i)(0)
      If tempFile.Path = UsedFile.Path Then
         Exit Sub
      End If
   Next

   varfilColItem = Array(UsedFile, ImportMode)
   m_ImportFileCollection.Add varfilColItem, UsedFile.Path

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.addMissingUsedFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


Private Sub exportAccessObject(ByRef aoType As AcObjectType, ByRef aoName As String)

On Error GoTo HandleErr

   Select Case aoType
      Case AcObjectType.acModule
         exportVbComponent aoName
      Case AcObjectType.acForm, AcObjectType.acReport
         exportAccessObjectCli aoType, aoName
      Case Else
         '
   End Select

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.exportAccessObject", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub exportAccessObjectCli(ByRef aoType As AcObjectType, ByRef aoName As String)

'/**
' * @todo Was ist mit Formularen und Berichten ohne Modul?
'**/
   
   Dim cli As CodeLibInfo

On Error GoTo HandleErr

   getCodeLibInfoFromAccessObject cli, aoType, aoName
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = getRepositoryFullPath(cli.RepositoryFile)
      createMissingFolder getParentFolderPath(cli.LocalFile)
      Application.SaveAsText aoType, aoName, cli.LocalFile
   End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert

#If DEBUGMODE Then
   Debug.Print aoName; " --> ";
   If Len(cli.RepositoryFile) > 0 Then
      Debug.Print cli.LocalFile
   Else
      Debug.Print "nicht exportiert"
   End If
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.exportAccessObjectCli", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr
   
   Set vbc = CurrentVbProject.VBComponents(vbcName)
   
   'Leerzeilen am Ende säubern
   Set cm = vbc.CodeModule
   Do While Len(Trim$(cm.Lines(cm.countOfLines, 1))) = 0 And cm.countOfLines > 1
      cm.DeleteLines cm.countOfLines
   Loop
   
   getCodeLibInfoFromVbComponent cli, vbc
   
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = getRepositoryFullPath(cli.RepositoryFile)
      createMissingFolder getParentFolderPath(cli.LocalFile)
      vbc.Export cli.LocalFile
   End If
   
#If DEBUGMODE Then
   Debug.Print vbcName; " --> "; cli.LocalFile
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.exportVbComponent", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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

On Error GoTo HandleErr
   
   If Len(Trim$(cli.Name)) = 0 Then
      Err.Raise vbObjectError, "ADF_ImportExport.importVbComponent", "Es wurde versucht ein Module ohne Namen zu erstellen."
      Exit Sub
   End If
   
   Set vbcCol = CurrentVbProject.VBComponents
   For Each vbc In vbcCol
      If vbc.Name = cli.Name Then
         If fileImportMode = clim_ImportMissingItems Then 'Abbruch, damit nicht überschrieben wird
            Exit Sub
         ElseIf vbc.Name = MODULNAME_CONFIG_APPLICATION Then ' _config_Application niemals überschreiben, wenn es schon existiert
            Exit Sub
         End If
         codeModuleExists = True
         Exit For
      End If
   Next

   If codeModuleExists Then ' Inhalt über CodeModule ändern, damit MS-Add-In für Quellcodeverwaltung keine Schwierigkeiten macht
      Set cm = vbc.CodeModule
      cm.DeleteLines 1, cm.countOfLines
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
      Do While Len(Trim$(cm.Lines(cm.countOfLines, 1))) = 0 And cm.countOfLines > 1
         cm.DeleteLines cm.countOfLines
      Loop
      
   Else
      vbcCol.import ImportFile.Path
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.importVbComponent", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

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
   
On Error GoTo HandleErr

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

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "deleteClassHeaderLinesFromCodeModul", Err.Description, ACLibErrorHandlerMode.aclibErrIgnore)
   Case Else ' ohne Meldung aus Ersetzen aussteigen und restlichen Code fortsetzen
      Resume ExitHere
   End Select

End Sub

Private Sub importAccessObject(ByRef acojbType As AcObjectType, ByRef cli As CodeLibInfo, _
                               ByRef ImportFile As Object, ByRef ImportMode As CodeLibImportMode)
   
   Dim col As Object ' AllForms oder AllReports
   Dim acObj As AccessObject
   
On Error GoTo HandleErr

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

   'Sonderfall: ACLibImportWizardForm
   If cli.Name = "ACLibImportWizardForm_NEU" Then
      MsgBox "Bitte das Formular ""ACLibImportWizardForm_NEU"" manuell in ""ACLibImportWizardForm"" umbennenen." & vbNewLine & _
             "Es konnte nicht mit dem Originalnamen erstellt werden, da der Import-Assistent ein gleichnamiges Formular geöffnet hat."
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.importAccessObject", Err.Description, ACLibErrorHandlerMode.aclibErrMsgBox)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   Dim countOfLines As Long
        
On Error GoTo HandleErr

   cli.Name = vbc.Name
   cli.Type = vbc.Type
   
   countOfLines = vbc.CodeModule.CountOfDeclarationLines
   If countOfLines < 3 Then
      'Eventuell nur Prozeduren vorhanden und nach den Options eine Leerzeile
      '=> zur Sicherheit den Code bis zur ersten Prozedur verwenden
      countOfLines = getFirstProcLine(vbc.CodeModule) - 1
      If countOfLines < 3 Then ' ist auch keine Prozedur vorhanden? => gesamten Code verwenden
         countOfLines = vbc.CodeModule.countOfLines
      End If
   End If
   If countOfLines = 0 Then 'kein Code vorhanden, der ausgewertet werden könnte
      Exit Sub
   End If
   checkString = vbc.CodeModule.Lines(1, countOfLines)
   checkString = findSubString(checkString, SEARCHSTRING_CODELIB_BEGIN, SEARCHSTRING_CODELIB_END)
   
   If Len(checkString) > 0 Then
      getCodeLibInfoRepositoryFile cli, checkString
      getCodeLibInfoRepositoryFileReplacement cli, checkString
      getCodeLibInfoLicenseFile cli, checkString
      getCodeLibInfoDependency cli, checkString
      getCodeLibInfoReferences cli, checkString
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoFromVbComponent", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub getCodeLibInfoFromAccessObject(ByRef cli As CodeLibInfo, _
               ByRef aoType As AcObjectType, ByRef aoName As String)

On Error GoTo HandleErr

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
   
      Case AcObjectType.acForm
         DoCmd.OpenForm aoName, acDesign, , , , acHidden
         Set tempForm = Forms(aoName)
         If tempForm.HasModule Then
            vbcName = "Form_" & Replace(aoName, " ", "_")
         End If
         DoCmd.Close acForm, aoName, acSaveNo
         
      Case AcObjectType.acReport
         DoCmd.OpenReport aoName, acViewDesign, , , acHidden
         Set tempReport = Reports(aoName)
         If tempReport.HasModule Then
            vbcName = "Report_" & Replace(aoName, " ", "_")
         End If
         DoCmd.Close acForm, aoName, acSaveNo
         
      Case Else ' ... ???
      
   End Select
   
   If Len(vbcName) > 0 Then
      Set vbc = CurrentVbProject.VBComponents(vbcName)
      getCodeLibInfoFromVbComponent cli, vbc
      If Len(cli.RepositoryFile) > 0 Then
         cli.LocalFile = getRepositoryFullPath(cli.RepositoryFile)
      End If ' Rest ignorieren .. wenn im Code nicht steht, wird auch nicht exportiert
   End If
   
#If DEBUGMODE Then
   Debug.Print aoName; " --> "; cli.LocalFile
#End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.exportAccessObjectCli", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Public Function GetCodeLibInfoFromFilePath(ByVal filePath As String, Optional ByVal FindDependency As Boolean = True) As CodeLibInfo
   
   Dim cli As CodeLibInfo
On Error GoTo HandleErr

   getCodeLibInfoFromFile cli, fso.GetFile(filePath), FindDependency
   GetCodeLibInfoFromFilePath = cli

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.GetCodeLibInfoFromFilePath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function


Private Sub getCodeLibInfoFromFile(ByRef cli As CodeLibInfo, ByVal InputFile As Object, _
                          Optional ByVal FindDependency As Boolean = True, _
                          Optional ByVal FindTests As Boolean = True)
   
   Dim pos As Long
   Dim checkString As String
   Dim tempString As String
   Dim i As Long
   Dim fileNumber As Long
   
On Error GoTo HandleErr

   fileNumber = FreeFile
   Open InputFile.Path For Binary Access Read As fileNumber
   checkString = String$(LOF(fileNumber), 0)
   Get fileNumber, , checkString
   Close fileNumber
   
   'Namen ermitteln
   cli.Name = findSubString(checkString, SEARCHSTRING_ATTRIBUTNAME_BEGIN, SEARCHSTRING_ATTRIBUTNAME_END, pos)
   
   If Len(cli.Name) = 0 Then ' aus Dateinamen
      cli.Name = Left$(InputFile.Name, InStrRev(InputFile.Name, ".") - 1)
   End If
   
   'Typ bestimmen
   If pos > 0 Then ' da SEARCHSTRING_ATTRIBUTNAME_BEGIN gefunden wurde,
                   ' ist das definitiv ein Module oder eine Klasse
      cli.Type = clet_StdModule
      If InStr(1, Left$(checkString, pos - 1), "CLASS", vbBinaryCompare) Then
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

      pos = 1
      For i = 1 To 3
         pos = InStr(pos, checkString, vbNewLine, vbBinaryCompare) + 1
         If pos <= 0 Then Exit For
      Next
      If pos > 1 Then
         tempString = findSubString(Mid$(checkString, pos - 1), vbNewLine, vbNewLine)
      End If
      
      If tempString = SEARCHSTRING_FORMIDENTIFER Then
         cli.Type = clet_Form
      ElseIf tempString = SEARCHSTRING_REPORTIDENTIFER Then
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
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoFromFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub getCodeLibInfoRepositoryFile(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   
On Error GoTo HandleErr

   cli.RepositoryFile = Replace(findSubString(sSourceString, SEARCHSTRING_FILE_BEGIN, SEARCHSTRING_FILE_END), "\", "/")

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoRepositoryFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub getCodeLibInfoRepositoryFileReplacement(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   
On Error GoTo HandleErr

   cli.RepositoryFileReplacement = Replace(findSubString(sSourceString, SEARCHSTRING_REPLACE_BEGIN, SEARCHSTRING_REPLACE_END), "\", "/")

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoRepositoryFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub getCodeLibInfoLicenseFile(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
   
On Error GoTo HandleErr

   cli.LicenseFile = Replace(findSubString(sSourceString, SEARCHSTRING_LICENSE_BEGIN, SEARCHSTRING_LICENSE_END), "\", "/")

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoRepositoryFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub getCodeLibInfoDependency(ByRef cli As CodeLibInfo, ByRef sSourceString As String)
'Private Const SEARCHSTRING_USE_BEGIN As String = "<use>"
'Private Const SEARCHSTRING_USE_END   As String = "</use>"
   
   Dim dep() As String

On Error GoTo HandleErr

   If getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_USE_BEGIN, SEARCHSTRING_USE_END, dep) > 0 Then
      cli.Dependency = dep
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoDependency", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub getCodeLibInfoTests(ByRef cli As CodeLibInfo, ByRef sSourceString As String)

   Dim tests() As String

On Error GoTo HandleErr

   If getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_TEST_BEGIN, SEARCHSTRING_TEST_END, tests) > 0 Then
      cli.TestFiles = tests
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoTests", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function getCodelibInfoTagStrings(ByRef sSourceString As String, _
                                    ByRef sStartTag As String, ByRef sEndTag As String, _
                                    ByRef TagStrings() As String) As Long
   
   Dim startPos As Long
   Dim endPos As Long
   
   Dim tagString As String
   Dim tagIdx As Long
   
On Error GoTo HandleErr

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

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "getCodelibInfoTagStrings", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub getCodeLibInfoReferences(ByRef cli As CodeLibInfo, ByRef sSourceString As String)

   Dim refString As String
   
   Dim refBlocks() As String
   Dim refBlockMaxIdx As Long
   Dim ref() As CodeLibInfoReference
   Dim i As Long
   
On Error GoTo HandleErr

   refBlockMaxIdx = -1 + getCodelibInfoTagStrings(sSourceString, SEARCHSTRING_REF_BEGIN, SEARCHSTRING_REF_END, refBlocks)
   
   If refBlockMaxIdx >= 0 Then
      ReDim ref(refBlockMaxIdx)
      For i = 0 To refBlockMaxIdx
         refString = refBlocks(i)
         ref(i).Name = findSubString(refString, SEARCHSTRING_REF_NAME_BEGIN, SEARCHSTRING_REF_NAME_END)
         ref(i).Major = findSubString(refString, SEARCHSTRING_REF_MAJOR_BEGIN, SEARCHSTRING_REF_MAJOR_END)
         ref(i).Minor = findSubString(refString, SEARCHSTRING_REF_MINOR_BEGIN, SEARCHSTRING_REF_MINOR_END)
         ref(i).GUID = findSubString(refString, SEARCHSTRING_REF_GUID_BEGIN, SEARCHSTRING_REF_GUID_END)
      Next
      
      cli.References = ref
      
   End If
   
ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getCodeLibInfoReferences", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


Private Function findSubString(ByRef SourceString As String, _
                               ByRef StartTag As String, ByRef EndTag As String, _
                               Optional ByRef LocatedStartPos As Long, Optional ByRef LocatedEndPos As Long) As String

   Dim startPos As Long, endPos As Long
   
On Error GoTo HandleErr

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
   

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.findSubString", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

Private Sub createMissingFolder(ByRef FullPath As String)
   
   Dim TempPath As String
   
On Error GoTo HandleErr

   If Len(Dir$(FullPath, vbDirectory)) = 0 Then
      TempPath = getParentFolderPath(FullPath)
      createMissingFolder TempPath
      MkDir FullPath
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.addMissingFolder", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Function getParentFolderPath(ByRef FullPath As String) As String

On Error GoTo HandleErr

   getParentFolderPath = Left$(FullPath, InStrRev(FullPath, "\", , vbTextCompare) - 1)

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.getParentFolderPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr

   refreshAllAccessObjects AcObjectType.acDefault, ImportMode, ShowMessageAfterRefresh

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "RefreshAll", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr

   refreshAllAccessObjects AcObjectType.acModule, ImportMode, ShowMessageAfterRefresh

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "RefreshAllModules", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub refreshAllAccessObjects(ByVal aoType As AcObjectType, _
                                    ByVal ImportMode As CodeLibImportMode, _
                                    ByVal ShowMessageAfterRefresh As Boolean)
   
'/**
' * @todo Regelung finden, falls FileCollection noch Einträge enthält
'**/

On Error GoTo HandleErr

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

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "refreshAllAccessObjects", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub addAccessObjectCollectionToRefreshFileCollection( _
                        ByRef aoType As AcObjectType, _
                        ByRef ImportMode As CodeLibImportMode)

   Dim aoCol As Object ' AllModules, ...
   Dim ao As AccessObject

On Error GoTo HandleErr

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

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "ACLibFileManager.ExportAllModules", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Sub addAccessObjectToRefreshFileCollection(ByRef ao As AccessObject, ByRef ImportMode As CodeLibImportMode)
   
   Dim cli As CodeLibInfo
   Dim tempFile As Object
   
On Error GoTo HandleErr

   getCodeLibInfoFromAccessObject cli, ao.Type, ao.Name
   If Len(cli.RepositoryFile) > 0 Then
      cli.LocalFile = getRepositoryFullPath(cli.RepositoryFile)
      Set tempFile = fso.GetFile(cli.LocalFile)
      addMissingFile tempFile, ImportMode
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "addAccessObjectToRefreshFileCollection", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

#If EARLYBINDING = 1 Then
Private Function getFirstProcLine(cm As VBIDE.CodeModule) As Long
#Else
Private Function getFirstProcLine(cm As Object) As Long
#End If

   Dim lngLine As Long
   Dim procKind As vbext_ProcKind
   Dim procName As String

On Error GoTo HandleErr

   For lngLine = cm.CountOfDeclarationLines + 1 To cm.countOfLines
      procName = cm.ProcOfLine(lngLine, procKind)
      If Len(procName) > 0 Then
         getFirstProcLine = cm.ProcBodyLine(procName, procKind)
         Exit For
      End If
   Next

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "getFirstProcLine", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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

On Error GoTo HandleErr

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
   
   strCodeLibInfo = strCodeLibInfo & _
      "'</codelib>" & vbNewLine
   
   Debug.Print strCodeLibInfo

ExitHere:
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "InsertCodeLibInfo", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub
