Attribute VB_Name = "FileTools"
Attribute VB_Description = "Funktionen für Dateioperationen"
'---------------------------------------------------------------------------------------
' Module: FileTools
'---------------------------------------------------------------------------------------
'/**
'\author    Josef Poetzl
'\short     File operation functions
' <remarks>
' </remarks>
'\ingroup file
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/FileTools.bas</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/file/FileToolsTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

#If USELOCALIZATION_DE = 1 Then
Private Const SELECTBOX_FILE_DIALOG_TITLE As String = "Datei auswählen"
Private Const SELECTBOX_FOLDER_DIALOG_TITLE As String = "Ordner auswählen"
Private Const SELECTBOX_OPENTITLE As String = "auswählen"
Private Const FILTERSTRING_ALL_FILES As String = "Alle Dateien (*.*)"
#Else
Private Const SELECTBOX_FILE_DIALOG_TITLE As String = "Select file"
Private Const SELECTBOX_FOLDER_DIALOG_TITLE As String = "Select folder"
Private Const SELECTBOX_OPENTITLE As String = "auswählen"
Private Const FILTERSTRING_ALL_FILES As String = "All Files (*.*)"
#End If

Private Const DEFAULT_TEMPPATH_NOENV As String = "C:\"
Private Const PATHLEN_MAX As Long = 255

Private Const SE_ERR_NOTFOUND As Long = 2
Private Const SE_ERR_NOASSOC  As Long = 31

Private Const VbaErrNo_FileNotFound As Long = 53

#If VBA7 Then

Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare PtrSafe Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long
         
Private Declare PtrSafe Function API_ShellExecuteA Lib "shell32.dll" ( _
         ByVal Hwnd As LongPtr, _
         ByVal lOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

#Else

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

Private Declare Function API_ShellExecuteA Lib "shell32.dll" ( _
         ByVal Hwnd As Long, _
         ByVal lOperation As String, _
         ByVal lpFile As String, _
         ByVal lpParameters As String, _
         ByVal lpDirectory As String, _
         ByVal nShowCmd As Long) As Long

#End If

'---------------------------------------------------------------------------------------
' Function: SelectFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Select file using dialogue
' </summary>
' <param name="InitDir">Initial Folder</param>
' <param name="DlgTitle">Title of dialogue</param>
' <param name="FilterString">Filter settings - Example: "(*.*)" oder "All (*.*)|text files (*.txt)|Images (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Multi-selection</param>
' <param name="ViewMode">View mode (0: Detail view, 1: Preview, 2: Properties, 3: List, 4: Thumbnail, 5: Large symbols, 6: Small symbols)</param>
' <returns>String (in case of multiple selection, the files are separated by chr(9))</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFile(Optional ByVal InitialDir As String = vbNullString, _
                           Optional ByVal DlgTitle As String = SELECTBOX_FILE_DIALOG_TITLE, _
                           Optional ByVal FilterString As String = FILTERSTRING_ALL_FILES, _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal ViewMode As Long = -1) As String

    SelectFile = WizHook_GetFileName(InitialDir, DlgTitle, SELECTBOX_OPENTITLE, FilterString, MultiSelectEnabled, , ViewMode, False)

End Function

'---------------------------------------------------------------------------------------
' Function: SelectFolder
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Folder selection dialogue
' </summary>
' <param name="InitDir">Initial Folder</param>
' <param name="DlgTitle">Title of dialogue</param>
' <param name="FilterString">Filter settings, Default:*</param>
' <param name="MultiSelect">Multi-selection</param>
' <param name="ViewMode">View mode (0: Detail view, 1: Preview, 2: Properties, 3: List, 4: Thumbnail, 5: Large symbols, 6: Small symbols)</param>
' <returns>String (in case of multiple selection, folders are separated by chr(9))</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFolder(Optional ByVal InitialDir As String = vbNullString, _
                             Optional ByVal DlgTitle As String = SELECTBOX_FOLDER_DIALOG_TITLE, _
                             Optional ByVal FilterString As String = "*", _
                             Optional ByVal MultiSelectEnabled As Boolean = False, _
                             Optional ByVal ViewMode As Long = -1) As String

   SelectFolder = WizHook_GetFileName(InitialDir, DlgTitle, SELECTBOX_OPENTITLE, FilterString, MultiSelectEnabled, , ViewMode, True)

End Function

Private Function WizHook_GetFileName( _
                           ByVal InitialDir As String, _
                           ByVal DlgTitle As String, _
                           ByVal OpenTitle As String, _
                           ByVal FilterString As String, _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal SplitDelimiter As String = "|", _
                           Optional ByVal ViewMode As Long = -1, _
                           Optional ByVal SelectFolderFlag As Boolean = False, _
                           Optional ByVal AppName As String) As String

'Summary of WizHook.GetFileName parameters: http://www.team-moeller.de/?Tipps_und_Tricks:Wizhook-Objekt:GetFileName
'View  0: Detailansicht
'      1: Vorschauansicht
'      2: Eigenschaften
'      3: Liste
'      4: Miniaturansicht
'      5: Große Symbole
'      6: Kleine Symbole

'flags 4: Set Current Dir
'      8: Mehrfachauswahl möglich
'     32: Ordnerauswahldialog
'     64: Wert im Parameter "View" berücksichtigen

   Dim SelectedFileString As String
   Dim WizHookRetVal As Long

   If InStr(1, InitialDir, " ") > 0 Then
      InitialDir = """" & InitialDir & """"
   End If

   Dim Flags As Long
   Flags = 0
   If MultiSelectEnabled Then Flags = Flags + 8
   If SelectFolderFlag Then Flags = Flags + 32

   If ViewMode >= 0 Then
      Flags = Flags + 64
   Else
      ViewMode = 0
   End If

   WizHook.Key = 51488399
   WizHookRetVal = WizHook.GetFileName( _
                        Access.Application.hWndAccessApp, AppName, DlgTitle, OpenTitle, _
                        SelectedFileString, InitialDir, FilterString, 0, ViewMode, Flags, True)
   If WizHookRetVal = 0 Then
      If MultiSelectEnabled Then SelectedFileString = Replace(SelectedFileString, vbTab, SplitDelimiter)
      WizHook_GetFileName = SelectedFileString
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: UNCPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Returns the UNC path
' </summary>
' <param name="Path">Path to convert</param>
' <param name="IgnoreErrors">true = ignore API errors</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function UncPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String
   
   Dim UNC As String * 512
   
   If VBA.Len(Path) = 1 Then Path = Path & ":"
   
   If WNetGetConnection(VBA.Left$(Path, 2), UNC, VBA.Len(UNC)) Then
   
      If IgnoreErrors Then
         UncPath = Path
      Else
         Err.Raise 5 ' Invalid procedure call or argument
      End If
   
   Else
   
      ' Ergebnis zurückgeben:
      UncPath = VBA.Left$(UNC, VBA.InStr(UNC, vbNullChar) - 1) & VBA.Mid$(Path, 3)
   
   End If
   
End Function

'---------------------------------------------------------------------------------------
' Property: TempPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Determine Temp folder
' </summary>
' <returns>String</returns>
' <remarks>
' Uses API GetTempPathA
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TempPath() As String

   Dim TempString As String

   TempString = Space$(PATHLEN_MAX)
   API_GetTempPath PATHLEN_MAX, TempString
   TempString = Left$(TempString, InStr(TempString, Chr$(0)) - 1)
   If Len(TempString) = 0 Then
      TempString = DEFAULT_TEMPPATH_NOENV
   End If
   TempPath = TempString

End Property

'---------------------------------------------------------------------------------------
' Function: TempPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Generate temp. file name
' </summary>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetNewTempFileName(Optional ByVal PathToUse As String = "", _
                         Optional ByVal FilePrefix As String = "", _
                         Optional ByVal FileExtension As String = "") As String

   Dim NewTempFileName As String
   
   If Len(PathToUse) = 0 Then
      PathToUse = TempPath
   End If

   NewTempFileName = String$(PATHLEN_MAX, 0)
   Call API_GetTempFilename(PathToUse, FilePrefix, 0&, NewTempFileName)

   NewTempFileName = Left$(NewTempFileName, InStr(NewTempFileName, Chr$(0)) - 1)

   'Delete file, as only name is needed
   Call Kill(NewTempFileName)

   If Len(FileExtension) > 0 Then 'Fileextension umschreiben
     NewTempFileName = Left$(NewTempFileName, Len(NewTempFileName) - 3) & FileExtension
   End If

   GetNewTempFileName = NewTempFileName

End Function

'---------------------------------------------------------------------------------------
' Function: ShortenFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Shorten file path to n characters
' </summary>
' <param name="FullFileName">Full path</param>
' <param name="MaxLen">required length</param>
' <returns>String</returns>
' <remarks>
' Helpful for the displays in narrow textboxes \n
' Example: <source>C:\Programms\...\Folder\File.txt</source>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ShortenFileName(ByVal FullFileName As Variant, ByVal MaxLen As Long) As String

   Dim FileString As String
   Dim Temp As String
   Dim TrimPos As Long

   FileString = Nz(FullFileName, vbNullString)
   If Len(FileString) <= MaxLen Then
      ShortenFileName = FileString
      Exit Function
   End If

   TrimPos = InStrRev(FileString, "\")
   Temp = Mid$(FileString, TrimPos)
   FileString = Left$(FileString, TrimPos - 1)

   TrimPos = MaxLen - Len(Temp) - 3
   If TrimPos < 2 Then
      FileString = "..." & Temp
   Else
      TrimPos = TrimPos \ 2
      FileString = Left$(FileString, TrimPos) & "..." & Right$(FileString, TrimPos) & Temp
   End If

   ShortenFileName = FileString

End Function

'---------------------------------------------------------------------------------------
' Function: FileNameWithoutPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Extract file name from complete path specification
' </summary>
' <param name="FullPath">File name incl. directory</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileNameWithoutPath(ByVal FullPath As Variant) As String

   Dim Temp As String
   Dim Pos As Long

   Temp = Nz(FullPath, vbNullString)
   Pos = InStrRev(Temp, "\")
   If Pos > 0 Then
      FileNameWithoutPath = Mid$(Temp, Pos + 1)
   Else
      FileNameWithoutPath = Temp
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: GetDirFromFullFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Determines the directory from the complete path of a file.
' </summary>
' <param name="FullFileName">complete file path</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDirFromFullFileName(ByVal FullFileName As String) As String
   GetDirFromFullFileName = PathFromFullFileName(FullFileName)
End Function

Public Function PathFromFullFileName(ByVal FullFileName As Variant) As String

   Dim DirPath As String
   Dim Pos As Long

   DirPath = FullFileName
   Pos = InStrRev(DirPath, "\")
   If Pos > 0 Then
      DirPath = Left$(DirPath, Pos)
   Else
      DirPath = vbNullString
   End If

   PathFromFullFileName = DirPath

End Function

'---------------------------------------------------------------------------------------
' Function: CreateDirectory
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates a directory including all missing parent directories
' </summary>
' <param name="FullPath">Directory to be created</param>
' <returns>Boolean: True = directory/folder created</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateDirectory(ByVal FullPath As String) As Boolean

   Dim PathBefore As String

   If Right$(FullPath, 1) = "\" Then
      FullPath = VBA.Left$(FullPath, Len(FullPath) - 1)
   End If

   If DirExists(FullPath) Then
      CreateDirectory = False
      Exit Function
   End If

   PathBefore = VBA.Mid$(FullPath, 1, VBA.InStrRev(FullPath, "\") - 1)
   If Not DirExists(PathBefore) Then
      If CreateDirectory(PathBefore) = False Then
         CreateDirectory = False
         Exit Function
      End If
   End If

   VBA.MkDir FullPath

   CreateDirectory = True

End Function

Public Sub CreateDirectoryIfMissing(ByVal FullPath As String)
   CreateDirectory FullPath
End Sub

'---------------------------------------------------------------------------------------
' Function: FileExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Check: file exists
' </summary>
' <param name="FullPath">Full path specification</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileExists(ByVal FullPath As String) As Boolean

   Do While VBA.Right$(FullPath, 1) = "\"
      FullPath = VBA.Left$(FullPath, Len(FullPath) - 1)
   Loop

   FileExists = (VBA.Len(VBA.Dir$(FullPath, vbReadOnly Or vbHidden Or vbSystem)) > 0) And (VBA.Len(FullPath) > 0)
   VBA.Dir$ "\" ' Avoiding error: issue #109

End Function

'---------------------------------------------------------------------------------------
' Function: DirExists
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Check: directory/folder exists
' </summary>
' <param name="FullPath">Full path specification</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DirExists(ByVal FullPath As String) As Boolean

   If VBA.Right$(FullPath, 1) <> "\" Then
      FullPath = FullPath & "\"
   End If

   DirExists = (VBA.Dir$(FullPath, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem) = ".")
   VBA.Dir$ "\" ' Avoiding error: issue #109
   
End Function

'---------------------------------------------------------------------------------------
' Function: GetFileUpdateDate
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Last modified date of a file
' </summary>
' <param name="FullFileName">Full path specification</param>
' <returns>Variant</returns>
' <remarks>
' Errors from API function are ignored
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileUpdateDate(ByVal FullFileName As String) As Variant
   If FileExists(FullFileName) Then
      On Error Resume Next
      GetFileUpdateDate = FileDateTime(FullFileName)
   Else
      GetFileUpdateDate = Null
   End If
End Function

'---------------------------------------------------------------------------------------
' Function: ConvertStringToFileName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates a file name from a string (replaces special characters)
' </summary>
' <param name="Text">Initial string for file names</param>
' <param name="ReplaceWith">Characters as a substitute for special characters</param>
' <param name="CharsToReplace">Characters that are replaced with ReplaceWith</param>
' <param name="CharsToDelete">Characters that will be removed</param>
' <returns>String</returns>
' <remarks>
' special characters: ? * " / ' : ( )
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ConvertStringToFileName(ByVal Text As String, _
                                   Optional ByVal ReplaceWith As String = "_", _
                                   Optional ByVal CharsToReplace As String = "/':()", _
                                   Optional ByVal CharsToDelete As String = "?*""") As String

   Dim FileName As String
   Dim i As Long

   FileName = Trim$(Text)

   For i = 1 To Len(CharsToDelete)
      FileName = Replace(FileName, Mid(CharsToReplace, i, 1), vbNullString)
   Next

   For i = 1 To Len(CharsToReplace)
      FileName = Replace(FileName, Mid(CharsToReplace, i, 1), ReplaceWith)
   Next

   ConvertStringToFileName = FileName

End Function

'---------------------------------------------------------------------------------------
' Function: GetFullPathFromRelativPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates a complete path specification from relative path specification and "base directory".
' </summary>
' <param name="RelativPath">relative path</param>
' <param name="BaseDir">Base directory</param>
' <returns>String</returns>
' <remarks>
' Example:
' GetFullPathFromRelativPath("..\..\Test.txt", "C:\Programms\xxx\") => "C:\test.txt"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFullPathFromRelativPath(ByVal RelativPath As String, _
                                           ByVal BaseDir As String) As String

   Dim FullPath As String
   Dim Pos As Long

   If Right$(BaseDir, 1) = "\" Then
      BaseDir = Left$(BaseDir, Len(BaseDir) - 1)
   End If

   FullPath = RelativPath
   If Mid$(FullPath, 2, 1) = ":" Or Left$(FullPath, 2) = "\\" Then ' absolut path !!!
      GetFullPathFromRelativPath = FullPath
      Exit Function
   ElseIf Left$(FullPath, 1) = "\" Then 'first dir
      Pos = InStr(3, BaseDir, "\")
      If Pos > 0 Then
         BaseDir = Left$(BaseDir, Pos - 1)
      End If
      GetFullPathFromRelativPath = BaseDir & FullPath
      Exit Function
   ElseIf FullPath = "." Then
      GetFullPathFromRelativPath = BaseDir
      Exit Function
   ElseIf Left$(FullPath, 2) = ".\" Then
      FullPath = Mid$(FullPath, 3)
   End If

   Do While Left$(FullPath, 3) = "..\"
      FullPath = Mid$(FullPath, 4)
      Pos = InStrRev(BaseDir, "\")
      If Pos > 0 Then
         BaseDir = Left$(BaseDir, Pos - 1)
      End If
   Loop

   GetFullPathFromRelativPath = BaseDir & "\" & FullPath

End Function

'---------------------------------------------------------------------------------------
' Function: GetRelativPathFromFullPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates a relative path from the complete path specification and source directory
' </summary>
' <param name="FullPath">Full path specification</param>
' <param name="BaseDir">Base directory</param>
' <param name="RelativePrefix">Add ".\" as relative path identifier</param>
' <returns>String</returns>
' <remarks>
' Example:
' <code>
' GetRelativPathFromFullPath("C:\test.txt", "C:\Programms\xxx\", True)
' => ".\..\..\test.txt"
' </code>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetRelativPathFromFullPath(ByVal FullPath As String, _
                                           ByVal BaseDir As String, _
                                  Optional ByVal EnableRelativePrefix As Boolean = False, _
                                  Optional ByVal DisableDecreaseBaseDir As Boolean = False) As String

   Dim RelativPath As String
   
   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   If Right$(BaseDir, 1) <> "\" Then BaseDir = BaseDir & "\"
   If FullPath = BaseDir Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If
   
   If Not DisableDecreaseBaseDir Then
      RelativPath = TryGetRelativPathWithDecreaseBaseDir(FullPath, BaseDir, EnableRelativePrefix)
   Else
      RelativPath = FullPath
      If Right$(BaseDir, 1) <> "\" Then BaseDir = BaseDir & "\"
      If Len(BaseDir) > 0 Then
         If Nz(InStr(1, FullPath, BaseDir, vbTextCompare), 0) > 0 Then
            RelativPath = Mid$(FullPath, Len(BaseDir) + 1)
            If EnableRelativePrefix Then
               RelativPath = ".\" & RelativPath
            End If
         End If
      End If
   End If
   
   GetRelativPathFromFullPath = RelativPath

End Function

Private Function TryGetRelativPathWithDecreaseBaseDir(ByVal FullPath As String, ByVal BaseDir As String, ByVal EnableRelativePrefix As Boolean) As String

   Dim RelativPath As String
   Dim DecreaseCounter As Long
   Dim Pos As Long
   Dim i As Long
   
   RelativPath = BaseDir

   Do While InStr(1, FullPath, RelativPath) = 0
      Pos = InStrRev(Left$(RelativPath, Len(RelativPath) - 1), "\")
      RelativPath = Left$(RelativPath, Pos)
      DecreaseCounter = DecreaseCounter + 1
      If Len(RelativPath) = 0 Then
         DecreaseCounter = 0
         Exit Do
      End If
   Loop
   
   If Len(RelativPath) > 0 Then
      RelativPath = Replace(FullPath, RelativPath, vbNullString)
      For i = 1 To DecreaseCounter
         RelativPath = "..\" & RelativPath
      Next

      If EnableRelativePrefix Then
         RelativPath = ".\" & RelativPath
      End If
   Else
      RelativPath = FullPath
   End If

   TryGetRelativPathWithDecreaseBaseDir = RelativPath

End Function

'---------------------------------------------------------------------------------------
' Sub: AddToZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Add file to Zip file
' </summary>
' <param name="ZipFile">Zip file</param>
' <param name="FullFileName">file to append</param>
' <remarks>
' CreateObject("Shell.Application").Namespace(zipFile & "").CopyHere sFile & ""
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddToZipFile(ByVal ZipFile As String, ByVal FullFileName As String)

   If Not FileExists(ZipFile) Then
      CreateZipFile ZipFile
   End If

   With CreateObject("Shell.Application")
      .NameSpace(ZipFile & "").CopyHere FullFileName & ""
   End With

End Sub

'---------------------------------------------------------------------------------------
' Function: ExtractFromZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Extract file from zip file
' </summary>
' <param name="ZipFile">Zip file</param>
' <param name="Destination">Destination folder</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ExtractFromZipFile(ByVal ZipFile As String, ByVal Destination As String) As String

   With CreateObject("Shell.Application")
      .NameSpace(Destination & "").CopyHere .NameSpace(ZipFile & "").Items
      ExtractFromZipFile = .NameSpace(ZipFile & "").Items.Item(0).Name
   End With

End Function

'---------------------------------------------------------------------------------------
' Function: CreateZipFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Creates an empty zip file
' </summary>
' <param name="ZipFile">Zip file (full path)</param>
' <param name="DeleteExistingFile">Delete existing Zip file</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateZipFile(ByVal ZipFile As String, Optional ByRef DeleteExistingFile As Boolean = False) As Boolean

   Dim FileHandle As Long

   If FileExists(ZipFile) Then
      If DeleteExistingFile Then
         Kill ZipFile
      Else
         CreateZipFile = False
         Exit Function
      End If
   End If

   FileHandle = FreeFile
   Open ZipFile For Output As #FileHandle
   Print #FileHandle, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, 0)
   Close #FileHandle

   CreateZipFile = FileExists(ZipFile)

End Function

'---------------------------------------------------------------------------------------
' Function: GetFileExtension
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Returns the file extension of a file returns.
' </summary>
' <param name="FilePath">File path or file name</param>
' <param name="WithDotBeforeExtension">True: returns extension excl. separator</param>
' <returns>File extension (String)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileExtension(ByVal FilePath As String, Optional ByVal WithDotBeforeExtension As Boolean = False) As String
   GetFileExtension = VBA.Strings.Mid$(FilePath, VBA.Strings.InStrRev(FilePath, ".") + (1 - Abs(WithDotBeforeExtension)))
End Function


'---------------------------------------------------------------------------------------
' Function: OpenFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Open file with API ShellExecute
' </summary>
' <param name="FileName">File path or file name</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenFile(ByVal FilePath As String, Optional ByVal ReadOnlyMode As Boolean = False) As Boolean

   Const FileNotFoundErrorTextTemplate As String = "File '{FilePath}' not found."
   Dim FileNotFoundErrorText As String

   If Len(VBA.Dir(FilePath)) = 0 Then
   
#If USELOCALIZATION = 1 Then
      FileNotFoundErrorText = Replace(L10n.Text(FileNotFoundErrorTextTemplate), "{FilePath}", FilePath)
#Else
      FileNotFoundErrorText = Replace(FileNotFoundErrorTextTemplate, "{FilePath}", FilePath)
#End If
      Err.Raise VbaErrNo_FileNotFound, "FileTools.OpenFile", FileNotFoundErrorText
      Exit Function
   End If

   OpenFile = ShellExecute(FilePath, "open")
   
End Function

'---------------------------------------------------------------------------------------
' Function: OpenFilePath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Open folder with API ShellExecute
' </summary>
' <param name="FilePath">folder path or file name</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenFilePath(ByVal FolderPath As String) As Boolean

   Const FolderNotFoundErrorTextTemplate As String = "File '{FolderPath}' not found."
   Dim FolderNotFoundErrorText As String

   If Len(VBA.Dir(FolderPath, vbDirectory)) = 0 Then
   
#If USELOCALIZATION = 1 Then
      FolderNotFoundErrorText = Replace(L10n.Text(FolderNotFoundErrorTextTemplate), "{FolderPath}", FolderPath)
#Else
      FolderNotFoundErrorText = Replace(FolderNotFoundErrorTextTemplate, "{FolderPath}", FolderPath)
#End If
      Err.Raise VbaErrNo_FileNotFound, "FileTools.OpenFilePath", FolderNotFoundErrorText
      Exit Function
   End If

   OpenFilePath = ShellExecute(FolderPath, "open")
   
End Function

Private Function ShellExecute(ByVal FilePath As String, _
                     Optional ByVal ApiOperation As String = vbNullString) As Boolean

   Const FileNotFoundErrorTextTemplate As String = "File '{FilePath}' not found."
   Dim FileNotFoundErrorText As String
   Dim Ret As Long
   Dim Directory As String
   Dim DeskWin As Long

   If Len(FilePath) = 0 Then
      ShellExecute = False
      Exit Function
   Else
      DeskWin = Application.hWndAccessApp
      Ret = API_ShellExecuteA(DeskWin, ApiOperation, FilePath, vbNullString, vbNullString, vbNormalFocus)
   End If
   
   If Ret = SE_ERR_NOTFOUND Then
#If USELOCALIZATION = 1 Then
      FileNotFoundErrorText = Replace(L10n.Text(FileNotFoundErrorTextTemplate), "{FilePath}", FilePath)
#Else
      FileNotFoundErrorText = Replace(FileNotFoundErrorTextTemplate, "{FilePath}", FilePath)
#End If
      Err.Raise VbaErrNo_FileNotFound, "FileTools.OpenFile", FileNotFoundErrorText
      ShellExecute = False
      Exit Function
   ElseIf Ret = SE_ERR_NOASSOC Then
      ShellExecute = False
      Exit Function
' ToDo: "Öffnen mit"-Dialog verwenden:
      'Wenn die Dateierweiterung noch nicht bekannt ist...
      'wird der "Öffnen mit..."-Dialog angezeigt.
'      Directory = Space$(260)
'      Ret = GetSystemDirectory(Directory, Len(Directory))
'      Directory = Left$(Directory, Ret)
'      Call ShellExecuteA(DeskWin, vbNullString, "RUNDLL32.EXE", "shell32.dll, OpenAs_RunDLL " & _
'         FilePath, Directory, vbNormalFocus)
   End If
   
   ShellExecute = True

End Function
