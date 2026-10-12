# File and folder dialogs, paths, temp files

Keywords: file dialog, select file, select folder, FileDialog, path, relative path, UNC path, temp file, file exists,
create folder, zip, Dateiauswahl, Ordnerauswahl, Pfad, relativer Pfad, temporäre Datei, Datei vorhanden, Ordner erstellen

## Import

- `file/FileTools.bas`

## Usage

```vba
' File dialog with filter (several files: separated by Chr(9))
Dim FilePath As String
FilePath = FileTools.SelectFile(CurrentProject.Path, "Select import file", "Excel (*.xlsx)|CSV (*.csv)")
If Len(FilePath) = 0 Then Exit Sub        ' cancelled

' Folder dialog
Dim FolderPath As String
FolderPath = FileTools.SelectFolder(CurrentProject.Path)

' Checks and folders
If Not FileTools.FileExists(FilePath) Then Err.Raise 53
FileTools.CreateDirectoryIfMissing CurrentProject.Path & "\export\2024"
Debug.Print FileTools.DirExists("C:\Temp")

' Path parts
Debug.Print FileTools.FileNameWithoutPath("C:\Windows\abc.txt")       ' abc.txt
Debug.Print FileTools.GetFileExtension("C:\Windows\abc.txt")          ' txt
Debug.Print FileTools.GetDirFromFullFileName("C:\Windows\abc.txt")

' Relative paths (e.g. to store paths relative to the application folder)
Debug.Print FileTools.GetRelativPathFromFullPath("C:\Windows\A1\A2\A3", "C:\Windows", True)   ' .\A1\A2\A3
Debug.Print FileTools.GetFullPathFromRelativPath(".\..\..\..\B1", "C:\Windows\A1\A2\A3\")     ' C:\Windows\B1

' Temp files, UNC path, safe file name
Debug.Print FileTools.TempPath
Debug.Print FileTools.GetNewTempFileName(, "exp", "csv")
Debug.Print FileTools.UncPath("X:\Data\file.accdb")                    ' \\server\share\Data\file.accdb
Debug.Print FileTools.ConvertStringToFileName("Report 1/2: (draft)")

' Open a file or folder with the associated application
FileTools.OpenFile FilePath
FileTools.OpenFilePath FolderPath

' Zip
FileTools.AddToZipFile "C:\Temp\export.zip", FilePath
```

## Notes

- `FileTools` is a standard module with `Option Private Module`: call the functions with or without the prefix
  `FileTools.`.
- The dialogs use the Access-internal `WizHook.GetFileName` (no reference to the Office library required).
- Zip functions use the Windows Shell (`Shell.Application`).

## See also

- [open-file-or-url.md](open-file-or-url.md) - open files and URLs, start programs (WinApiShellTools)
- `file/RecursiveFileSearcher.cls` - search folder trees for files
- tests `_test/file/FileToolsTests.cls`
