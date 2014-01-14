VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LibFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class Module: LibFiles
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' \brief        Dateien in Tabelle abspeichern und in Verzeichnis exportieren
' \ingroup file
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/LibFiles.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
' !!! Auch VBA als Bibliotheks-Präfix angeben,
'     damit dieses Modul auch in einer mdb ausgeführt werden kann,
'     dessen Verweis noch nicht passt
'
'---------------------------------------------------------------------------------------
Option Compare Text
Option Explicit

'---------------------------------------------------------------------------------------------------
' File Table
'---------------------------------------------------------------------------------------------------
Private Const m_TABLE_AppFiles As String = "usys_AppFiles"
Private Const m_TABLE_FIELD_FileName As String = "FileName"
Private Const m_TABLE_FIELD_EarlyBindingRefName As String = "EarlyBindingRefName"
Private Const m_TABLE_FIELD_DestinationFolder As String = "DestinationFolder"
Private Const m_TABLE_FIELD_Version As String = "Version"
Private Const m_TABLE_FIELD_FileData As String = "FileData"

Private Const m_TABLE_AllFields_SqlString As String = _
            "select " & m_TABLE_FIELD_FileName & ", " & _
                        m_TABLE_FIELD_EarlyBindingRefName & ", " & _
                        m_TABLE_FIELD_DestinationFolder & ", " & _
                        m_TABLE_FIELD_Version & ", " & _
                        m_TABLE_FIELD_FileData & _
            " from " & m_TABLE_AppFiles

Private Const m_TABLE_InfoFields_SqlString As String = _
            "select " & m_TABLE_FIELD_FileName & ", " & _
                        m_TABLE_FIELD_EarlyBindingRefName & ", " & _
                        m_TABLE_FIELD_DestinationFolder & ", " & _
                        m_TABLE_FIELD_Version & _
            " from " & m_TABLE_AppFiles

Private Const m_TABLE_DATA_CurrentProjectPathCode = "[APPDIR]"

Private m_LibPath As String
Private m_ReferencesUpdateAllowed As Boolean

'---------------------------------------------------------------------------------------------------
' API
'---------------------------------------------------------------------------------------------------

Private Const MAX_PATH As Long = 260&
Private Const CSIDL_FLAG_CREATE As Long = &H8000&
Private Const CSIDL_FLAG_DONT_VERIFY As Long = &H4000&
Private Const SHGFP_TYPE_CURRENT As Long = 0&

Public Enum CSIDL_FOLDER
   CSIDL_PERSONAL = &H5& ' Eigene Dateien
   CSIDL_STARTUP = &H7& ' Autostart
   CSIDL_APPDATA = &H1A& ' Anwendungsdaten
   CSIDL_LOCAL_APPDATA = &H1C& ' Lokale Anwendungsdaten
   CSIDL_COMMON_APPDATA = &H23& ' Anwendungsdaten (Allgemein)
   CSIDL_WINDOWS = &H24& ' Windows
   CSIDL_SYSTEM = &H25& ' Windows-System
   CSIDL_PROGRAM_FILES = &H26& ' Programme
   CSIDL_PROGRAM_FILESX86 = &H2A&
   CSIDL_PROGRAM_FILES_COMMON = &H2B& ' Gemeinsame Dateien
   CSIDL_PROGRAM_FILES_COMMONX86 = &H2C&
End Enum

#If VBA7 Then

   Private Declare PtrSafe Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
     ByVal lptstrFilename As String, _
     lpdwHandle As Long) As Long
   
   Private Declare PtrSafe Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
     ByVal lptstrFilename As String, _
     ByVal dwHandle As Long, _
     ByVal dwLen As Long, _
     ByRef lpData As Any _
     ) As Long
   
   Private Declare PtrSafe Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
     pBlock As Any, _
     ByVal lpSubBlock As String, _
     lplpBuffer As Any, _
     puLen As Long _
     ) As Long
   
   Private Declare PtrSafe Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
     Dest As Any, _
     ByVal Source As Long, _
     ByVal Length As Long)

   Private Declare PtrSafe Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

   
   Private Declare PtrSafe Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" ( _
         ByVal hwndOwner As Long, ByVal nFolder As Long, _
         ByVal hToken As Long, ByVal dwFlags As Long, _
         ByVal pszPath As String) As Long

#Else

   Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
     ByVal lptstrFilename As String, _
     lpdwHandle As Long) As Long
   
   Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
     ByVal lptstrFilename As String, _
     ByVal dwHandle As Long, _
     ByVal dwLen As Long, _
     ByRef lpData As Any _
     ) As Long
   
   Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
     pBlock As Any, _
     ByVal lpSubBlock As String, _
     lplpBuffer As Any, _
     puLen As Long _
     ) As Long
   
   Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
     Dest As Any, _
     ByVal Source As Long, _
     ByVal Length As Long)
   
   Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

   Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" ( _
         ByVal hwndOwner As Long, ByVal nFolder As Long, _
         ByVal hToken As Long, ByVal dwFlags As Long, _
         ByVal pszPath As String) As Long


#End If



Private Type VS_FIXEDFILEINFO
  dwSignature As Long
  dwStrucVersion As Long
  dwFileVersionMS As Long
  dwFileVersionLS As Long
  dwProductVersionMS As Long
  dwProductVersionLS As Long
  dwFileFlagsMask As Long
  dwFileFlags As Long
  dwFileOS As Long
  dwFileType As Long
  dwFileSubtype As Long
  dwFileDateMS As Long
  dwFileDateLS As Long
End Type

Private Type FILEINFOOUT
  FileVersion As String
  ProductVersion As String
End Type
'---------------------------------------------------------------------------------------------------

#If UseAssemblyDebugFolder Then

Private m_AssemblyDebugFolder As String
Public Event GetAssemblyDebugFolder(ByRef AssemblyDebugFolder As String)

#End If

Private Sub Class_Initialize()
   'LibPath = DefaultLibPath
   m_ReferencesUpdateAllowed = Not IsMDE
End Sub

'---------------------------------------------------------------------------------------
' Sub: ReInitialize
'---------------------------------------------------------------------------------------
'/**                                            '<-- Start Doxygen-Block
' <summary>
' (Re)Initialisiert die DotNetLib Integration
' </summary>
' <remarks>
' Muss bei jedem Anwendungsstart ausgeführt werden
' </remarks>
'**/                                            '<-- Ende Doxygen-Block
'---------------------------------------------------------------------------------------
Public Sub ReInitialize()
   LibFiles.CheckFiles
   LibFiles.SetLibPathAsCurrentDir
End Sub

Friend Property Get DefaultLibPath() As String
#If UseAssemblyDebugFolder Then
   If Len(AssemblyDebugFolder) > 0 Then
      DefaultLibPath = AssemblyDebugFolder
   Else
#Else
      DefaultLibPath = ApplicationLibPath & "\lib\"
#End If
#If UseAssemblyDebugFolder Then
   End If
#End If
End Property

Private Function ApplicationLibPath() As String
' MS.net-dlls laufen nicht in Netzwerkpfad

   Dim p As String
   p = CodeProject.Path
   If p <> UNCPath(p) Then
      p = GetSpecFolder(CSIDL_LOCAL_APPDATA) & "\" & ApplicationName
   End If
   ApplicationLibPath = p
   
End Function

Private Function ApplicationName() As String
   Dim s As String
   s = Access.CodeProject.Name
   s = VBA.Left$(s, VBA.InStrRev(s, ".") - 1)
   ApplicationName = s
End Function

#If UseAssemblyDebugFolder Then
   
Friend Property Get AssemblyDebugFolder() As String
   If Len(m_AssemblyDebugFolder) = 0 Then
      RaiseEvent GetAssemblyDebugFolder(m_AssemblyDebugFolder)
   End If
   AssemblyDebugFolder = m_AssemblyDebugFolder
End Property

Friend Property Let AssemblyDebugFolder(ByVal NewValue As String)
   m_AssemblyDebugFolder = NewValue
End Property

#End If

'---------------------------------------------------------------------------------------
' Property: LibPath
'---------------------------------------------------------------------------------------
'/**
' <summary>
' LibPath
' </summary>
' <param name="Vollständige Pfadangabe"></param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Friend Property Get LibPath() As String
   If Len(m_LibPath) = 0 Then
      m_LibPath = DefaultLibPath
   End If
   LibPath = m_LibPath
End Property

Friend Property Let LibPath(ByVal NewValue As String)
   If VBA.Right$(NewValue, 1) <> "\" Then
      NewValue = NewValue & "\"
   End If
   m_LibPath = NewValue
End Property

Friend Sub SetLibPathAsCurrentDir(Optional ByVal CurLibPath As String = vbNullString)
   
   If VBA.Len(CurLibPath) = 0 Then
      CurLibPath = LibPath
      If VBA.Len(CurLibPath) = 0 Then
         CurLibPath = DefaultLibPath
      End If
   End If
   VBA.ChDrive VBA.Left$(CurLibPath, 1)
   VBA.ChDir CurLibPath
   
End Sub

Friend Sub CheckFiles()
   
   Dim rst As DAO.Recordset

On Error GoTo Err_

   CheckLibPath

   Set rst = OpenAppFileInfoRecordset
   With rst
      Do While Not .EOF
         Call CheckFile(.Fields(m_TABLE_FIELD_FileName), Nz(.Fields(m_TABLE_FIELD_DestinationFolder), vbNullString), _
                        Nz(.Fields(m_TABLE_FIELD_Version).Value, vbNullString), _
                        Nz(.Fields(m_TABLE_FIELD_EarlyBindingRefName), vbNullString))
         .MoveNext
      Loop
      .Close
   End With
   Set rst = Nothing

Exit_:
   Exit Sub

Err_:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "LibFiles.CheckFiles()" & vbNewLine & "- " & Err.Source, _
                   Err.Description, Err.HelpFile, Err.HelpContext

End Sub

Private Sub CheckFile(ByVal FileName As String, ByVal DestinationFolder As String, _
             Optional ByVal sTabFileVersion As String = vbNullString, _
             Optional ByVal EarlyBindingRefName As String = vbNullString, _
             Optional ByVal AutoUpdate As Boolean = True)

   Dim InstalledVersionString As String
   Dim DoUpdate As Boolean
   Dim FullFilePath As String
   
   FullFilePath = GetFullFilePath(FileName, DestinationFolder, True)

   If VBA.Len(VBA.Dir$(FullFilePath)) > 0 Then
      If VBA.Len(sTabFileVersion) > 0 Then
         InstalledVersionString = GetFileVersion(FullFilePath)
         If IsNewVersion(sTabFileVersion, InstalledVersionString) Then
            If AutoUpdate Then
               DoUpdate = True
            ElseIf MsgBox("Es ist eine neuere Version von '" & FileName & "' vorhanden" & vbNewLine & _
                      "Soll die vorhandene Datei überschrieben werden?", _
                      vbQuestion + vbYesNo) = vbYes Then
               DoUpdate = True
            Else
               DoUpdate = False
            End If
         End If
      End If
   Else
      DoUpdate = True
   End If
   
   If DoUpdate Then
      Call ExportFile(FileName)
   End If

   If m_ReferencesUpdateAllowed Then
   If VBA.Len(EarlyBindingRefName) > 0 Then
      ' immer durchführen, auch wenn Dll vorhanden ist,
      ' damit immer auf die richtige dll verwiesen wird
      UpdateReference EarlyBindingRefName, FullFilePath
   End If
   End If

End Sub

Private Sub CheckLibPath()
   CreateDirectoryIfMissing LibPath
End Sub

Private Sub CreateDirectoryIfMissing(ByVal FullPath As String)
   
   Dim PathBefore As String
   Dim TextPos As Long
   Dim fso As Object 'Scripting.FileSystemObject

   If VBA.Len(VBA.Dir$(FullPath, vbDirectory)) > 0 Then 'Verzeichnis ist bereits vorhanden
      Exit Sub
   End If

   If VBA.Right$(FullPath, 1) = "\" Then
      FullPath = VBA.Left$(FullPath, VBA.Len(FullPath) - 1)
   End If

   TextPos = InStrRev(FullPath, "\") ' Z:\\dsf\
   If TextPos > 3 Then
      PathBefore = VBA.Mid$(FullPath, 1, TextPos - 1)
      If VBA.Len(VBA.Dir$(PathBefore, vbDirectory)) = 0 Then
         Call CreateDirectoryIfMissing(PathBefore)
      End If
   End If
  
   Set fso = VBA.CreateObject("Scripting.FileSystemObject")
   fso.CreateFolder FullPath
   Set fso = Nothing

End Sub

Private Function GetFullFilePath(ByVal FileName As String, _
                        Optional ByVal DestinationFolder As String, _
                        Optional ByVal CreateMissingDirectory As Boolean = False) As String
   
   If VBA.Len(DestinationFolder) = 0 Then
      DestinationFolder = LibPath
      If VBA.Len(DestinationFolder) = 0 Then
         Err.Raise vbObjectError, "GetFullFilePath", "Missing directory info"
      End If
   Else
      DestinationFolder = Replace(DestinationFolder, m_TABLE_DATA_CurrentProjectPathCode, ApplicationLibPath)
   End If

   If VBA.Right$(DestinationFolder, 1) <> "\" Then
      DestinationFolder = DestinationFolder & "\"
   End If

   If CreateMissingDirectory Then
      CreateDirectoryIfMissing DestinationFolder
   End If

   GetFullFilePath = DestinationFolder & FileName

End Function

Private Sub UpdateReference(ByVal RefName As String, ByVal FullPath As String)

On Error Resume Next
   RemoveReference RefName ' funktioniert nur, wenn Zugriff auf Datei möglich ist (irgendwie nonsens)
On Error GoTo 0
   Access.References.AddFromFile FullPath

End Sub

Private Sub RemoveReference(ByVal RefName As String)

   Dim ref As Access.Reference
   Dim refs As Access.References

   Set refs = Access.References
   For Each ref In refs
      If ref.Name = RefName Then
         refs.Remove ref
         Exit For
      End If
   Next

End Sub

Private Function IsNewVersion(ByVal sTabFileVersion As String, ByVal InstalledVersionString As String) As Boolean

   Dim installedVersion() As String
   Dim savedVersion() As String
   Dim i As Long

   If VBA.StrComp(sTabFileVersion, InstalledVersionString, vbTextCompare) = 0 Then
      IsNewVersion = False
      Exit Function
   End If

   installedVersion = Split(InstalledVersionString, ".")
   savedVersion = Split(sTabFileVersion, ".")

   For i = 0 To UBound(installedVersion)
      If Val(savedVersion(i)) > Val(installedVersion(i)) Then
         IsNewVersion = True
         Exit For
      End If
   Next

End Function

Private Property Get IsMDE() As Boolean

   Dim db As DAO.Database
   Dim prop As DAO.Property

   Set db = CodeDb
   For Each prop In db.Properties
      If prop.Name = "MDE" Then
         IsMDE = (prop.Value = "T")
         Exit Property
      End If
   Next

End Property

Friend Sub ImportFilesToTable()

   Dim rst As DAO.Recordset
   Dim FileNameField As DAO.Field
   Dim DestinationFolderField As DAO.Field
   Dim FileVersionField As DAO.Field
   
   Dim DestinationFolder As String
   Dim FullFilePath As String

On Error GoTo Err_

   Set rst = OpenAppFileInfoRecordset

   Set FileNameField = rst.Fields(m_TABLE_FIELD_FileName)
   Set DestinationFolderField = rst.Fields(m_TABLE_FIELD_DestinationFolder)
   Set FileVersionField = rst.Fields(m_TABLE_FIELD_Version)

   Do While Not rst.EOF
      ImportFileToTable Nz(FileNameField.Value, vbNullString), _
                        Nz(DestinationFolderField.Value, vbNullString), _
                        (Len(Nz(FileVersionField.Value, vbNullString)) > 0)
      rst.MoveNext
   Loop
   rst.Close

Exit_:
   Exit Sub

Err_:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If
   Err.Raise Err.Number, "LibFiles.CreateAppFile:" & Err.Source, Err.Description

End Sub

Public Sub ImportFileToTable(ByVal FileName As String, ByVal DestinationFolder As String, _
               Optional ByVal SaveVersion As Boolean = False, _
               Optional ByVal EarlyBindingRefName As String = vbNullString, _
               Optional ByVal ExternFilePath As String = vbNullString)

   Dim FileNr As Integer
   Dim BinFile() As Byte
   Dim ArraySize As Long
   Dim FileVersion As String
   Dim FullFilePath As String
   
On Error GoTo HandleErr

#If UseAssemblyDebugFolder Then
   If Len(AssemblyDebugFolder) > 0 Then
      FullFilePath = GetFullFilePath(FileName, AssemblyDebugFolder)
      If ExternFilePath <> vbNullString Then FullFilePath = AssemblyDebugFolder & "\" & FileName
   Else
#Else
   FullFilePath = GetFullFilePath(FileName, DestinationFolder)
   If ExternFilePath <> vbNullString Then FullFilePath = ExternFilePath & "\" & FileName
#End If
#If UseAssemblyDebugFolder Then
   End If
#End If
   
   If VBA.Len(VBA.Dir$(FullFilePath)) = 0 Then
      Err.Raise vbObjectError, "SaveAppFile", "file not found"
      Exit Sub
   End If

   FileNr = FreeFile
   Open FullFilePath For Binary As #FileNr
   ArraySize = LOF(FileNr) - 1
   If ArraySize >= 0 Then
      ReDim BinFile(ArraySize)
      Get #FileNr, , BinFile()
   End If
   Close #FileNr

   If ArraySize < 0 Then
      Err.Raise vbObjectError, "SaveAppFile", "Empty file (filesize = 0)"
      Exit Sub
   End If

   Dim rst As DAO.Recordset
   Set rst = OpenAppFileImportRecordset(FileName)
   If rst.EOF Then
      rst.AddNew
      rst.Fields(m_TABLE_FIELD_FileName) = FileName
      rst.Fields(m_TABLE_FIELD_DestinationFolder) = DestinationFolder
      If EarlyBindingRefName <> vbNullString Then rst.Fields(m_TABLE_FIELD_EarlyBindingRefName) = EarlyBindingRefName
   Else
      rst.Edit
   End If
   rst.Fields(m_TABLE_FIELD_FileData).AppendChunk BinFile
   If SaveVersion Then
      FileVersion = GetFileVersion(FullFilePath)
      If VBA.Len(FileVersion) = 0 Then
         FileVersion = Format$(Nz(FileDateTime(FullFilePath), vbNullString), "yyyy.mm.dd")
      End If
      rst.Fields(m_TABLE_FIELD_Version) = FileVersion
   End If
   rst.Update
   rst.Close
   Set rst = Nothing

ExitHere:
   Exit Sub
   
HandleErr:
   Dim tabCreateErrCnt As Long
   Dim lngErrNumber As Long, strErrDescription As String

   lngErrNumber = Err.Number
   strErrDescription = Err.Description

   If lngErrNumber = 3078 And tabCreateErrCnt = 0 Then
      tabCreateErrCnt = 1
      lngErrNumber = lngErrNumber
      Err.Clear
      If CreateAppFileTable Then
         Resume
      End If
   End If

   Err.Raise lngErrNumber, "LibFiles.SaveAppFile", strErrDescription

End Sub


Friend Sub ExportFile(Optional ByVal FileName As String = vbNullString)

   Dim rst As DAO.Recordset
   Dim SqlString As String
   
On Error GoTo Err_

   Set rst = OpenAppFileExportRecordset(FileName)
   ExportFilesFromRecordset rst
   rst.Close

Exit_:
   Exit Sub

Err_:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If
   Err.Raise Err.Number, "LibFiles.CreateAppFile:" & Err.Source, Err.Description

End Sub

Private Sub ExportFilesFromRecordset(ByVal rst As DAO.Recordset)
   
   Dim FileNameField As DAO.Field
   Dim DestinationFolderField As DAO.Field
   Dim FileDataField As DAO.Field
   
   Dim DestinationFolder As String
   Dim FullFilePath As String

   If rst.EOF Then
      Err.Raise vbObjectError, "ExportFilesFromRecordset", "Missing file data"
   End If

   Set FileNameField = rst.Fields(m_TABLE_FIELD_FileName)
   Set DestinationFolderField = rst.Fields(m_TABLE_FIELD_DestinationFolder)
   Set FileDataField = rst.Fields(m_TABLE_FIELD_FileData)

   Do
      DestinationFolder = Nz(DestinationFolderField.Value, vbNullString)
      FullFilePath = GetFullFilePath(FileNameField.Value, DestinationFolder, True)
      ExportFileFromRecordField FileDataField, FullFilePath
      rst.MoveNext
   Loop Until rst.EOF

End Sub

Private Sub ExportFileFromRecordField(FileDataField As DAO.Field, FullFilePath As String)

   Dim FieldSize As Long
   Dim BinFile() As Byte
   Dim FileNr As Integer

   FieldSize = FileDataField.FieldSize
   If FieldSize = 0 Then
      Err.Raise vbObjectError, "ExportFile", "Missing file data"
   End If

   ReDim BinFile(FieldSize - 1)

   BinFile = FileDataField.GetChunk(0, FieldSize)
   
   ' Datei erstellen
   FileNr = FreeFile
   Open FullFilePath For Binary As #FileNr
   Put #FileNr, , BinFile()
   Close #FileNr

End Sub

Public Function CreateAppFileTable() As Boolean
   
   Const CreateTableSql = _
      "create table " & m_TABLE_AppFiles & " ( " & _
                     m_TABLE_FIELD_FileName & " varchar(255) primary key, " & _
                     m_TABLE_FIELD_EarlyBindingRefName & " varchar(255), " & _
                     m_TABLE_FIELD_DestinationFolder & " varchar(255), " & _
                     m_TABLE_FIELD_Version & " varchar(20), " & _
                     m_TABLE_FIELD_FileData & " image)"

   CodeDb.Execute CreateTableSql, dbFailOnError

   CreateAppFileTable = True

End Function

Private Function OpenAppFileInfoRecordset() As DAO.Recordset

   Set OpenAppFileInfoRecordset = CodeDb.OpenRecordset(m_TABLE_InfoFields_SqlString, dbOpenForwardOnly, dbReadOnly)

End Function

Private Function OpenAppFileImportRecordset(ByVal FileName As String) As DAO.Recordset

   Dim strSQL As String

   strSQL = m_TABLE_AllFields_SqlString

   If VBA.Len(FileName) > 0 Then
      strSQL = strSQL & " where " & m_TABLE_FIELD_FileName & " = '" & FileName & "'"
   End If

   Set OpenAppFileImportRecordset = CodeDb.OpenRecordset(strSQL, dbOpenDynaset)

End Function

Private Function OpenAppFileExportRecordset(ByVal FileName As String) As DAO.Recordset

   Dim strSQL As String

   strSQL = m_TABLE_AllFields_SqlString

   If VBA.Len(FileName) > 0 Then
      strSQL = strSQL & " where " & m_TABLE_FIELD_FileName & " = '" & FileName & "'"
   End If

   Set OpenAppFileExportRecordset = CodeDb.OpenRecordset(strSQL, dbOpenForwardOnly, dbReadOnly)

End Function

'---------------------------------------------------------------------------------------------------
' API
'---------------------------------------------------------------------------------------------------
Private Function GetVersion(ByVal sPath As String, _
                            ByRef FInfo As FILEINFOOUT) As Boolean

  Dim lRet As Long, lSize As Long, lHandle As Long
  Dim lVerBufLen As Long, lVerPointer As Long
  Dim FileInfo As VS_FIXEDFILEINFO
  Dim sBuffer() As Byte

  lSize = GetFileVersionInfoSize(sPath, lHandle)
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  ReDim sBuffer(lSize)
  lRet = GetFileVersionInfo(sPath, 0&, lSize, sBuffer(0))
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  lRet = VerQueryValue(sBuffer(0), "\", lVerPointer, lVerBufLen)
  If lSize = 0 Then
    GetVersion = False
    Exit Function
  End If
  
  Const LenFileInfo As Long = 52 ' = Len(FileInfo) .. aber = VBA.Len(FileInfo) funktioniert nicht
  Call MoveMemory(FileInfo, lVerPointer, LenFileInfo)
  
  With FileInfo
  
    FInfo.FileVersion = _
      VBA.Trim$(VBA.Str$((.dwFileVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      VBA.Trim$(VBA.Str$(.dwFileVersionMS And &HFFFF&)) & "." & _
      VBA.Trim$(VBA.Str$((.dwFileVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      VBA.Trim$(VBA.Str$(.dwFileVersionLS And &HFFFF&))
    
    FInfo.ProductVersion = _
      VBA.Trim$(VBA.Str$((.dwProductVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      VBA.Trim$(VBA.Str$(.dwProductVersionMS And &HFFFF&)) & "." & _
      VBA.Trim$(VBA.Str$((.dwProductVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      VBA.Trim$(VBA.Str$(.dwProductVersionLS And &HFFFF&))
      
  End With
  
  GetVersion = True

End Function

Private Function GetFileVersion(ByVal sFile As String) As String
   Dim FileVersionInfo As FILEINFOOUT
   If GetVersion(sFile, FileVersionInfo) Then
      GetFileVersion = FileVersionInfo.FileVersion
   Else
      GetFileVersion = vbNullString
   End If
End Function

Private Function UNCPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String

  Dim UNC As String * 512

  If VBA.Len(Path) = 1 Then Path = Path & ":"

  If WNetGetConnection(Left$(Path, 2), UNC, Len(UNC)) Then

    ' API-Routine gibt Fehler zurück:
    If IgnoreErrors Then
      UNCPath = Path
    Else
      Err.Raise 5 ' Invalid procedure call or argument
    End If

  Else

    ' Ergebnis zurückgeben:
    UNCPath = Left$(UNC, InStr(UNC, vbNullChar) - 1) _
            & Mid$(Path, 3)

  End If

End Function

Public Function GetSpecFolder(ByVal lCSIDL As CSIDL_FOLDER, _
      Optional ByVal bCreate As Boolean = False, _
      Optional ByVal bVerify As Boolean = False) As String
      
   Dim sPath As String, RetVal As Long, lFlags As Long

   sPath = VBA.String$(MAX_PATH, 0)
   lFlags = lCSIDL
   If bCreate Then lFlags = lFlags Or CSIDL_FLAG_CREATE
   If Not bVerify Then lFlags = lFlags Or CSIDL_FLAG_DONT_VERIFY
   RetVal = SHGetFolderPath(0, lFlags, 0, SHGFP_TYPE_CURRENT, sPath)
   Select Case RetVal
   Case 0
      ' Verzeichnis gefunden
      GetSpecFolder = VBA.Left$(sPath, InStr(1, sPath, Chr$(0)) - 1)
   Case 1
      ' lCSIDL ist gültig, aber das Verzeichnis existiert nicht
      ' CSIDL_FLAG_CREATE erzeugt es automatisch
      Err.Raise vbObjectError + 1, "GetSpecFolder", "Verzeichnis existiert nicht"
   Case &H80070057
      ' Ungültiges Verzeichnis
      Err.Raise vbObjectError + 2, "GetSpecFolder", "Ungültiger Verzeichnisbezeichner (CSIDL)"
   Case Else
      Err.Raise vbObjectError + 999, "GetSpecFolder", "Unerwartete Rückgabe der API-Funktion SHGetFolderPath"
   End Select

End Function
