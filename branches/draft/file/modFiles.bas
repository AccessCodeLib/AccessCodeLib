Attribute VB_Name = "modFiles"
Attribute VB_Description = "Funktionen für Dateioperationen"
'---------------------------------------------------------------------------------------
' Module: modFiles (Josef Pötzl, 2009-12-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Funktionen für Dateioperationen
' </summary>
' <remarks>
' </remarks>
'\ingroup file
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>file/modFiles.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <use>base/defGlobal.bas</use>
'  <test>_test/file/Test_modFiles.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Zuordnung der Prozeduren zur Doxygen-Gruppe:
'/** \addtogroup file
'@{ **/

Private Const m_SELECTBOX_File_DlgTitle As String = "Datei auswählen"
Private Const m_SELECTBOX_Folder_DlgTitle As String = "Ordner auswählen"
Private Const m_SELECTBOX_OpenTitle As String = "auswählen"

Private Const m_DEFAULT_TEMPPATH_NoEnv As String = "C:\"
Private Const m_MAXPATHLEN As Long = 255

Private Declare Function WNetGetConnection Lib "mpr.dll" Alias "WNetGetConnectionA" ( _
         ByVal lpszLocalName As String, ByVal lpszRemoteName As String, cbRemoteName As Long) As Long

Private Declare Function API_GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long




'---------------------------------------------------------------------------------------
' Function: SelectFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei mittels Dialog auswählen
' </summary>
' <param name="InitDir">Startverzeichnis</param>
' <param name="DlgTitle">Dialogtitel</param>
' <param name="FilterString">Filterwerten - Beispiel: "(*.*)" oder "Alle (*.*)|Textdateien (*.txt)|Bilder (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Mehrfachauswahl</param>
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Große Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFile(Optional ByVal InitialDir As String = vbNullString, _
                           Optional ByVal DlgTitle As String = m_SELECTBOX_File_DlgTitle, _
                           Optional ByVal FilterString As String = "Alle Dateien (*.*)", _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal ViewMode As Long = -1) As String

On Error GoTo HandleErr

    SelectFile = WizHook_GetFileName(InitialDir, DlgTitle, m_SELECTBOX_OpenTitle, FilterString, MultiSelectEnabled, , ViewMode, False)

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "SelectFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: SelectFolder (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Auswahldialog zur Verzeichnisauswahl
' </summary>
' <param name="InitDir">Startverzeichnis</param>
' <param name="DlgTitle">Dialogtitel</param>
' <param name="FilterString">Filterwerten - Beispiel: "(*.*)" oder "Alle (*.*)|Textdateien (*.txt)|Bilder (*.png;*.jpg;*.gif)</param>
' <param name="MultiSelect">Mehrfachauswahl</param>
' <param name="viewMode">Anzeigeart (0: Detailansicht, 1: Vorschauansicht, 2: Eigenschaften, 3: Liste, 4: Miniaturansicht, 5: Große Symbole, 6: Kleine Symbole)</param>
' <returns>String (bei Mehfachauswahl sind die Dateien durch chr(9) getrennt)</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function SelectFolder(Optional ByVal InitialDir As String = vbNullString, _
                             Optional ByVal DlgTitle As String = m_SELECTBOX_Folder_DlgTitle, _
                             Optional ByVal FilterString As String = "*", _
                             Optional ByVal MultiSelectEnabled As Boolean = False, _
                             Optional ByVal ViewMode As Long = -1) As String

On Error GoTo HandleErr

   SelectFolder = WizHook_GetFileName(InitialDir, DlgTitle, m_SELECTBOX_OpenTitle, FilterString, MultiSelectEnabled, , ViewMode, True)

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "SelectFolder", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function


Private Function WizHook_GetFileName( _
                           ByVal InitialDir As String, _
                           ByVal DlgTitle As String, _
                           ByVal OpenTitle As String, _
                           ByVal FilterString As String, _
                           Optional ByVal MultiSelectEnabled As Boolean = False, _
                           Optional ByVal SplitDelimiter As String = "|", _
                           Optional ByVal ViewMode As Long = -1, _
                           Optional ByVal SelectFolderFlag As Boolean = False) As String

'Zusammenfassung der Parameter von WizHook.GetFileName: http://www.team-moeller.de/?Tipps_und_Tricks:Wizhook-Objekt:GetFileName
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

   Dim selectedFileString As String
   Dim wizHookRetVal As Long

On Error GoTo HandleErr

   If InStr(1, InitialDir, " ") > 0 Then
      InitialDir = """" & InitialDir & """"
   End If

   Dim flags As Long
   flags = 0
   If MultiSelectEnabled Then flags = flags + 8
   If SelectFolderFlag Then flags = flags + 32

   If ViewMode >= 0 Then
      flags = flags + 64
   Else
      ViewMode = 0
   End If

   WizHook.Key = 51488399
   wizHookRetVal = WizHook.GetFileName( _
                        Access.Application.hWndAccessApp, CurrentApplicationName, DlgTitle, OpenTitle, _
                        selectedFileString, InitialDir, FilterString, 0, ViewMode, flags, True)
   If wizHookRetVal = 0 Then
      If MultiSelectEnabled Then selectedFileString = Replace(selectedFileString, vbTab, SplitDelimiter)
      WizHook_GetFileName = selectedFileString
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "WizHook_GetFileName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Public Function UNCPath(ByVal Path As String, Optional ByVal IgnoreErrors As Boolean = True) As String

  Dim UNC As String * 512

On Error GoTo HandleErr

  If Len(Path) = 1 Then Path = Path & ":"

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

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "UNCPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function

'---------------------------------------------------------------------------------------
' Property: TempPath (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Temp-Verzeichnis ermitteln
' </summary>
' <param name="Param"></param>
' <returns>String</returns>
' <remarks>
' Verwendet API GetTempPathA
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TempPath() As String

   Dim strTemp As String

On Error GoTo HandleErr

   strTemp = Space$(m_MAXPATHLEN)
   API_GetTempPath m_MAXPATHLEN, strTemp
   strTemp = Left$(strTemp, InStr(strTemp, Chr$(0)) - 1)
   If Len(strTemp) = 0 Then
      strTemp = m_DEFAULT_TEMPPATH_NoEnv
   End If
   TempPath = strTemp

ExitHere:
   Exit Property

HandleErr:
   Select Case HandleError(Err.Number, "TempPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Property

'---------------------------------------------------------------------------------------
' Function: ShortFileName (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateipfad auf n Zeichen kürzen
' </summary>
' <param name="vFile">Vollständiger Pfad</param>
' <param name="lMaxLen">die gewünschte Länge</param>
' <returns>String</returns>
' <remarks>
' Hilfreich für die Anzeigen in schmalen Textfeldern \n
' Beispiel: <source>C:\Programme\...\Verzeichnis\Dateiname.txt</source>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ShortFileName(ByVal vFile As Variant, Optional ByVal lMaxLen As Long = 40) As String

   Dim strFile As String
   Dim strTemp As String
   Dim lngPos As Long

On Error GoTo HandleErr

   If IsNull(vFile) Then
      strFile = vbNullString
   Else
      strFile = vFile
   End If
   If Len(strFile) > lMaxLen Then
      lngPos = InStrRev(strFile, "\")
      strTemp = Mid$(strFile, lngPos)
      strFile = Left$(strFile, lngPos - 1)

      lngPos = lMaxLen - Len(strTemp) - 3
      If lngPos < 2 Then
         strTemp = "..." & strTemp
      Else
         lngPos = lngPos \ 2
         strTemp = Left$(strFile, lngPos) & "..." & Right$(strFile, lngPos) & strTemp
      End If
      strFile = strTemp
   End If

   ShortFileName = strFile

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ShortFileName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: FileNameWithoutPath (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateinamen ohne Verzeichnis
' </summary>
' <param name="vFile">Dateiname inkl. Verzeichnis</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileNameWithoutPath(ByVal vFile As Variant) As String

   Dim sFile As String
   Dim lngPos As Long

On Error GoTo HandleErr

   sFile = Nz(vFile, vbNullString)
   lngPos = InStrRev(sFile, "\")
   If lngPos > 0 Then
      FileNameWithoutPath = Mid$(sFile, lngPos + 1)
   Else
      FileNameWithoutPath = sFile
   End If

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "FileNameWithoutPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case Else
      Resume ExitHere
   End Select

End Function


'---------------------------------------------------------------------------------------
' Function: CreateDirectory (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erstelle ein Verzeichnis inkl. aller fehlenden übergeordneten Verzeichnisse
' </summary>
' <param name="spath">Zu erstellendes Verzeichnis</param>
' <returns>Boolean: True = Verzeichnis wurde erstellt</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateDirectory(ByVal sPath As String) As Boolean

   Dim strPathBefore As String

On Error GoTo HandleErr

   If Len(Dir$(sPath, vbDirectory)) > 0 Then 'Verzeichnis ist bereits vorhanden
      CreateDirectory = False
      Exit Function
   End If

   strPathBefore = Mid$(sPath, 1, InStrRev(sPath, "\") - 1)
   If Len(Dir$(strPathBefore, vbDirectory)) = 0 Then
      If CreateDirectory(strPathBefore) = False Then
         CreateDirectory = False
         Exit Function
      End If
   End If

   MkDir sPath

   CreateDirectory = True

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "CreateDirectory", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: FileExists (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft Existens einer Datei
' </summary>
' <param name="PathName">Vollständige Pfadangabe</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function FileExists(ByVal PathName As String) As Boolean

On Error GoTo HandleErr

   Do While Right$(PathName, 1) = "\"
      PathName = Left$(PathName, Len(PathName) - 1)
   Loop
   FileExists = (Len(Dir$(PathName, vbReadOnly Or vbHidden Or vbSystem)) > 0)
      '6 = vbNormal or vbHidden or vbSystem

ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "FileExists", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: DirExists (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft Existenz eines Verzeichnisses
' </summary>
' <param name="PathName">Vollständige Pfadangabe</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DirExists(ByVal PathName As String) As Boolean

On Error GoTo HandleErr

   If Right$(PathName, 1) <> "\" Then
      PathName = PathName & "\"
   End If

   DirExists = (Dir$(PathName, vbDirectory Or vbReadOnly Or vbHidden Or vbSystem) = ".")

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "DirExists", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetFileUpdateDate (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Letztes Änderungsdatum einer Datei
' </summary>
' <param name="FullFileName">Vollständige Pfadangabe</param>
' <returns>Variant</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileUpdateDate(ByVal FullFileName As String) As Variant
On Error Resume Next
   If Len(Dir$(FullFileName)) > 0 Then
      GetFileUpdateDate = FileDateTime(FullFileName)
   Else
      GetFileUpdateDate = Null
   End If
End Function


'---------------------------------------------------------------------------------------
' Function: GetClearedStringForFileName (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt aus einer Zeichenkette einen Dateinamen (ersetzt Sonderzeichen)
' </summary>
' <param name="strName">Ausgangsstring für Dateinamen</param>
' <param name="ReplacementSign">Zeichen als Ersatz für Sonderzeichen</param>
' <returns>String</returns>
' <remarks>
' Sonderzeichen: ? * " / ' : ( )
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetClearedStringForFileName(ByVal strName As String, Optional ByVal ReplacementSign As String = "_") As String

On Error Resume Next

   Dim strTemp As String

   strTemp = Replace(strTemp, "?", vbNullString)
   strTemp = Replace(strTemp, "*", vbNullString)
   strTemp = Replace(strTemp, """", vbNullString)

   strTemp = Replace(strName, "/", ReplacementSign)
   strTemp = Replace(strTemp, "'", ReplacementSign)
   strTemp = Replace(strTemp, ":", ReplacementSign)
   strTemp = Replace(strTemp, "(", ReplacementSign)
   strTemp = Replace(strTemp, ")", ReplacementSign)
   'strTemp = Replace(strTemp, " ", ReplacementSign)

   GetClearedStringForFileName = strTemp

End Function

'---------------------------------------------------------------------------------------
' Function: GetFullPathFromRelativPath (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erezugt aus relativer Pfadangabe und "Basisverzeichnis" eine vollständige Pfadangabe
' </summary>
' <param name="sRelativPath">relativer Pfad</param>
' <param name="sBaseFolder">Ausgangsverzeichnis</param>
' <returns>String</returns>
' <remarks>
' Beispiel:
' GetFullPathFromRelativPath("..\..\Test.txt", "C:\Programme\xxx\") => "C:\test.txt"
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFullPathFromRelativPath(ByVal sRelativPath As String, ByVal sBaseFolder As String) As String

   Dim strPath As String
   Dim strBase As String
   Dim lngPos As Long

On Error GoTo HandleErr

   strBase = sBaseFolder
   If Right$(strBase, 1) = "\" Then
      strBase = Left$(strBase, Len(strBase) - 1)
   End If

   strPath = sRelativPath
   If Mid$(strPath, 2, 1) = ":" Or Left$(strPath, 2) = "\\" Then ' absolut path !!!
      GetFullPathFromRelativPath = strPath
      Exit Function
   ElseIf Left$(strPath, 1) = "\" Then 'first dir
      lngPos = InStr(3, strBase, "\")
      If lngPos > 0 Then
         strBase = Left$(strBase, lngPos - 1)
      End If
      GetFullPathFromRelativPath = strBase & strPath
      Exit Function
   ElseIf strPath = "." Then
      GetFullPathFromRelativPath = strBase
      Exit Function
   ElseIf Left$(strPath, 2) = ".\" Then
      strPath = Mid$(strPath, 3)
   End If

   Do While Left$(strPath, 3) = "..\"
      strPath = Mid$(strPath, 4)
      lngPos = InStrRev(strBase, "\")
      If lngPos > 0 Then
         strBase = Left$(strBase, lngPos - 1)
      End If
   Loop

   GetFullPathFromRelativPath = strBase & "\" & strPath

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetFullPathFromRelativPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetRelativPathFromFullPath (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt einen relativen Pfad aus vollständiger Pfadangabe und Ausgangsverzeichnis
' </summary>
' <param name="sFullPath">vollständiger Pfadangabe</param>
' <param name="sBaseFolder">Ausgangsverzeichnis</param>
' <param name="RelativePrefix">".\" als Kennung für relativen Pfad ergänzen</param>
' <returns>String</returns>
' <remarks>
' Beispiel:
' <code>
' GetRelativPathFromFullPath("C:\test.txt", "C:\Programme\xxx\", True)
' => ".\..\..\test.txt"
' </code>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetRelativPathFromFullPath( _
                        ByVal sFullPath As String, ByVal sBaseFolder As String, _
                        Optional ByVal EnableRelativePrefix As Boolean = False) As String

   Dim strPath As String
   Dim strRetPath As String
   Dim lngPos As Long

   Dim lngRetCounter As Long, i As Long

On Error GoTo HandleErr

   If sFullPath = sBaseFolder Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   If Right$(sBaseFolder, 1) <> "\" Then sBaseFolder = sBaseFolder & "\"
   If sFullPath = sBaseFolder Then
      GetRelativPathFromFullPath = "."
      Exit Function
   End If

   strPath = sBaseFolder

   Do While InStr(1, sFullPath, strPath) = 0
      lngPos = InStrRev(Left$(strPath, Len(strPath) - 1), "\")
      strPath = Left$(strPath, lngPos)
      lngRetCounter = lngRetCounter + 1
      If Len(strPath) = 0 Then
         lngRetCounter = 0
         Exit Do
      End If
   Loop

   If Len(strPath) > 0 Then
      strRetPath = Replace(sFullPath, strPath, vbNullString)
      For i = 1 To lngRetCounter
         strRetPath = "..\" & strRetPath
      Next

      If EnableRelativePrefix Then
         strRetPath = ".\" & strRetPath
      End If

   Else
      strRetPath = sFullPath
   End If

   GetRelativPathFromFullPath = strRetPath

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetRelativPathFromFullPath", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

'---------------------------------------------------------------------------------------
' Function: GetDirFromFilePath (2009-11-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittels aus vollständer Pfadangabe einer Datei das Verzeichnis
' </summary>
' <param name="sFileName">vollständer Pfadangabe</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetDirFromFilePath(ByVal sFileName As String) As String

   Dim strPath As String
   Dim lngPos As Long

On Error GoTo HandleErr

   strPath = sFileName
   lngPos = InStrRev(strPath, "\")
   If lngPos > 0 Then
      strPath = Left$(strPath, lngPos)
   Else
      strPath = vbNullString
   End If

   GetDirFromFilePath = strPath

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetPathFromFullFileName", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function


'---------------------------------------------------------------------------------------
' Sub: AddToZipFile (2009-11-09)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei an Zip-Datei anhängen.
' </summary>
' <param name="Param"></param>
' <returns></returns>
' <remarks>
' CreateObject("Shell.Application").Namespace(zipFile & "").CopyHere sFile & ""
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub AddToZipFile(ByVal zipFile As String, ByVal sFile As String)

On Error GoTo HandleErr

   If Len(Dir$(zipFile)) = 0 Then
      NewZip zipFile
   End If

   With CreateObject("Shell.Application")
      .Namespace(zipFile & "").CopyHere sFile & ""
   End With

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "AddToZipFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Sub

Private Function ExtractFromZipFile(ByVal zipFile As String, ByVal Destination As String) As String

On Error GoTo HandleErr

   With CreateObject("Shell.Application")
      .Namespace(Destination & "").CopyHere .Namespace(zipFile & "").Items
      ExtractFromZipFile = .Namespace(zipFile & "").Items.Item(0).Name
   End With

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "ExtractFromZipFile", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function

Private Sub NewZip(ByVal zipFile As String)

On Error GoTo HandleErr

    If Len(Dir$(zipFile)) > 0 Then Kill zipFile
    Open zipFile For Output As #1
    Print #1, Chr$(80) & Chr$(75) & Chr$(5) & Chr$(6) & String$(18, 0)
    Close #1

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case HandleError(Err.Number, "NewZip", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
