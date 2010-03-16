Attribute VB_Name = "modWinAPI_SpecFolder"
'---------------------------------------------------------------------------------------
' Module: modWinAPI_SpecFolder (Josef Pötzl, 2010-03-13)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Kapselung von SHGetFolderPath
' </summary>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_SpecFolder.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private Const CSIDL_FLAG_CREATE = &H8000&

Public Enum CSIDL_FOLDER
   CSIDL_DESKTOP = &H0& ' Desktop
   CSIDL_INTERNET = &H1& ' Internet
   CSIDL_PROGRAMS = &H2& ' Startmenü: Programme
   CSIDL_CONTROLS = &H3& ' Systemsteuerung
   CSIDL_PRINTERS = &H4& ' Drucker
   CSIDL_PERSONAL = &H5& ' Eigene Dateien
   CSIDL_FAVORITES = &H6& ' Favoriten
   CSIDL_STARTUP = &H7& ' Autostart
   CSIDL_RECENT = &H8& ' Zuletzt benutzte Dokumente
   CSIDL_SENDTO = &H9& ' Senden an
   CSIDL_BITBUCKET = &HA& ' Papierkorb
   CSIDL_STARTMENU = &HB& ' Startmenü
   CSIDL_DESKTOPDIRECTORY = &H10& ' Desktopverzeichnis
   CSIDL_DRIVES = &H11& ' Mein Computer
   CSIDL_NETWORK = &H12& ' Netzwerk
   CSIDL_NETHOOD = &H13& ' Netzwerkumgebung
   CSIDL_FONTS = &H14& ' Fonts
   CSIDL_TEMPLATES = &H15& ' Vorlagen
   CSIDL_COMMON_STARTMENU = &H16& ' "All Users"-Profil - Startmenü
   CSIDL_COMMON_PROGRAMS = &H17& ' "All Users"-Profil - Programme
   CSIDL_COMMON_STARTUP = &H18& ' "All Users"-Profil - Autostart
   CSIDL_COMMON_DESKTOPDIRECTORY = &H19& ' Desktopverzeichnis (Allgemein)
   CSIDL_APPDATA = &H1A& ' Anwendungsdaten
   CSIDL_PRINTHOOD = &H1B& ' Druckumgebung
   CSIDL_LOCAL_APPDATA = &H1C& ' Lokale Anwendungsdaten
   CSIDL_ALTSTARTUP = &H1D&
   CSIDL_COMMON_ALTSTARTUP = &H1E&
   CSIDL_COMMON_FAVORITES = &H1F&
   CSIDL_INTERNET_CACHE = &H20& ' Temporäre Internetdateien (MSIE)
   CSIDL_COOKIES = &H21& ' Cookies (MSIE)
   CSIDL_HISTORY = &H22& ' Verlauf (MSIE)
   CSIDL_COMMON_APPDATA = &H23& ' Anwendungsdaten (Allgemein)
   CSIDL_WINDOWS = &H24& ' Windows
   CSIDL_SYSTEM = &H25& ' Windows-System
   CSIDL_PROGRAM_FILES = &H26& ' Programme
   CSIDL_MYPICTURES = &H27& ' Eigene Bilder
   CSIDL_PROFILE = &H28&
   CSIDL_SYSTEMX86 = &H29&
   CSIDL_PROGRAM_FILESX86 = &H2A&
   CSIDL_PROGRAM_FILES_COMMON = &H2B& ' Gemeinsame Dateien
   CSIDL_PROGRAM_FILES_COMMONX86 = &H2C&
   CSIDL_COMMON_TEMPLATES = &H2D& ' Vorlagen (Allgemein)
   CSIDL_COMMON_DOCUMENTS = &H2E& ' Dokumente (Allgemein)
   CSIDL_COMMON_ADMINTOOLS = &H2F& ' Verwaltung (Allgemein)
   CSIDL_ADMINTOOLS = &H30&
   CSIDL_CONNECTIONS = &H31&
   CSIDL_FOLDER_MASK = &HFF&
End Enum

Private Const CSIDL_FLAG_DONT_VERIFY = &H4000&
Private Const SHGFP_TYPE_CURRENT = 0
Private Const MAX_PATH = 260

Private Declare Function SHGetFolderPath Lib "shfolder" Alias "SHGetFolderPathA" ( _
   ByVal hwndOwner As Long, ByVal nFolder As Long, _
   ByVal hToken As Long, ByVal dwFlags As Long, _
   ByVal pszPath As String) As Long

Public Function GetSpecFolder(lCSIDL As CSIDL_FOLDER, _
      Optional bCreate As Boolean = False, _
      Optional bVerify As Boolean = False) As String
      
   Dim sPath As String, RetVal As Long, lFlags As Long
  
On Error GoTo HandleErr

   sPath = String(MAX_PATH, 0)
   lFlags = lCSIDL
   If bCreate Then lFlags = lFlags Or CSIDL_FLAG_CREATE
   If Not bCreate Then lFlags = lFlags Or CSIDL_FLAG_DONT_VERIFY
   RetVal = SHGetFolderPath(0, lFlags, 0, SHGFP_TYPE_CURRENT, sPath)
   Select Case RetVal
   Case 0
      ' Verzeichnis gefunden
      GetSpecFolder = Left(sPath, InStr(1, sPath, Chr(0)) - 1)
   Case 1
      ' lCSIDL ist gültig, aber das Verzeichnis existiert nicht
      ' CSIDL_FLAG_CREATE erzeugt es automatisch
      Err.Raise vbObjectError + 1, "GetSpecFolder", "Verzeichnis existiert nicht"
   Case &H80070057
      ' Ungültiges Verzeichnis
      Err.Raise vbObjectError + 2, "GetSpecFolder", "Ungültiger Verzeichnisbezeichner (CSIDL)"
   End Select


ExitHere:
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetSpecFolder", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function
