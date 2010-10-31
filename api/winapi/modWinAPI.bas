Attribute VB_Name = "modWinAPI"
Attribute VB_Description = "Gebräuchliche WinAPI-Funktionen"
'---------------------------------------------------------------------------------------
' Module: modWinAPI
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gebräuchliche WinAPI-Funktionen
' </summary>
' <remarks>
' Sammlung von API-Deklarationen, die oft benötigt werden
' </remarks>
'\ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
   pDest As Any, _
   pSource As Any, _
   ByVal dwLength As Long)

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
   ByVal Hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
'--
Private Const WM_MSG_SETICON As Long = &H80
Private Const WM_PARAM_ICON_SMALL As Long = 0

Private Declare Function LoadImage Lib "user32" Alias "LoadImageA" ( _
   ByVal hInst As Long, _
   ByVal lpszName As String, _
   ByVal uType As Long, _
   ByVal cxDesired As Long, _
   ByVal cyDesired As Long, _
   ByVal fuLoad As Long) As Long
   
'uType
'Private Const IMAGE_BITMAP As Long = 0
Private Const IMAGE_ICON As Long = 1
'Private Const IMAGE_CURSOR As Long = 2

'fuLoad
Private Const LR_LOADFROMFILE As Long = &H10

'--

Private Declare Function ShellExecuteA Lib "shell32.dll" ( _
   ByVal Hwnd As Long, _
   ByVal lOperation As String, _
   ByVal lpFile As String, _
   ByVal lpParameters As String, _
   ByVal lpDirectory As String, _
   ByVal nShowCmd As Long) As Long
   
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Const SE_ERR_NOTFOUND As Long = 2
Private Const SE_ERR_NOASSOC  As Long = 31

Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" ( _
   ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long
   
Private Const STARTF_USESHOWWINDOW As Long = &H1
Private Const NORMAL_PRIORITY_CLASS As Long = &H20

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" ( _
   ByVal lpApplicationName As String, ByVal lpCommandLine As String, _
   ByVal lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   lpEnvironment As Any, ByVal lpCurrentDirectory As String, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long

Private Const INFINITE As Long = &HFFFFFFFF ' = -1&
Private Const WAIT_TIMEOUT As Long = &H102&

Private Declare Function WaitForInputIdle Lib "user32" (ByVal hProcess As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'---------------------------------------------------------------------------------------
' Kapselungen
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Function: ShellExecuteOpenFile (Josef Pötzl, 2010-04-19)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datei mit ShellExecute öffnen
' </summary>
' <param name="sFile">vollständiger Dateiname inkl. Verzeichnis</param>
' <param name="sAPIOperation">"open", "print", ...</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ShellExecuteOpenFile(ByVal sFile As String, _
                            Optional ByVal sAPIOperation As String = vbNullString) As Boolean

   Dim lRet As Long
   Dim sDirectory As String
   Dim lngDeskWin As Long
   
   If sFile = vbNullString Then
      ShellExecuteOpenFile = False
      Exit Function
   Else
      lngDeskWin = GetDesktopWindow()
      lRet = ShellExecuteA(lngDeskWin, sAPIOperation, sFile, vbNullString, vbNullString, vbNormalFocus)
   End If
   
   If lRet = SE_ERR_NOTFOUND Then
      'Datei nicht gefunden
      MsgBox "Datei nicht gefunden" & vbNewLine & vbNewLine & _
            sFile
      ShellExecuteOpenFile = False
      Exit Function
   ElseIf lRet = SE_ERR_NOASSOC Then
      'Wenn die Dateierweiterung noch nicht bekannt ist...
      'wird der "Öffnen mit..."-Dialog angezeigt.
      sDirectory = Space$(260)
      lRet = GetSystemDirectory(sDirectory, Len(sDirectory))
      sDirectory = Left$(sDirectory, lRet)
      Call ShellExecuteA(lngDeskWin, vbNullString, "RUNDLL32.EXE", "shell32.dll, OpenAs_RunDLL " & _
         sFile, sDirectory, vbNormalFocus)
   End If
   
   ShellExecuteOpenFile = True

End Function

'---------------------------------------------------------------------------------------
' Function: ShellExecuteSendMail
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Email mit Standard-Programm versenden
' </summary>
' <param name="sTo">Empfänger-Adresse</param>
' <param name="sSubject">Betreff-Zeile</param>
' <param name="sBody">Email-Text</param>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ShellExecuteSendMail(ByVal sTo As String, _
                                     ByVal sSubject As String, _
                                     ByVal sBody As String) As Boolean

   Dim lRet As Long
   Dim strLpFile As String
   
   If Len(sTo) = 0 Then
      ShellExecuteSendMail = False
      Exit Function
   End If
   
   If sSubject > vbNullString Then
      strLpFile = "subject=" & sSubject
   End If
   If sBody > vbNullString Then
      If strLpFile > vbNullString Then
         strLpFile = strLpFile & "&body=" & sBody
      Else
         strLpFile = "body=" & sBody
      End If
   End If
   If strLpFile > vbNullString Then
       strLpFile = "mailto:" & sTo & "?" & strLpFile
   Else
      strLpFile = "mailto:" & sTo
   End If

   
   
   lRet = ShellExecuteA(GetDesktopWindow(), "open", strLpFile, vbNullString, vbNullString, vbNormalFocus)
   ShellExecuteSendMail = (lRet <> 0)

End Function

'---------------------------------------------------------------------------------------
' Function: LaunchAppSynchronous (Josef Pötzl, 2010-04-19)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Anwnedung Synchron ausführen
' </summary>
' <param name="strExecutablePathAndName">Ausführbare Datei</param>
' <param name="sParam">Startparameter der Anwendung</param>
' <param name="lShowCommand">Fenstermodus</param>
' <returns>Boolean</returns>
' <remarks>
' Code hält so lange an, bis die gestartete Anwendung beendet wurde.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LaunchAppSynchronous(ByVal strExecutablePathAndName As String, _
                     Optional ByVal sParam As String = vbNullString, _
                     Optional ByVal lShowCommand As Long = vbNormalFocus) As Boolean
   
   'http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=3716&lngWId=1

   Dim lngResponse As Long
   Dim typStartUpInfo As STARTUPINFO
   Dim typProcessInfo As PROCESS_INFORMATION

   LaunchAppSynchronous = False

   With typStartUpInfo
      .cb = Len(typStartUpInfo)
      .lpReserved = vbNullString
      .lpDesktop = vbNullString
      .lpTitle = vbNullString
      .dwFlags = STARTF_USESHOWWINDOW
      .wShowWindow = lShowCommand
   End With

   'Launch the application by creating a ne
   '    w process
   lngResponse = CreateProcess(vbNullString, strExecutablePathAndName & " " & sParam, 0, 0, True, NORMAL_PRIORITY_CLASS, ByVal 0&, vbNullString, typStartUpInfo, typProcessInfo)


   If lngResponse Then
      'Wait for the application to terminate b
      '    efore moving on
      Call WaitForTermination(typProcessInfo)
      LaunchAppSynchronous = True
   Else
      LaunchAppSynchronous = False
   End If

End Function

Private Sub WaitForTermination(ByRef typProcessInfo As PROCESS_INFORMATION)
   'This wait routine allows other applicat
   '    ion events
   'to be processed while waiting for the p
   '    rocess to
   'complete.
   
   Dim lngResponse As Long
   'Let the process initialize
   Call WaitForInputIdle(typProcessInfo.hProcess, INFINITE)
   'We don't need the thread handle so get
   '    rid of it
   Call CloseHandle(typProcessInfo.hThread)
   'Wait for the application to end

   Do
      lngResponse = WaitForSingleObject(typProcessInfo.hProcess, 0)
      If lngResponse <> WAIT_TIMEOUT Then
         'No timeout, app is terminated
         Exit Do
      End If
      DoEvents
      Loop While True

      'Kill the last handle of the process
      Call CloseHandle(typProcessInfo.hProcess)

End Sub

'---------------------------------------------------------------------------------------
' Sub: SetFormIconFromFile (Josef Pötzl, 2009-12-19)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Spezielles Icon für ein Formular einstellen
' </summary>
' <param name="FormRef">Referenz zum Access.Form</param>
' <param name="IconFilePath">Icon-Datei (vollständige Pfadangabe)</param>
' <remarks>
'
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub SetFormIconFromFile(ByRef formRef As Access.Form, ByVal IconFilePath As String)
   
On Error Resume Next ' ... Fehlermeldung würde bei dieser "unwichtigen" Funktion nur stören
  
   Const ICONPIXELSIZE As Long = 16
   
   Dim imageHandle As Long
   
   imageHandle = LoadImage(0, IconFilePath, IMAGE_ICON, _
                           ICONPIXELSIZE, ICONPIXELSIZE, LR_LOADFROMFILE)
   If imageHandle <> 0 Then
      SendMessage formRef.Hwnd, WM_MSG_SETICON, WM_PARAM_ICON_SMALL, ByVal imageHandle
   End If
   
End Sub
