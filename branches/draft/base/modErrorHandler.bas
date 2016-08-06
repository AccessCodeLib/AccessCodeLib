Attribute VB_Name = "modErrorHandler"
Attribute VB_Description = "Prozeduren für die Fehlerbehandlung"
'---------------------------------------------------------------------------------------
' Modul: modErrorHandler (2009-12-15)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prozeduren für die Fehlerbehandlung
' </summary>
' <remarks></remarks>
'\ingroup base
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>base/modErrorHandler.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorHandlerMode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ErrorHandler Modes (Fehlerbehandlungsvarianten)
' </summary>
' <list type="table">
'   <item><term>aclibErrRaise (0)</term><description>Weitergabe an Anwendung</description></item>
'   <item><term>aclibErrMsgBox (1)</term><description>Fehler in MsgBox anzeigen</description></item>
'   <item><term>aclibErrIgnore (2)</term><description>keine Meldung ausgeben</description></item>
'   <item><term>aclibErrFile (4)</term><description>Fehlerinformation in Datei schreiben</description></item>
' </list>
' <remarks>
'   Die Werte {0,1,2} schließen sich gegenseitig aus. Der Werte 4 (aclibErrFile) kann beliebig zu {0,1,2} addiert werden.
'   Beispiel: Init aclibErrRaise + aclibErrFile
' </remarks>
'**/
Public Enum ACLibErrorHandlerMode
   [_aclibErr_default] = -1
   aclibErrRaise = 0&    'Weitergabe an Anwendung
   aclibErrMsgBox = 1&   'MsgBox
   aclibErrIgnore = 2&   'keine Meldung ausgeben
   aclibErrFile = 4&     'Ausgabe in Datei
End Enum

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorResumeMode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verarbeitungsparamter bei aufgetretene Fehler
' </summary>
' <list type="table">
'   <item><term>aclibErrExit (0)</term><description>Abbruch (Funktionsaustritt)</description></item>
'   <item><term>aclibErrResume (1)</term><description>Resume, Problem von außen behoben</description></item>
'   <item><term>aclibErrResumeNext (2)</term><description>Resume next, im Code an nächster Stelle weiterarbeiten</description></item>
' </list>
' <remarks>Wird bei Error-Events genutzt</remarks>
'**/
Public Enum ACLibErrorResumeMode
   aclibErrExit = 0       'Abbruch
   aclibErrResume = 1     'Resume, Problem wurde (von außen) behoben
   aclibErrResumeNext = 2 'Resume next, im Code weiterarbeiten
End Enum

'---------------------------------------------------------------------------------------
' Enum: ACLibErrorNumbers
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ErrorHandler Modes (Fehlerbehandlungsvarianten)
' </summary>
'**/
Public Enum ACLibErrorNumbers
   ERRNR_NOOBJECT = vbObjectError + 1001
   ERRNR_NOCONFIG = vbObjectError + 1002
   ERRNR_INACTIVE = vbObjectError + 1003
   ERRNR_FORBIDDEN = vbObjectError + 9001
End Enum

'Voreinstellungen:
Private Const DEFAULT_ERRORHANDLERMODE As Long = ACLibErrorHandlerMode.[_aclibErr_default]
Private Const DEFAULT_ERRORRESUMEMODE As Long = ACLibErrorResumeMode.aclibErrExit

Private Const ERRORSOURCE_DELIMITERSYMBOL As String = "->"


'Hilfsvariablen
Private m_DefaultErrorHandlerMode As Long 'Zwischenspeicher für Fehlerbehandlungsart
Private m_ErrorHandlerLogFile As String   'Konfiguration des Logfiles

'---------------------------------------------------------------------------------------
' Property: DefaultErrorHandlerMode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Standardverhalten der Fehlerbehandlung
' </summary>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DefaultErrorHandlerMode() As ACLibErrorHandlerMode
On Error Resume Next
    DefaultErrorHandlerMode = m_DefaultErrorHandlerMode
End Property

'---------------------------------------------------------------------------------------
' Property: DefaultErrorHandlerMode
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Standardverhalten der Fehlerbehandlung
' </summary>
' <param name="errMode">ACLibErrorHandlerMode</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let DefaultErrorHandlerMode(ByVal ErrMode As ACLibErrorHandlerMode)
On Error Resume Next
    m_DefaultErrorHandlerMode = ErrMode
End Property

'---------------------------------------------------------------------------------------
' Property: ErrorHandlerLogFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Log file für Fehlermeldungen
' </summary>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ErrorHandlerLogFile() As String
On Error Resume Next
    ErrorHandlerLogFile = m_ErrorHandlerLogFile
End Property

'---------------------------------------------------------------------------------------
' Property: ErrorHandlerLogFile
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Log file für Fehlermeldungen
' </summary>
' <param name="errMode">ACLibErrorHandlerMode</param>
'**/
'---------------------------------------------------------------------------------------
Public Property Let ErrorHandlerLogFile(ByVal Path As String)
On Error Resume Next
'/**
' * @todo Prüfung auf Existenz der Datei oder zumindest des Verzeichnisses
'**/
    m_ErrorHandlerLogFile = Path
End Property

'---------------------------------------------------------------------------------------
' Function: HandleError (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Standard-Prozedur für Fehlerbehandlung
' </summary>
' <param name="lErrorNumber"></param>
' <param name="sSource"></param>
' <param name="sErrDescription"></param>
' <param name="lErrHandlerMode"></param>
' <returns>ACLibErrorResumeMode</returns>
' <remarks>
'Beispiel:
'==<code>
'Private Sub Beispiel() \n
'\n
'On Error GoTo HandleErr \n
'
'[...]
'
'ExitHere:
'On Error Resume Next
'   Exit Sub
'
'HandleErr:
'   Select Case HandleError(Err.Number, "Beispiel", Err.Description)
'   Case ACLibErrorResumeMode.aclibErrResume
'      Resume
'   Case ACLibErrorResumeMode.aclibErrResumeNext
'      Resume Next
'   Case Else
'      Resume ExitHere
'   End Select
'
'End Sub
'<code>==
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function HandleError(ByVal ErrNumber As Long, ByVal ErrSource As String, _
                   Optional ByVal ErrDescription As String, _
                   Optional ByVal ErrHandlerMode As ACLibErrorHandlerMode = DEFAULT_ERRORHANDLERMODE _
            ) As ACLibErrorResumeMode
'hier wäre auch das Aktivieren eine anderen ErrorHandlers möglich (z. B. ErrorHandler-Klasse)

   If ErrHandlerMode = ACLibErrorHandlerMode.[_aclibErr_default] Then
      ErrHandlerMode = m_DefaultErrorHandlerMode
   End If
   
   HandleError = procHandleError(ErrNumber, ErrSource, ErrDescription, ErrHandlerMode)

End Function

Private Function procHandleError(ByRef ErrNumber As Long, ByRef ErrSource As String, _
                                 ByRef ErrDescription As String, _
                                 ByVal ErrHandlerMode As ACLibErrorHandlerMode _
             ) As ACLibErrorResumeMode

   Dim NewErrSource As String
   Dim NewErrDescription As String
   Dim CurrentErrSource As String
   
   NewErrDescription = Err.Description
   CurrentErrSource = Err.Source
   
On Error Resume Next
   
   NewErrSource = ErrSource
   If Len(NewErrSource) = 0 Then
      NewErrSource = CurrentErrSource
   ElseIf CurrentErrSource <> GetApplicationVbProjectName Then
      NewErrSource = NewErrSource & ERRORSOURCE_DELIMITERSYMBOL & CurrentErrSource
   End If
   
   If Len(ErrDescription) > 0 Then
      NewErrDescription = ErrDescription
   End If
   
   'Ausgabe in Datei
   If (ErrHandlerMode And ACLibErrorHandlerMode.aclibErrFile) Then
      printToFile ErrNumber, NewErrSource, NewErrDescription
      ErrHandlerMode = ErrHandlerMode - ACLibErrorHandlerMode.aclibErrFile
   End If

   'Fehlerbehandlung
   Err.Clear
On Error GoTo 0
   Select Case ErrHandlerMode
      Case ACLibErrorHandlerMode.aclibErrRaise 'Weitergabe an Anwendung
         Err.Raise ErrNumber, NewErrSource, NewErrDescription
      Case ACLibErrorHandlerMode.aclibErrMsgBox  'Msgbox
         ShowErrorMessage ErrNumber, NewErrSource, NewErrDescription
      Case ACLibErrorHandlerMode.aclibErrIgnore  'Fehlermeldung übergehen
         '
      Case Else '(sollte eigentlich nie eintreten) .. an Anwendung weitergeben
         Err.Raise ErrNumber, NewErrSource, NewErrDescription
   End Select

   'return resume mode
   procHandleError = DEFAULT_ERRORRESUMEMODE ' Das würde erst bei einer Klasse etwas bringen

End Function

Public Sub ShowErrorMessage(ByVal ErrNumber As Long, ByRef ErrSource As String, ByRef ErrDescription As String)
   
   Dim ErrMsgBoxTitle As String
   Dim Pos As Long
   Dim TempString As String

On Error Resume Next
   
   Const LineBreakPos As Long = 50
   
   Pos = InStr(1, ErrSource, ERRORSOURCE_DELIMITERSYMBOL, vbBinaryCompare)
   If Pos > 1 Then
      ErrMsgBoxTitle = Left$(ErrSource, Pos - 1)
   Else
      ErrMsgBoxTitle = ErrSource
   End If
   
   If Len(ErrSource) > LineBreakPos Then
      Pos = InStr(LineBreakPos, ErrSource, ERRORSOURCE_DELIMITERSYMBOL)
      If Pos > 0 Then
         Do While Pos > 0
            TempString = TempString & Left$(ErrSource, Pos - 1) & vbNewLine
            ErrSource = Mid$(ErrSource, Pos)
            Pos = InStr(LineBreakPos, ErrSource, ERRORSOURCE_DELIMITERSYMBOL)
         Loop
         ErrSource = TempString & ErrSource
      End If
   End If
   
   VBA.MsgBox "Error " & ErrNumber & ": " & vbNewLine & ErrDescription & vbNewLine & vbNewLine & "(" & ErrSource & ")", _
         vbCritical + vbSystemModal + vbMsgBoxSetForeground, ErrMsgBoxTitle

End Sub

Private Sub printToFile(ByRef ErrNumber As Long, ByRef ErrSource As String, _
                        ByRef ErrDescription As String)
    
   Dim FileSource As String
   Dim f As Long
   Dim WriteToFile As Boolean
   Dim PathToErrLogFile As String
   
On Error Resume Next
   
   WriteToFile = True
   
   FileSource = "[" & ErrSource & "]"
   PathToErrLogFile = ErrorHandlerLogFile
   If Len(PathToErrLogFile) = 0 Then
      PathToErrLogFile = CurrentProject.Path & "\Error.log"
   End If
   f = FreeFile
   Open PathToErrLogFile For Append As #f
      Print #f, Format$(Now(), _
            "yyyy-mm-tt hh:nn:ss "); FileSource; _
            " Error "; CStr(ErrNumber); ": "; ErrDescription
   Close #f
   
End Sub

Private Function GetApplicationVbProjectName() As String
   
   Dim VbProjectName As String
   Dim DbFile As String
   Dim vbp As Object
   
On Error Resume Next
   
   VbProjectName = Access.VBE.ActiveVBProject.Name
   DbFile = CurrentDb.Name 'Auf UNCPath verzichtet, damit dieses Modul unabhängig bleibt
   If Access.VBE.ActiveVBProject.FileName <> DbFile Then
      For Each vbp In Access.VBE.VBProjects
         If vbp.FileName = DbFile Then
            VbProjectName = vbp.Name
         End If
      Next
   End If
    
   GetApplicationVbProjectName = VbProjectName
   
End Function
