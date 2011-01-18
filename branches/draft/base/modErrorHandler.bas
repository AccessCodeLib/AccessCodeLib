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
Private Const m_conDefaultErrorHandlerMode As Long = ACLibErrorHandlerMode.[_aclibErr_default]
Private Const m_conDefaultErrorResumeMode As Long = ACLibErrorResumeMode.aclibErrExit

Private Const m_ErrorSourceDelimiterSymbol As String = "->"


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
Public Property Let DefaultErrorHandlerMode(ByVal errMode As ACLibErrorHandlerMode)
On Error Resume Next
    m_DefaultErrorHandlerMode = errMode
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
Public Function HandleError(ByVal lErrorNumber As Long, ByVal sSource As String, _
                   Optional ByVal sErrDescription As String, _
                   Optional ByVal lErrHandlerMode As ACLibErrorHandlerMode = m_conDefaultErrorHandlerMode _
            ) As ACLibErrorResumeMode
'hier wäre auch das Aktivieren eine anderen ErrorHandlers möglich (z. B. ErrorHandler-Klasse)

   If lErrHandlerMode = ACLibErrorHandlerMode.[_aclibErr_default] Then
      lErrHandlerMode = m_DefaultErrorHandlerMode
   End If
   
   HandleError = procHandleError(lErrorNumber, sSource, sErrDescription, lErrHandlerMode)

End Function

Private Function procHandleError(ByRef lErrorNumber As Long, ByRef sSource As String, _
                                 ByRef sErrDescription As String, _
                                 ByVal lErrHandlerMode As ACLibErrorHandlerMode _
             ) As ACLibErrorResumeMode

   Dim strSource As String
   Dim strErrDescription As String
   Dim strErrSource As String
   
   strErrDescription = Err.Description
   strErrSource = Err.Source
   
On Error Resume Next
   
   strSource = sSource
   If Len(strSource) = 0 Then
      strSource = strErrSource
   ElseIf strErrSource <> getApplicationVbProjectName Then
      strSource = strSource & m_ErrorSourceDelimiterSymbol & strErrSource
   End If
   
   If Len(sErrDescription) > 0 Then
      strErrDescription = sErrDescription
   End If
   
   'Ausgabe in Datei
   If (lErrHandlerMode And ACLibErrorHandlerMode.aclibErrFile) Then
      printToFile lErrorNumber, strSource, strErrDescription
      lErrHandlerMode = lErrHandlerMode - ACLibErrorHandlerMode.aclibErrFile
   End If

   'Fehlerbehandlung
   Err.Clear
On Error GoTo 0
   Select Case lErrHandlerMode
      Case ACLibErrorHandlerMode.aclibErrRaise 'Weitergabe an Anwendung
         Err.Raise lErrorNumber, strSource, strErrDescription
      Case ACLibErrorHandlerMode.aclibErrMsgBox  'Msgbox
         ShowErrorMessage lErrorNumber, strSource, strErrDescription
      Case ACLibErrorHandlerMode.aclibErrIgnore  'Fehlermeldung übergehen
         '
      Case Else '(sollte eigentlich nie eintreten) .. an Anwendung weitergeben
         Err.Raise lErrorNumber, strSource, strErrDescription
   End Select

   'return resume mode
   procHandleError = m_conDefaultErrorResumeMode ' Das würde erst bei einer Klasse etwas bringen

End Function

Public Sub ShowErrorMessage(ByVal lErrorNumber As Long, ByRef sSource As String, ByRef sErrorDescription As String)
   
   Dim strMsgBoxTitle As String
   Dim Pos As Long
   Dim TempString As String

On Error Resume Next
   
   Const conLineBreakPos As Long = 50
   
   Pos = InStr(1, sSource, m_ErrorSourceDelimiterSymbol, vbBinaryCompare)
   If Pos > 1 Then
      strMsgBoxTitle = Left$(sSource, Pos - 1)
   Else
      strMsgBoxTitle = sSource
   End If
   
   If Len(sSource) > conLineBreakPos Then
      Pos = InStr(conLineBreakPos, sSource, m_ErrorSourceDelimiterSymbol)
      If Pos > 0 Then
         Do While Pos > 0
            TempString = TempString & Left$(sSource, Pos - 1) & vbNewLine
            sSource = Mid$(sSource, Pos)
            Pos = InStr(conLineBreakPos, sSource, m_ErrorSourceDelimiterSymbol)
         Loop
         sSource = TempString & sSource
      End If
   End If
   
   VBA.MsgBox "Error " & lErrorNumber & ": " & vbNewLine & sErrorDescription & vbNewLine & vbNewLine & "(" & sSource & ")", _
         vbCritical + vbSystemModal + vbMsgBoxSetForeground, strMsgBoxTitle

End Sub

Private Sub printToFile(ByRef lErrorNumber As Long, ByRef sSource As String, _
                        ByRef sErrDescription As String)
    
   Dim strFileSource As String
   Dim iFile As Long
   Dim bolWriteToFile As Boolean
   Dim PathToErrLogFile As String
   
On Error Resume Next
   
   bolWriteToFile = True
   
   strFileSource = "[" & sSource & "]"
   PathToErrLogFile = ErrorHandlerLogFile
   If Len(PathToErrLogFile) = 0 Then
      PathToErrLogFile = CurrentProject.Path & "\Error.log"
   End If
   iFile = FreeFile
   Open PathToErrLogFile For Append As #iFile
      Print #iFile, Format$(Now(), _
            "yyyy-mm-tt hh:nn:ss "); strFileSource; _
            " Error "; CStr(lErrorNumber); ": "; sErrDescription
   Close #iFile
   
End Sub

Private Function getApplicationVbProjectName() As String
   
   Dim strVbProjectName As String
   Dim strDbFile As String
   Dim vbp As Object
   
On Error Resume Next
   
   strVbProjectName = Access.VBE.ActiveVBProject.Name
   strDbFile = CurrentDb.Name 'Auf UNCPath verzichtet, damit dieses Modul unabhängig bleibt
   If Access.VBE.ActiveVBProject.FileName <> strDbFile Then
      For Each vbp In Access.VBE.VBProjects
         If vbp.FileName = strDbFile Then
            strVbProjectName = vbp.Name
         End If
      Next
   End If
    
   getApplicationVbProjectName = strVbProjectName
   
End Function
