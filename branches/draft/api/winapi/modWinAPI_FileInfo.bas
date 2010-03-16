Attribute VB_Name = "modWinAPI_FileInfo"
Attribute VB_Description = "Dateiinformationen mit Win-API auslesen"
'---------------------------------------------------------------------------------------
' Module: modWinAPI_FileInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dateiinformationen mit Win-API auslesen
' </summary>
' <remarks>
' </remarks>
'\ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_FileInfo.bas</file>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------
'
' Code basiert auf http://support.microsoft.com/kb/509493/
'
Option Compare Text
Option Explicit

'/**
'\addtogroup WinAPI
'@{ **/

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
  ByVal lptstrFilename As String, _
  lpdwHandle As Long) As Long


Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
  ByVal lptstrFilename As String, _
  ByVal dwHandle As Long, _
  ByVal dwLen As Long, _
  lpData As Any _
  ) As Long

      
Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" ( _
  pBlock As Any, _
  ByVal lpSubBlock As String, _
  lplpBuffer As Any, _
  puLen As Long _
  ) As Long

Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
  dest As Any, _
  ByVal Source As Long, _
  ByVal Length As Long)

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

Private Function GetVersion(ByVal sPath As String, _
                           ByRef FInfo As FILEINFOOUT) As Boolean

  Dim lRet As Long, lSize As Long, lHandle As Long
  Dim lVerBufLen As Long, lVerPointer As Long
  Dim FileInfo As VS_FIXEDFILEINFO
  Dim sBuffer() As Byte

On Error GoTo HandleErr

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
  
  Call MoveMemory(FileInfo, lVerPointer, Len(FileInfo))
  
  With FileInfo
  
    FInfo.FileVersion = _
      Trim$(Str$((.dwFileVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwFileVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwFileVersionLS And &HFFFF&))
    
    FInfo.ProductVersion = _
      Trim$(Str$((.dwProductVersionMS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionMS And &HFFFF&)) & "." & _
      Trim$(Str$((.dwProductVersionLS And &HFFFF0000) \ &H10000)) & "." & _
      Trim$(Str$(.dwProductVersionLS And &HFFFF&))
      
  End With
  
  GetVersion = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetVersion", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

End Function


'#####################################################
'
' Ergänzung:
'
'---------------------------------------------------------------------------------------
' Function: GetFileVersion
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ermittelt die Version aus einer Datei
' </summary>
' <param name="sFile">vollständiger Pfad zur Datei</param>
' <returns>Versionskennung</returns>
' <remarks>
' Nützlich zum Auslesen von Versionen aus dll-Dateien
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function GetFileVersion(ByVal sFile As String) As String
   
   Dim verInfo As FILEINFOOUT
   
On Error GoTo HandleErr

   If GetVersion(sFile, verInfo) Then
      GetFileVersion = verInfo.FileVersion
   Else
      GetFileVersion = vbNullString
   End If

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case HandleError(Err.Number, "GetFileVersion", Err.Description, ACLibErrorHandlerMode.aclibErrRaise)
   Case ACLibErrorResumeMode.aclibErrResume
      Resume
   Case ACLibErrorResumeMode.aclibErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Function


'/** @} **/ '<-- Ende der Doxygen-Gruppen-Zuordnung
