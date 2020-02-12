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
Option Private Module

Private Declare PtrSafe Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" ( _
  ByVal lptstrFilename As String, _
  lpdwHandle As LongPtr) As Long

Private Declare PtrSafe Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" ( _
  ByVal lptstrFilename As String, _
  ByVal dwHandle As LongPtr, _
  ByVal dwLen As Long, _
  lpData As Any _
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
   Dim VerInfo As FILEINFOOUT
   If GetVersion(sFile, VerInfo) Then
      GetFileVersion = VerInfo.FileVersion
   Else
      GetFileVersion = vbNullString
   End If
End Function
