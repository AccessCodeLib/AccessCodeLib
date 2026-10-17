Attribute VB_Name = "modWinAPI_Registry"
'---------------------------------------------------------------------------------------
' Package: api.winapi.modWinAPI_Registry
'---------------------------------------------------------------------------------------
'
' Access to Windows Registry
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi/modWinAPI_Registry.bas</file>
'  <license>_codelib/license.bas</license>
'  <description>Reads the Windows registry: check whether a key exists, read a value, check whether a COM class is registered</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Public Const HKEY_CLASSES_ROOT   As Long = &H80000000
Public Const HKEY_CURRENT_USER   As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE  As Long = &H80000002

Private Const KEY_QUERY_VALUE As Long = &H1&
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8&
Private Const KEY_NOTIFY As Long = &H10&
Private Const READ_CONTROL As Long = &H20000

'Private Const REG_NONE As Long = 0&
Private Const REG_SZ As Long = 1&
Private Const REG_EXPAND_SZ As Long = 2&
'Private Const REG_BINARY As Long = 3&
Private Const REG_DWORD As Long = 4&
'Private Const REG_DWORD_LITTLE_ENDIAN As Long = 4&
'Private Const REG_DWORD_BIG_ENDIAN As Long = 5&
'Private Const REG_LINK As Long = 6&
'Private Const REG_MULTI_SZ As Long = 7&


Private Const SYNCHRONIZE As Long = &H100000
Private Const KEY_READ As Long = (READ_CONTROL Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) _
                                 And (Not SYNCHRONIZE)
Private Const ERROR_SUCCESS As Long = 0&

Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" ( _
         ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As Long, _
         ByVal samDesired As Long, ByRef phkResult As LongPtr) As Long

Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long

Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" ( _
         ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
         ByRef lpType As Long, ByRef lpData As Any, ByRef lpcbData As Any) As Long


'---------------------------------------------------------------------------------------
' Function: ComClassIsRegistered
'---------------------------------------------------------------------------------------
'
' Checks whether a COM class is registered
'
' Parameters:
'     ClassName - Class identifier
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function ComClassIsRegistered(ByVal ClassName As String) As Boolean
   
   Dim lKey As LongPtr
   
On Error Resume Next

   If RegOpenKeyEx(HKEY_CLASSES_ROOT, ClassName, 0&, KEY_READ, lKey) = ERROR_SUCCESS Then
      If lKey <> 0 Then
         ComClassIsRegistered = True
         Call RegCloseKey(lKey)
      End If
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: RegKeyExist
'---------------------------------------------------------------------------------------
'
' Checks whether a key exists
'
' Parameters:
'     Root - HKEY_CLASSES_ROOT, ...
'     Key  - Key/path
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function RegKeyExist(ByVal Root As Long, ByVal Key As String) As Long
'Pr�fen ob ein Schl�ssel existiert

   Dim Result As Long, hKey As LongPtr

   Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
   If Result = ERROR_SUCCESS Then Call RegCloseKey(hKey)
   RegKeyExist = Result

End Function

'---------------------------------------------------------------------------------------
' Function: RegValueGet
'---------------------------------------------------------------------------------------
'
' Read a value from the registry
'
' Parameters:
'     Root      - HKEY_CLASSES_ROOT, ...
'     Key       - Key/path
'     ValueName - Name of the entry
'     Value     - R�ckgabewert (ByFef)
'
' Returns:
'     Long
'
' Remarks:
'     API: RegOpenKeyEx, RegQueryValueEx, RegCloseKey
'
'---------------------------------------------------------------------------------------
Public Function RegValueGet(ByVal Root As Long, ByVal Key As String, ByVal ValueName As String, ByRef Value As Variant) As Long

   Dim Result As Long, hKey As LongPtr
   Dim dwType As Long, lpData As Long, Buffer As String, lpcdData As Long
   
   'Wert aus einem Feld der Registry auslesen
   Result = RegOpenKeyEx(Root, Key, 0, KEY_READ, hKey)
   If Result = ERROR_SUCCESS Then
      Result = RegQueryValueEx(hKey, ValueName, 0&, dwType, ByVal 0&, lpcdData)
      If Result = ERROR_SUCCESS Then
         Select Case dwType
            Case REG_SZ, REG_EXPAND_SZ
               Buffer = Space$(lpcdData - 1) '(l + 1)
               Result = RegQueryValueEx(hKey, ValueName, 0&, _
                                        dwType, ByVal Buffer, lpcdData)
               If Result = ERROR_SUCCESS Then
                  Value = Buffer
               End If
                  
            Case REG_DWORD
              Result = RegQueryValueEx(hKey, ValueName, 0&, dwType, lpData, lpcdData)
              If Result = ERROR_SUCCESS Then Value = lpData
            
         End Select
      End If
   End If
   
   If Result = ERROR_SUCCESS Then Result = RegCloseKey(hKey)
   RegValueGet = Result
    
End Function
