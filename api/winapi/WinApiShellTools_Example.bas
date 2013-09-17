Attribute VB_Name = "WinApiShellTools_Example"
'---------------------------------------------------------------------------------------
' Class Module: WinApiShellTools_Example
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' \brief        Beispiel zur Verwendung der WinApiShellTools Klasse.
' \ingroup WinAPI
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>api/winapi//WinApiShellTools_Example.bas</file>
'  <use>api/winapi/WinApiShellTools.cls</use>
'  <license>_codelib/license.bas</license>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Public Sub WinApiShellTools_Example()

    Dim WAST As WinApiShellTools
    Set WAST = New WinApiShellTools
        
    'cmd mit der Berechtigung des aktuellen Benutzers ausf�hren.
    WAST.Execute "cmd"
        
    'cmd als Administrator mit erweiterten Benutzerrechten ausf�hren.
    'Bei aktivierter Benutzerkontensteuerung (ab Windows Vista) erscheint der UAC-Dialog
    WAST.ExecuteAsAdmin "cmd"
    
    Set WAST = Nothing

End Sub
