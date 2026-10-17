Attribute VB_Name = "DotNetLibsSetup"
'---------------------------------------------------------------------------------------
' Package: _dotnetlib.integration.DotNetLibsSetup
'---------------------------------------------------------------------------------------
'
' Setup module: when imported, it automatically creates the table usys_AppFiles,
' imports the DotNetLib DLLs and calls the initialization of the
' DotNetLib integration.
'
' Author:
'     Sten Schmidt
'
' Remarks:
'     No longer required after the import.
'
'---------------------------------------------------------------------------------------

'<codelib>
'  <file>_dotnetlib/integration/DotNetLibsSetup.bas</file>
'  <use>_dotnetlib/integration/DotNetLibs.cls</use>
'  <use>file/LibFiles.cls</use>
'  <use>data/dao/DaoTools.bas</use>
'  <license>_codelib/license.bas</license>
'  <execute>DotNetLibsSetup_Setup_usys_AppFiles("%PrivateRoot%\_dotnetlib\lib")</execute>
'  <execute>DotNetLibsSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const m_SetupModulName As String = "DotNetLibsSetup"
Private Const m_RegInfoTableName As String = "usys_AppFiles"
Private Const m_ExportDestination As String = "[APPDIR]\Lib"

Public Function DotNetLibsSetup_Setup_usys_AppFiles(ByVal dllImportSourcePath As String)
    
    Dim dllFiles As New LibFiles
    
    If Not DaoTools.TableDefExists(m_RegInfoTableName, CurrentDb()) Then
        dllFiles.CreateAppFileTable
    End If
    
    dllFiles.ImportFileToTable "AccessCodeLib.Data.Common.Sql.dll", m_ExportDestination, True, , dllImportSourcePath
    dllFiles.ImportFileToTable "AccessCodeLib.Data.SqlTools.Converter.dll", m_ExportDestination, True, , dllImportSourcePath
    dllFiles.ImportFileToTable "AccessCodeLib.Data.SqlTools.dll", m_ExportDestination, True, , dllImportSourcePath
    dllFiles.ImportFileToTable "AccessCodeLib.Data.SqlTools.interop.dll", m_ExportDestination, True, , dllImportSourcePath
    dllFiles.ImportFileToTable "AccessCodeLib.Data.SqlTools.interop.tlb", m_ExportDestination, True, "ACLibSqlTools", dllImportSourcePath

    dllFiles.ReInitialize
    
End Function

Public Function DotNetLibsSetup_RemoveSelf()
        
    '/*
    ' * TODO: Setup Modul wieder entfernen
    '*/
    
    Debug.Print "Das Modul DotNetLibsSetup kann aus dem Projekt entfernt werden"

End Function
