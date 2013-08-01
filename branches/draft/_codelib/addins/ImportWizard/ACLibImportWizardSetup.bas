Attribute VB_Name = "ACLibImportWizardSetup"
'---------------------------------------------------------------------------------------
' Modul: ACLibImportWizardSetup
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Setup Modul für den ImportWizard
' </summary>
' <remarks>
' Erstellt die Tabelle USysRegInfo.
' Fuehrt Projekteinstellungen durch.
' Wird nach dem Import wieder entfernt.
' </remarks>
'**/
'<codelib>
'  <file>_codelib/addins/ImportWizard/ACLibImportWizardSetup.bas</file>
'  <use>_codelib/addins/ImportWizard/_config_Application.bas</use>
'  <use>_codelib/addins/ImportWizard/defGlobal_ACLibImportWizard.bas</use>
'  <use>_codelib/addins/ImportWizard/ACLibFileManager.cls</use>
'  <use>data/dao/DaoTools.bas</use>
'  <execute>ACLibImportWizardSetup_SetupUSysRegInfo()</execute>
'  <execute>ACLibImportWizardSetup_SetupTodoInfo()</execute>
'  <execute>ACLibImportWizardSetup_SetupProjectProperties()</execute>
'  <execute>ACLibImportWizardSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const m_SetupModulName As String = "ACLibImportWizardSetup"
Private Const m_VbeProjectName As String = "ACLibImportWizard"
Private Const m_VbeProjectDescription As String = "Access-Add-In für den Import von Dateien aus der Access Code Library (http://access-codelib.net)"
Private Const m_RegInfoTableName As String = "USysRegInfo"
Private Const m_ACLibImportWizardAddinFileName As String = "ACLibImportWizard.mda"
Private Const m_ACLibImportWizardAddinMenuEntry As String = "ACLib Import Wizard"

Public Function ACLibImportWizardSetup_SetupUSysRegInfo()
    
    Dim Query As String
    Dim rs As DAO.Recordset

    If Not DaoTools.TableDefExists(m_RegInfoTableName, CurrentDb()) Then

        Query = "CREATE TABLE " & m_RegInfoTableName & " " & _
                "([Subkey] varchar(255)," & _
                " [Type] long," & _
                " [ValName] varchar(255)," & _
                " [Value] varchar(255))"

        CurrentDb.Execute Query
        
        Application.SetHiddenAttribute acTable, m_RegInfoTableName, True
        
        Set rs = CurrentDb.OpenRecordset(m_RegInfoTableName)
            
            rs.AddNew
            rs!SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_ACLibImportWizardAddinMenuEntry
            rs!Type = 4
            rs!ValName = "BitmapID"
            rs!Value = "339"
            rs.Update
            
            rs.AddNew
            rs!SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_ACLibImportWizardAddinMenuEntry
            rs!Type = 0
            rs.Update
            
            rs.AddNew
            rs!SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_ACLibImportWizardAddinMenuEntry
            rs!Type = 1
            rs!ValName = "Library"
            rs!Value = "|ACCDIR\" & m_ACLibImportWizardAddinFileName
            rs.Update
            
            rs.AddNew
            rs!SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_ACLibImportWizardAddinMenuEntry
            rs!Type = 1
            rs!ValName = "Expression"
            rs!Value = "=StartApplication()"
            rs.Update
            
            rs.AddNew
            rs!SubKey = "HKEY_CURRENT_ACCESS_PROFILE\Menu Add-Ins\" & m_ACLibImportWizardAddinMenuEntry
            rs!Type = 4
            rs!ValName = "Version"
            rs!Value = "1"
            rs.Update
            
        Set rs = Nothing
    End If
        
End Function

Public Function ACLibImportWizardSetup_SetupTodoInfo()
        MsgBox "TODO: Benennen Sie die Datenbank um in: " & m_ACLibImportWizardAddinFileName
End Function

Public Function ACLibImportWizardSetup_SetupProjectProperties()
    defGlobal_ACLibImportWizard.CurrentACLibFileManager.CurrentVbProject.Name = m_VbeProjectName
    defGlobal_ACLibImportWizard.CurrentACLibFileManager.CurrentVbProject.Description = m_VbeProjectDescription
End Function

Public Function ACLibImportWizardSetup_RemoveSelf()
    
    Dim aclibFileMgr As New ACLibFileManager
    Dim currVbeProject As Object
    Set currVbeProject = aclibFileMgr.CurrentVbProject

    currVbeProject.VBComponents.Remove currVbeProject.VBComponents(m_SetupModulName)

End Function
