Attribute VB_Name = "OptionManagerSetup"
'---------------------------------------------------------------------------------------
' Modul: OptionManagerSetup
'---------------------------------------------------------------------------------------
'/**
' \author       Andreas Vogt
' <summary>
' Setup Modul für den OptionManager
' </summary>
' <remarks>
' Erstellt die Tabelle tabOptions.
' Erstellt ein Hilfsmodul mit Enum und Const-Var
' Wird nach dem Import wieder entfernt.
' </remarks>
'**/
'<codelib>
'  <file>usability/OptionManagerSetup.bas</file>
'  <use>usability/OptionManager.cls</use>
'  <ref><name>VBIDE</name><major>5</major><minor>3</minor><guid>{0002E157-0000-0000-C000-000000000046}</guid></ref>
'  <execute>OptionManagerSetup_SetupTable()</execute>
'  <execute>OptionManagerSetup_CreateHelperModule()</execute>
'  <execute>OptionManagerSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const m_OptionTableName = "tabOptions"
Private Const m_HelperModuleName = "OptionManagerHelper"
Private Const m_SetupModuleName = "OptionManagerSetup"
Private Const m_EnumName = "SettingKeys"

Public Function OptionManagerSetup_SetupTable()

    Dim strSQL As String
    
    If Not IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_OptionTableName & "' AND (Type IN (1, 4, 5))")) Then Exit Function
    
    With CurrentDb
        strSQL = "Create Table " & m_OptionTableName & " (strKey varchar(50) Primary Key, strValue varchar(255))"
        .Execute strSQL
        .Execute "Insert into " & m_OptionTableName & " (strKey, strValue) Values ('Name', 'Vogt')"
        .Execute "Insert into " & m_OptionTableName & " (strKey, strValue) Values ('Vorname', 'Andreas')"
        .Execute "Insert into " & m_OptionTableName & " (strKey, strValue) Values ('Beruf','Software-Entwickler')"
    End With
    Application.RefreshDatabaseWindow

End Function

Public Function OptionManagerSetup_CreateHelperModule()

    If Not IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_HelperModuleName & "' AND (Type = -32761)")) Then Exit Function

    With Application.VBE.ActiveVBProject.VBComponents
        With .Add(vbext_ct_StdModule)
           .Name = m_HelperModuleName
           With .CodeModule
                .DeleteLines 1, .CountOfDeclarationLines
           
                .InsertLines 1, "Option Compare Database" & vbNewLine & _
                                "Option Explicit" & vbNewLine & _
                                vbNewLine & _
                                "Public Const OptionManagerDefaultDataSource As String = """ & m_OptionTableName & """" & vbNewLine & _
                                vbNewLine & _
                                "Public Enum " & m_EnumName & vbNewLine & _
                                "    [_undefined] = 0" & vbNewLine & _
                                "End Enum"
           
           End With
        End With
    End With
    
    DoCmd.RunCommand acCmdCompileAndSaveAllModules
    Application.RefreshDatabaseWindow
    DoCmd.Close acModule, m_HelperModuleName
    
End Function

Public Function OptionManagerSetup_RemoveSelf()

    Dim currVbeProject As Object
    Set currVbeProject = Application.VBE.ActiveVBProject

    currVbeProject.VBComponents.Remove currVbeProject.VBComponents(m_SetupModuleName)

End Function
