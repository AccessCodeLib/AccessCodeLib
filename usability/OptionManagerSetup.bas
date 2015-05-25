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
'  <execute>OptionManagerSetup_CreateEnum()</execute>
'  <execute>OptionManagerSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const m_OptionTableName = "tabOptions"
Private Const m_HelperModuleName = "OptionManagerHelper"
Private Const m_SetupModuleName = "OptionManagerSetup"

Public Function OptionManagerSetup_SetupTable()
    Dim strSQL As String
    strSQL = "Create Table " & m_OptionTableName & " (id AUTOINCREMENT Primary Key, strKey varchar(50), strValue varchar(255))"
    CurrentDb.Execute strSQL
    Application.RefreshDatabaseWindow
End Function

Public Function OptionManagerSetup_CreateHelperModule()
    If IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_HelperModuleName & "' AND (Type = -32761)")) = False Then Exit Function

    Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_StdModule)
    DoCmd.Save acModule, m_HelperModuleName

    Application.RefreshDatabaseWindow
End Function

Public Function OptionManagerSetup_CreateEnum()
    Dim CODL As Long

    With Application.VBE.ActiveVBProject.VBComponents(m_HelperModuleName).CodeModule
        CODL = .CountOfDeclarationLines

        .InsertLines CODL + 1, ""
        .InsertLines CODL + 2, "Public Enum ltOptions"
        .InsertLines CODL + 3, "    dummy = 0"
        .InsertLines CODL + 4, "End Enum"
    End With
    DoCmd.Save acModule, m_HelperModuleName
End Function

Public Function OptionManagerSetup_RemoveSelf()

    Dim currVbeProject As Object
    Set currVbeProject = Application.VBE.ActiveVBProject

    currVbeProject.VBComponents.Remove currVbeProject.VBComponents(m_SetupModuleName)
End Function
