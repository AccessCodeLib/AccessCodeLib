Attribute VB_Name = "OptionManagerSetup"
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
'  <execute>OptionManagerSetup_SetupTable()</execute>
'  <execute>OptionManagerSetup_CreateHelperModule()</execute>
'  <execute>OptionManagerSetup_CreateConstantandEnum()</execute>
'  <execute>OptionManagerSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit

Private Const m_OptionTableName = "tabOptions"
Private Const m_HelperModuleName = "OptionManagerHelper"
Private Const m_SetupModuleName = "OptionManagerSetup"

Private Function OptionManagerSetup_SetupTable()
    Dim strSQL As String
    strSQL = "Create Table " & m_OptionTableName & " (id AUTOINCREMENT Primary Key, strKey varchar(50), strValue varchar(255))"
    CurrentDb.Execute strSQL
    Application.RefreshDatabaseWindow
End Function

Private Function OptionManagerSetup_CreateHelperModule()
    If IsNull(DLookup("[Name]", "MSysObjects", "[Name] = m_HelperModuleName AND (Type = -32761)")) = False Then Exit Sub

    With Application.VBE.ActiveVBProject.VBComponents
        .Add vbext_ct_StdModule
        .Name = m_HelperModuleName
    End With
    DoCmd.Save acModule, strModulename
    
    Application.RefreshDatabaseWindow
End Function

Private Function OptionManagerSetup_CreateConstantandEnum()
    Dim CODL As Long
    
    With Application.VBE.ActiveVBProject.VBComponents(m_HelperModuleName).CodeModule
        CODL = .CountOfDeclarationLines
        
        .InsertLines CODL + 1, "Public Const OptionManagerhelper_FieldArr = """""
        .InsertLines CODL + 2, ""
        .InsertLines CODL + 3, "Public Enum ltOptions"
        .InsertLines CODL + 4, "    dummy = 0"
        .InsertLines CODL + 5, "End Enum"
    End With
End Function

Private Function OptionManagerSetup_RemoveSelf()
    
    Dim currVbeProject As Object
    Set currVbeProject = Application.VBE.ActiveVBProject

    currVbeProject.VBComponents.Remove currVbeProject.VBComponents(m_SetupModulName)
End Function
