Attribute VB_Name = "Package_FilterStringBuilder"
'---------------------------------------------------------------------------------------
' Modul: Package_FilterStringBuilder
'---------------------------------------------------------------------------------------
'
'<codelib>
'  <file>_examples/Package_FilterStringBuilder_Setup.bas</file>
'  <description>Import FilterStringBuild and example code modules</description>
'  <use>data/FilterStringBuilder.cls</use>
'  <use>data/FilterStringBuilder_Examples.bas</use>
'  <execute>FilterStringBuilderSetup_CreateExampleModule()</execute>
'  <execute>FilterStringBuilderSetup_RemoveSelf()</execute>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Text
Option Explicit

Private Const m_SetupModuleName = "Package_FilterStringBuilder"
Private Const m_ExampleModuleName = "modFilterStringBuilderExamples"

Public Function FilterStringBuilderSetup_CreateExampleModule()

    If Not IsNull(DLookup("[Name]", "MSysObjects", "[Name] = '" & m_ExampleModuleName & "' AND (Type = -32761)")) Then Exit Function

    With Application.VBE.ActiveVBProject.VBComponents
        With .Add(1) ' 1 = vbext_ct_StdModule
           .Name = m_ExampleModuleName
           With .CodeModule
                .DeleteLines 1, .CountOfDeclarationLines
           
                .InsertLines 1, "Option Compare Text" & vbNewLine & _
                                "Option Explicit" & vbNewLine & _
                                vbNewLine & _
                                "Private Sub TestFilter()" & vbNewLine & _
                                vbNewLine & _
                                "    With New FilterStringBuilder" & vbNewLine & _
								"        .SelectSqlDialect SQL_DAO" & vbNewLine & _
								"        .Add ""DateField"", SQL_Date, SQL_GreaterThan + SQL_Equal, Date" & vbNewLine & _
								"        Debug.Print .ToString" & vbNewLine & _
								"    End With" & vbNewLine & _
                                vbNewLine & _
                                "End Sub"
           
           End With
        End With
    End With
    
End Function

Public Function FilterStringBuilderSetup_RemoveSelf()

    Dim currVbeProject As Object
    Set currVbeProject = Application.VBE.ActiveVBProject

    currVbeProject.VBComponents.Remove currVbeProject.VBComponents(m_SetupModuleName)

End Function
