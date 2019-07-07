Attribute VB_Name = "SqlTools_DotNetLib_Examples"
Option Compare Text
Option Explicit
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>/data/SqlTools_DotNetLib_Examples.bas</file>
'  <use>/data/SqlTools_DotNetLib.bas</use>
'</codelib>
'---------------------------------------------------------------------------------------
'


'--------------------------------------------------------------------
' Beispiele
'--------------------------------------------------------------------
Public Sub Converter_Direct_Use()
    
    Dim SqlAnweisungGenerated As String
    Dim SqlAnweisungExpected As String

    SqlAnweisungGenerated = SqlTools.SqlGenerator(SqlTools.SqlConverters.DaoSqlConverter) _
                  .Select("F1", "F2").SelectField("Count(*)", "", "Anzahl") _
                  .From("Tab1") _
                  .Where("F3", RelationalOperators_Equal + RelationalOperators_GreaterThan, 5) _
                  .GroupBy("F1", "F2") _
                  .HavingString("Count(*) > 1") _
                  .ToString()
                  
    SqlAnweisungExpected = "Select F1, F2, Count(*) As Anzahl From Tab1 Where (F3 >= 5) Group By F1, F2 Having (Count(*) > 1)"
    
    Debug.Print SqlAnweisungGenerated, SqlAnweisungExpected
End Sub

Public Sub Converter_Set_Properties_Later()
' den like-Ausdruck beachten

    Dim Generator As ACLibSqlTools.SqlGenerator

    Set Generator = SqlTools.SqlGenerator()
    With Generator
       .SelectAll
       .From "Tab1"
       .Where "F4", RelationalOperators_Like, "a*"
       .Where "F3", RelationalOperators_GreaterThan, 2
       .Where "F4", RelationalOperators_Like, "a*"
       .OrderBy "F1", "F2", "F3", "F4"
    End With

    'DAO
    Set Generator.Converter = SqlTools.SqlConverters.DaoSqlConverter
    
    Debug.Print Generator.ToString()
    'Jet/Adodb
    Set Generator.Converter = SqlTools.SqlConverters.JetAdodbSqlConverter
    
    Debug.Print Generator.ToString()

End Sub

Public Sub Converter_Separate_Usage()
    ' den like-Ausdruck beachten

    Dim SqlAnweisung As ACLibSqlTools.SqlStatement
    Dim Generator As ACLibSqlTools.SqlGenerator
    Dim Converter As ACLibSqlTools.ISqlConverter

    Set Generator = SqlTools.SqlGenerator()
    With Generator
        .SelectAll
        .From "Tab1"
        .Where "F4", RelationalOperators_Like, "a*"
        .Where "F 3", RelationalOperators_GreaterThan, 2
        .Where "F-4", RelationalOperators_Like, "a*"
        .OrderBy "F1", "F2", "F 3", "F-4"
    End With
    Set SqlAnweisung = Generator.SqlStatement

    'DAO
    Set Converter = SqlTools.SqlConverters.DaoSqlConverter
    Debug.Print Converter.GenerateSqlString(SqlAnweisung)
    
    'Jet/Adodb
    Set Converter = SqlTools.SqlConverters.JetAdodbSqlConverter
    Debug.Print Converter.GenerateSqlString(SqlAnweisung)
    
End Sub


Private Sub Converter_getrennt_einsetzen()
' den like-Ausdruck beachten

   Dim SqlAnweisung As ACLibSqlTools.SqlStatement
   Dim Generator As ACLibSqlTools.SqlGenerator
   Dim Converter As ACLibSqlTools.ISqlConverter

   Set Generator = SqlTools.SqlGenerator()
   With Generator
      .SelectAll
      .From "Tab1"
      .Where "F4", RelationalOperators_Like, "a*"
      .Where "F 3", RelationalOperators_GreaterThan, 2
      .Where "F-4", RelationalOperators_Like, "a*"
      .OrderBy "F1", "F2", "F 3", "F-4"
   End With
   Set SqlAnweisung = Generator.SqlStatement

   'DAO
   Set Converter = SqlTools.SqlConverters.DaoSqlConverter
   Debug.Print Converter.GenerateSqlString(SqlAnweisung)

   'Jet/Adodb
   Set Converter = SqlTools.SqlConverters.JetAdodbSqlConverter
   Debug.Print Converter.GenerateSqlString(SqlAnweisung)

End Sub
