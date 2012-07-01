VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlTools_DotNetLibTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Text
Option Explicit

'<codelib>
'  <file>_test/data/SqlTools_DotNetLibTests.cls</file>
'  <license>_codelib/license.bas</license>
'</codelib>

'AccUnit:TestClass

' AccUnit infrastructure for advanced AccUnit features. Do not remove these lines.
Implements SimplyVBUnit.ITestFixture
Implements AccUnit_Integration.ITestManagerBridge
Private TestManager As AccUnit_Integration.TestManager
Private Sub ITestManagerBridge_InitTestManager(ByVal NewTestManager As AccUnit_Integration.ITestManagerComInterface): Set TestManager = NewTestManager: End Sub
Private Function ITestManagerBridge_GetTestManager() As AccUnit_Integration.ITestManagerComInterface: Set ITestManagerBridge_GetTestManager = TestManager: End Function
Private Sub ITestFixture_AddTestCases(ByVal Tests As SimplyVBUnit.TestCaseCollector): TestManager.AddTestCases Tests: End Sub

'--------------------------------------------------------------------
' Tests
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
    
    Assert.That SqlAnweisungGenerated, Iz.EqualTo(SqlAnweisungExpected)
End Sub

Public Sub Converter_Set_Properties_Later()
' den like-Ausdruck beachten

    Dim ExpectedDAOQuery As String
        ExpectedDAOQuery = "Select * From Tab1 Where (F4 Like 'a*') And (F3 > 2) And (F4 Like 'a*') Order By F1, F2, F3, F4"
   
    Dim ExpectedJetADOQuery As String
        ExpectedJetADOQuery = "Select * From Tab1 Where (F4 Like 'a%') And (F3 > 2) And (F4 Like 'a%') Order By F1, F2, F3, F4"

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
    
    Assert.That Generator.ToString(), Iz.EqualTo(ExpectedDAOQuery)

    'Jet/Adodb
    Set Generator.Converter = SqlTools.SqlConverters.JetAdodbSqlConverter
    
    Assert.That Generator.ToString(), Iz.EqualTo(ExpectedJetADOQuery)

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
    Dim DAOQuery As String
        DAOQuery = "Select * From Tab1 Where (F4 Like 'a*') And ([F 3] > 2) And ([F-4] Like 'a*') Order By F1, F2, [F 3], [F-4]"
        
    Assert.That Converter.GenerateSqlString(SqlAnweisung), Iz.EqualTo(DAOQuery)

    'Jet/Adodb
    Set Converter = SqlTools.SqlConverters.JetAdodbSqlConverter
    Dim JetADOQuery As String
        JetADOQuery = "Select * From Tab1 Where (F4 Like 'a%') And ([F 3] > 2) And ([F-4] Like 'a%') Order By F1, F2, [F 3], [F-4]"
    
    Assert.That Converter.GenerateSqlString(SqlAnweisung), Iz.EqualTo(JetADOQuery)
    
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
