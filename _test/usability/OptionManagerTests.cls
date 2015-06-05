VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "OptionManagerTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'--------------------------------------------------------------------
' AccUnit:TestClass
'--------------------------------------------------------------------
'/**
' @brief Tests für OptionManager
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_test/usability/OptionManagerTests.cls</file>
'  <use>usability/OptionManager.cls</use>
'  <ref><name>SimplyVBUnit</name><major>7</major><minor>0</minor><guid>{AD061B4A-38BF-4847-BA00-0B2F9D60C3FB}</guid></ref>
'  <ref><name>AccUnit_Integration</name><major>0</major><minor>9</minor><guid>{4D92B0E4-E59B-4DD5-8B52-B1AEF82B8941}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

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

Public Sub SettingByName_InsertedKey_OutputKeyValue()
    ' Arrange
   Const KeyName As String = "TestKey1"
    Const TestValue As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    Dim Actual As String
    ' Act
   OptionManager.SettingByName(KeyName) = TestValue
    Actual = OptionManager.SettingByName(KeyName)
    OptionManager.DeleteByName KeyName
    
    ' Assert
   Assert.That Actual, Iz.EqualTo(TestValue)
End Sub

Public Sub Setting_InsertedKeyValueBefore_EnumItemValueNotvbNullString()
    ' Arrange
   Const KeyName As String = "TestKey1b"
    Const TestValue As String = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890"
    Dim Actual As String
    
    ' Act
   OptionManager.SettingByName(KeyName) = TestValue
    Actual = OptionManager.Setting(1)
    OptionManager.DeleteByName KeyName
    
    ' Assert
   Assert.That Actual, Iz.not.EqualTo(vbNullString)
End Sub

Public Sub Setting_ChangedEnumItemValue_EnumItemValueEqualNewValue()
    ' Arrange
   Const newValue As String = "test_4711"
    Dim oldValue As String
    Dim expectedValue As String
    
    ' Act
   oldValue = OptionManager.Setting(1) 'ein Item mit Index >=1 ist nach dem Input oben vorhanden!
   OptionManager.Setting(1) = newValue
    expectedValue = OptionManager.Setting(1)
    OptionManager.Setting(1) = oldValue
    
    ' Assert
   Assert.That newValue, Iz.EqualTo(expectedValue)
End Sub

Public Sub DeleteByName_InsertedKeyValueBefore_KeyDeleted()
    ' Arrange
   Dim getValue As String
    Const KeyName As String = "TestKey1"
    
    ' Act
   OptionManager.SettingByName(KeyName) = vbNullString
    OptionManager.DeleteByName KeyName
    getValue = OptionManager.SettingByName(KeyName)
    
    ' Assert
   Assert.That getValue, Iz.EqualTo(vbNullString)
End Sub

Public Sub KeyName_CreatedEnumKey_EnumKeyExpected()
    ' Arrange
   Const NewKeyName As String = "test_key2"
    Dim ActualKeyName As String
    Dim lastIndex As Long
    
    ' Act
   OptionManager.SettingByName(NewKeyName) = ""
    lastIndex = DCount("*", OptionManager.DataSource)
    ActualKeyName = OptionManager.KeyName(lastIndex)
    OptionManager.DeleteByName NewKeyName
    
    ' Assert
   Assert.That ActualKeyName, Iz.EqualTo(NewKeyName)
End Sub

Public Sub Count_RowNumInsertPlus1_NumberOfRows()
    ' Arrange
   Const NewKeyName As String = "test_key3"
    Dim lastIndex As Long
    Dim expectedValue As Long

    ' Act
   lastIndex = DCount("*", OptionManager.DataSource) + 1
    OptionManager.SettingByName(NewKeyName) = vbNullString
    expectedValue = OptionManager.Count
    OptionManager.DeleteByName NewKeyName
    
    ' Assert
   Assert.That lastIndex, Iz.EqualTo(expectedValue)
End Sub

Public Sub DataSource_Tabname_IsNotvbNullString()
    ' Arrange
   Dim TabName As String
    
    ' Act
   TabName = OptionManager.DataSource
    
    ' Assert
   Assert.That TabName, Iz.not.EqualTo(vbNullString)
End Sub

Public Sub UpdateEnum_InsertKey_CheckKeyInEnumCode()

    Const TestTabName As String = "tabOptions"
    Const TestKeyName As String = "test_UpdateEnum"
    Const EnumDeclarationCode As String = TestKeyName & " = 1"
    
    Dim EnumFound As Boolean
    Dim TestApp As Access.Application
    
    Set TestApp = CreateTestApplication

    With TestApp.CurrentDb
        .Execute "delete from " & TestTabName
        .Execute "insert into " & TestTabName & " (strKey)" & _
                 " values ('" & TestKeyName & "')"
    End With
    
    Dim OptMngr As Object
    Set OptMngr = TestApp.Run("GetOptionManager")

    CallByName OptMngr, "UpdateEnum", VbMethod
    
    With TestApp.VBE.ActiveVBProject.VBComponents("OptionManagerhelper")
        With .CodeModule
            EnumFound = .Find(EnumDeclarationCode, 1, 1, 0, 0)
        End With
    End With
    
    RemoveTestApp TestApp
    
    ' Assert
   Assert.That EnumFound, Iz.True
   
End Sub

Private Sub RemoveTestApp(ByRef TestApp As Access.Application)

    Dim TestAppDb As String
    TestAppDb = TestApp.CurrentDb.Name
    
    TestApp.CloseCurrentDatabase
    DoEvents
    Set TestApp = Nothing
    
    Kill TestAppDb

End Sub

Private Function CreateTestApplication() As Access.Application

    Dim TestApp As Access.Application
    Set TestApp = New Access.Application
    TestApp.Visible = True
    
    TestApp.OpenCurrentDatabase CreateTestDatabase, False
    
    TestApp.CurrentDb.Execute "Create Table tabOptions (strKey varchar(50) Primary Key, strValue varchar(255))"
    
    ImportCodeModulToApp TestApp, "OptionManagerHelper"
    ImportCodeModulToApp TestApp, "OptionManager"
    
    CreateFactoryCodeModul TestApp
    
    Set CreateTestApplication = TestApp

End Function

Private Sub CreateFactoryCodeModul(ByVal TestApp As Access.Application)

    With TestApp.VBE.ActiveVBProject.VBComponents.Add(1) ' vbext_ct_StdModule = 1
        .Name = "modOptionManagerFactory"
        .CodeModule.InsertLines 5, "Public Function GetOptionManager() As OptionManager : Set GetOptionManager = OptionManager : End Function"
    End With
    TestApp.DoCmd.Save acModule, "modOptionManagerFactory"
    TestApp.RunCommand acCmdCompileAndSaveAllModules

End Sub

Private Sub ImportCodeModulToApp(ByVal TestApp As Access.Application, ByVal CodeModulName As String)
    
    Dim FileName As String
    
    FileName = CurrentProject.Path & "\" & CodeModulName & ".tmp"
    
    Application.SaveAsText acModule, CodeModulName, FileName
    TestApp.LoadFromText acModule, CodeModulName, FileName
    
    Kill FileName

End Sub

Private Function CreateTestDatabase() As String
   
   
   Dim TestDbName As String
   
   TestDbName = CurrentProject.Path & "\OptionManagerTestDb" & Mid(CurrentDb.Name, InStrRev(CurrentDb.Name, "."))
   
   If Len(Dir(TestDbName)) > 0 Then
    Kill TestDbName
   End If
   
   Dim db As DAO.Database
   Set db = DBEngine.CreateDatabase(TestDbName, dbLangGeneral)
   db.Close
   
   CreateTestDatabase = TestDbName
   
End Function


Public Sub NochEinTest()
    ' Arrange
   Dim TabName As String
    
    ' Act
   TabName = "abc"
    
    ' Assert
   Assert.That TabName, Iz.not.EqualTo(vbNullString)
End Sub
