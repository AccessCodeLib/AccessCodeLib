VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UnicodeToolsTests"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'/*
' * TODO: More and better Tests
' */

Public Sub IsUnicode_UnicodeString_UNICODE_MASK()
    
    'Der erste Testteil schlägt fehlt, wenn die Testklasse mit dem ImportWizard
    'importiert wird, da der String "test" dann nicht als Unicode vorliegt.
    'Wird "test" im VBE-Editor neu eingegeben, läuft der Test erfolgreich durch.
    
    'Dim t1 As String
    '    t1 = StrConv("test", vbUnicode)
    '
    'Assert.That UnicodeTools.IsUnicode(t1, UNICODE_MASK), Iz.True
    
    Dim t2 As String
        t2 = "test"
    
    Assert.That UnicodeTools.IsUnicode(t2, UNICODE_MASK), Iz.False
End Sub

Public Sub IsUnicode_UnicodeString_NOT_ASCII_MASK()
    
    Dim t1 As String
        t1 = StrConv("test", vbUnicode)
        
    Assert.That UnicodeTools.IsUnicode(t1, NOT_ASCII_MASK), Iz.True
    
    Dim t2 As String
        t2 = "test"
    
    Assert.That UnicodeTools.IsUnicode(t2, NOT_ASCII_MASK), Iz.True
End Sub
