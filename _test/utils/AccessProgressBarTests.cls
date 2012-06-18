VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AccessProgressBarTests"
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


Public Sub AccessProgressBar___Init_And_PerformStep___AutoStart_And_DefaultClearOnLastStep()
    
    Dim Step As Long: Step = 1
    Dim x As Long: x = 100000 '* 5
    Dim y As Long
    
    AccessProgressBar.Init "...working...", x, Step
    
    For y = 1 To x Step Step
        AccessProgressBar.PerformStep
    Next y
    
    Assert.That AccessProgressBar.IsInitialized, Iz.False
    
End Sub

Public Sub AccessProgressBar_PerformStep_ReturnCurrentStep()
    
    Dim Step As Long: Step = 2
    Dim x As Long: x = 10
    Dim y As Long
    
    AccessProgressBar.Init "...working...", x, Step
    
    For y = Step To x Step Step
        Assert.That AccessProgressBar.PerformStep, Iz.EqualTo(y)
    Next y
    
    AccessProgressBar.Clear
    
End Sub

Public Sub AccessProgressBar_Set_Properties_Not_Using_Init()
    
    AccessProgressBar.Clear
    
    AccessProgressBar.Step = 1
    AccessProgressBar.ValueMin = 1
    AccessProgressBar.ValueMax = 10
    AccessProgressBar.Text = "...working..."
    
    Assert.That AccessProgressBar.IsInitialized, Iz.False
    
    AccessProgressBar.Start
    
    Assert.That AccessProgressBar.IsInitialized, Iz.True
    
    Dim y As Long
    For y = AccessProgressBar.Step To AccessProgressBar.ValueMax Step AccessProgressBar.Step
        Assert.That AccessProgressBar.PerformStep(False), Iz.EqualTo(y)
    Next y
    
    Assert.That AccessProgressBar.IsInitialized, Iz.True
    
    AccessProgressBar.Clear
    
    Assert.That AccessProgressBar.IsInitialized, Iz.False
    
End Sub
