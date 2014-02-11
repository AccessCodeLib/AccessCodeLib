VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoTransactionHandlerTests"
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

Public Sub WorkspaceID_Default_IsZero()

    With New DaoTransactionHandler
        Assert.AreEqual 0, .WorkspaceID
        .BeginTrans
        .CommitTrans
    End With

End Sub

Public Sub Rollback_RaiseError_IfTransIsNotStarted()
On Error GoTo ERR_
    
    With New DaoTransactionHandler
        .ErrorNumber = 999
        .Rollback
    End With
    
ERR_:
    Assert.AreEqual 999, Err.Number
End Sub

Public Sub CommitTrans_RaiseError_IfTransIsNotStarted()
On Error GoTo ERR_
    
    With New DaoTransactionHandler
        .ErrorNumber = 999
        .CommitTrans
    End With
    
ERR_:
    Assert.AreEqual 999, Err.Number
End Sub
