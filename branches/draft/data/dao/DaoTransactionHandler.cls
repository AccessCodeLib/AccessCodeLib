VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DaoTransactionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: DaoTransactionHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Sten Schmidt
' <summary>
' Behandlung von Dao Transaktionen
' </summary>
' <remarks></remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/DaoTransactionHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <test>_test/data/dao/DaoTransactionHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------

Option Compare Database
Option Explicit
 
Private m_TransIsStarted As Boolean
Private m_WorkspaceID As Integer
Private m_ErrorNumber As Long
 
Private Sub Class_Initialize()
    m_WorkspaceID = 0
    m_ErrorNumber = vbObjectError + 20140211
    m_TransIsStarted = False
End Sub
 
Private Sub Class_Terminate()
    Me.Dispose
End Sub
 
Public Sub Dispose()
    If m_TransIsStarted Then Me.Rollback
End Sub
 
Public Property Get TransIsStarted() As Boolean
    TransIsStarted = m_TransIsStarted
End Property
 
Public Property Get WorkspaceID() As Integer
    WorkspaceID = m_WorkspaceID
End Property
 
Public Property Let WorkspaceID(ByVal Value As Integer)
    m_WorkspaceID = Value
End Property
 
Public Property Get ErrorNumber() As Long
    ErrorNumber = m_ErrorNumber
End Property
 
Public Property Let ErrorNumber(ByVal Value As Long)
    m_ErrorNumber = Value
End Property
 
Public Sub BeginTrans()
   
    If m_TransIsStarted Then Err.Raise m_ErrorNumber, "DaoTransactionHandler", "Die Transaktion wurde bereits gestartet."
 
    Application.DBEngine.Workspaces(m_WorkspaceID).BeginTrans
    m_TransIsStarted = True
   
End Sub
 
Public Sub CommitTrans()
 
    If Not m_TransIsStarted Then Err.Raise m_ErrorNumber, "DaoTransactionHandler", "Die Transaktion wurde nicht gestartet."
 
    Application.DBEngine.Workspaces(m_WorkspaceID).CommitTrans
    m_TransIsStarted = False
   
End Sub
 
Public Sub Rollback()
 
    If Not m_TransIsStarted Then Err.Raise m_ErrorNumber, "DaoTransactionHandler", "Die Transaktion wurde nicht gestartet."
 
    Application.DBEngine.Workspaces(m_WorkspaceID).Rollback
    m_TransIsStarted = False
   
End Sub
