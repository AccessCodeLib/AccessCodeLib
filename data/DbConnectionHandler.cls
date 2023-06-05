VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: data.DbConnectionHandler
'---------------------------------------------------------------------------------------
' Database connection via ADODB or DAO (incl. ODBC)
'
' Standardises ADO and DAO methods
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoHandler.cls</use>
'  <use>data/ado/AdodbHandler.cls</use>
'  <use>data/odbc/OdbcHandler.cls</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'Objekte für Datenzugriff
Private WithEvents m_DaoHandler As DaoHandler
Attribute m_DaoHandler.VB_VarHelpID = -1
Private WithEvents m_AdodbHandler As AdodbHandler
Attribute m_AdodbHandler.VB_VarHelpID = -1
Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1

' Group: Events
' Error events for possible error handling via the frontend application

'---------------------------------------------------------------------------------------
' Event: ErrorMissingCurrentDb
'---------------------------------------------------------------------------------------
' Event for missing CurrentDb reference
' (Passing the <data.dao.DaoHandler> or <data.odbc.OdbcHandler> event: ErrorMissingCurrentDb)
'
' Parameters:
'     Msg   -  Message about the cause of the event triggering. (Not yet used)
'     NewCurrentDbRef - DAO.Database reference to be used for CurrentDb
'
' Remarks:
'     The required DAO.Database reference can be passed on via the ByRef parameter NewCurrentDbRef.
'
'---------------------------------------------------------------------------------------
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

'---------------------------------------------------------------------------------------
' Event: ErrorMissingCurrentConnection
'---------------------------------------------------------------------------------------
'
' Event for missing ADODB connection reference
' (Passing the AdodbHandler event: <data.adodb.AdodbHandler::ErrorMissingCurrentConnection>)
'
' Parameters:
'     Msg   -  Message about the cause of the event triggering. (Not yet used)
'     NewCurrentConnection - ADODB.Connectionreference to be used for CurrentConnection
'
' Remarks:
'     The appropriate ADODB.Database reference can be passed on via the ByRef parameter NewCurrentConnection.
'     However, passing on a reference is not mandatory, since, it is stored in the AdodbHandler property <data.adodb.AdodbHandler::CurrentConnection>.
'     The connection is then created from the <data.adodb.AdodbHandler::ConnectionString>
'
'---------------------------------------------------------------------------------------
Public Event ErrorMissingCurrentConnection(ByVal Msg As String, ByRef NewCurrentConnection As ADODB.Connection)

'---------------------------------------------------------------------------------------
' Event: ErrorEmptyOledbConnectionString
'---------------------------------------------------------------------------------------
'
' Event for missing OLEDB connection string
' (Passing the AdodbHandler event: <data.adodb.AdodbHandler::ErrorEmptyOledbConnectionString> )
'
' Parameters:
'     Msg                  - Message about the cause of the event triggering. (Not yet used)
'     NewConnectionString  - OLEDB Connectionstring
'
' Remarks:
'     The appropriate OLEDB connection string can be passed on via the ByRef parameter NewConnectionString.
'
'---------------------------------------------------------------------------------------
Public Event ErrorEmptyOledbConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'---------------------------------------------------------------------------------------
' Event: ErrorEmptyOdbcConnectionString
'---------------------------------------------------------------------------------------
'
' Event for missing ODBC connection string
' (Passing the OdbcHandler event: <data.odbc.OdbcHandler::ErrorEmptyConnectionString>)
'
' Parameters:
'     Msg                  - Message about the cause of the event triggering. (Not yet used)
'     NewConnectionString  - ODBC-Connectionstring
'
' Remarks:
'     The appropriate ODBC connection string can be passed on via the ByRef parameter NewConnectionString.
'
'---------------------------------------------------------------------------------------
Public Event ErrorEmptyOdbcConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'---------------------------------------------------------------------------------------
' Event: Disposed
'---------------------------------------------------------------------------------------
'
' Event message when object references have been removed
'
'---------------------------------------------------------------------------------------
Public Event Disposed()


' Group: Dispose support

Private m_Disposed As Boolean

'
' Init / Terminate
' ----------------
'
'Private Sub Class_Initialize()
''
'End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

'---------------------------------------------------------------------------------------
' Sub: Dispose
'---------------------------------------------------------------------------------------
'
' Remove object references
'
'---------------------------------------------------------------------------------------
Public Sub Dispose()

   If m_Disposed Then Exit Sub

On Error Resume Next

   ClearObjectReferences
   
   m_Disposed = True
   RaiseEvent Disposed
   
End Sub

Private Sub ClearObjectReferences()

On Error Resume Next

   If Not m_AdodbHandler Is Nothing Then
      m_AdodbHandler.Dispose
      Set m_AdodbHandler = Nothing
   End If
   
   If Not m_DaoHandler Is Nothing Then
      m_DaoHandler.Dispose
      Set m_DaoHandler = Nothing
   End If
   
   If Not m_OdbcHandler Is Nothing Then
      m_OdbcHandler.Dispose
      Set m_OdbcHandler = Nothing
   End If
   
End Sub


'#######################################################################################
' Group: DAO data connection
' DAO access methods
' --------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Property: DAO
'---------------------------------------------------------------------------------------
'
' Instance of <data.dao.DaoHandler>
'
' Returns:
'     <data.dao.DaoHandler>
'
'---------------------------------------------------------------------------------------
Public Property Get DAO() As DaoHandler
   If m_DaoHandler Is Nothing Then
      Set m_DaoHandler = New DaoHandler
   End If
   Set DAO = m_DaoHandler
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'
' CurrentDb = <data.dao.DaoHandler::CurrentDb>
'
' Returns:
'     DAO.Database
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
   Set CurrentDb = Me.DAO.CurrentDb
End Property

Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
   Set Me.DAO.CurrentDb = NewCurrentDb
End Property


'#######################################################################################
' Group: ADODB data connection
' ADODB access methods
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Property: ADODB
'---------------------------------------------------------------------------------------
'
' Instance of <data.adodb.AdodbHandler>
'
' Returns:
'     <data.adodb.AdodbHandler>
'
'---------------------------------------------------------------------------------------
Public Property Get ADODB() As AdodbHandler
   If m_AdodbHandler Is Nothing Then
      Set m_AdodbHandler = New AdodbHandler
   End If
   Set ADODB = m_AdodbHandler
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentConnection
'---------------------------------------------------------------------------------------
'
' ADODB connection to the current DBMS connection = <data.adodb.AdodbHandler::CurrentConnection>
'
' Returns:
'     ADODB.Connection
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnection() As ADODB.Connection
   Set CurrentConnection = Me.ADODB.CurrentConnection
End Property

Public Property Set CurrentConnection(ByRef NewConnection As ADODB.Connection)
   Set Me.ADODB.CurrentConnection = NewConnection
End Property

'---------------------------------------------------------------------------------------
' Property: OledbConnectionString
'---------------------------------------------------------------------------------------
'
' OledbConnectionString = <data.adodb.AdodbHandler::ConnectionString>
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Friend Property Get OledbConnectionString() As String
   OledbConnectionString = Me.ADODB.ConnectionString
End Property

Public Property Let OledbConnectionString(ByVal NewConnectionString As String)
   Me.ADODB.ConnectionString = NewConnectionString
End Property


'#######################################################################################
' Group: ODBC data connection
' ODBC access methods
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
' Property: ODBC
'---------------------------------------------------------------------------------------
'
' Instance of <data.odbc.OdbcHandler>
'
' Returns:
'     <data.odbc.OdbcHandler>
'
'---------------------------------------------------------------------------------------
Public Property Get ODBC() As OdbcHandler
   If m_OdbcHandler Is Nothing Then
     Set m_OdbcHandler = New OdbcHandler
   End If
   Set ODBC = m_OdbcHandler
End Property

'---------------------------------------------------------------------------------------
' Property: OdbcConnectionString
'---------------------------------------------------------------------------------------
'
' OdbcConnectionString = <data.odbc.OdbcHandler::ConnectionString>
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Friend Property Get OdbcConnectionString() As String
   OdbcConnectionString = Me.ODBC.ConnectionString
End Property

Public Property Let OdbcConnectionString(ByVal NewConnectionString As String)
   Me.ODBC.ConnectionString = NewConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'
' CurrentDbBE = <data.odbc.OdbcHandler::CurrentDbBE>
'
' Returns:
'     DAO.Database of ODBC database
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentDbBE() As DAO.Database
   Set CurrentDbBE = Me.ODBC.CurrentDbBE
End Property

Public Property Set CurrentDbBE(ByRef NewDb As DAO.Database)
   If Not (NewDb Is Nothing) Then
      Set Me.ODBC.CurrentDbBE = NewDb
   End If
End Property


'#######################################################################################
'
' Data access events
' -------------------

' AdodbHandler events
' --------------------
Private Sub m_AdodbHandler_ErrorMissingCurrentConnection(ByVal Msg As String, ByRef NewCurrentConnection As ADODB.Connection)
   RaiseEvent ErrorMissingCurrentConnection(Msg, NewCurrentConnection)
End Sub

Private Sub m_AdodbHandler_ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)
   RaiseEvent ErrorEmptyOledbConnectionString(Msg, NewConnectionString)
End Sub

' DaoHandler events
' ------------------
Private Sub m_DaoHandler_ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)
   RaiseEvent ErrorMissingCurrentDb(Msg, NewCurrentDbRef)
End Sub

' OdbcHandler events
' ------------------
Private Sub m_OdbcHandler_ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)
   RaiseEvent ErrorEmptyOdbcConnectionString(Msg, NewConnectionString)
End Sub

Private Sub m_OdbcHandler_ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

   '1. Try to pass from DAO class:
   Set NewCurrentDbRef = Me.DAO.CurrentDb
   
   '2. Trigger event if Me.DAO.CurrentDb also returns nothing
   If NewCurrentDbRef Is Nothing Then
      RaiseEvent ErrorMissingCurrentDb(Msg, NewCurrentDbRef)
      If Not NewCurrentDbRef Is Nothing Then
         Set Me.DAO.CurrentDb = NewCurrentDbRef
      End If
   End If
   
End Sub
