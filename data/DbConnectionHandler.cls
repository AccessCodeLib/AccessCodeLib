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
' Klassenmodul: DbConnectionHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datenbankverbindung per ADO oder DAO (inkl. ODBC)
' </summary>
' <remarks>
' Vereinheitlicht ADO- und DAO-Methoden
' </remarks>
'**/
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

'Objekte f?r Datenzugriff
Private WithEvents m_DaoHandler As DaoHandler
Attribute m_DaoHandler.VB_VarHelpID = -1
Private WithEvents m_AdodbHandler As AdodbHandler
Attribute m_AdodbHandler.VB_VarHelpID = -1
Private WithEvents m_OdbcHandler As OdbcHandler
Attribute m_OdbcHandler.VB_VarHelpID = -1

'Error-Events f?r m?gliche Fehlerbehandlung ?ber die FE-Anwendung

'/**
' <summary>
' Ereignis f?r fehlende CurrentDb-Referenz
' (Weitergabe des <see cref=DaoHandler#ErrorMissingCurrentDb>DaoHandler</see>- bzw. <see cref=OdbcHandler#ErrorMissingCurrentDb>OdbcHandler</see>-Ereignisses: ErrorMissingCurrentDb)
' </summary>
' <param name="Msg">Meldung ?ber Ursache der Ereignisausl?sung. (Wird noch nicht genutzt)</param>
' <param name="NewCurrentDbRef">DAO.Database-Referenz, die f?r CurrentDb verwendet werden soll</param>
' <remarks>?ber den ByRef-Parameter NewCurrentDbRef kann die passende DAO.Database-Referenz weitergegeben werden.</remarks>
'**/
Public Event ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

'/**
' <summary>
' Ereignis f?r fehlende ADODB-Connection-Referenz
' (Weitergabe des AdodbHandler-Ereignisses: <see cref=AdodbHandler#ErrorMissingCurrentConnection>ErrorMissingCurrentConnection</see>)
' </summary>
' <param name="Msg">Meldung ?ber Ursache der Ereignisausl?sung. (Wird noch nicht genutzt)</param>
' <param name="NewCurrentConnection">ADODB.Connection-Referenz, die f?r CurrentConnection verwendet werden soll</param>
' <remarks>
' ?ber den ByRef-Parameter NewCurrentConnection kann die passende ADODB.Database-Referenz weitergegeben werden.
' Eine Referenz-Weitergabe ist allerdings nicht zwingend erforderlich,
' da z. B. in der ADODBHandler-Methode <see cref=AdodbHandler#CurrentConnection>CurrentConnection</see>
' die Connection dann aus dem <see cref=AdodbHandler#ConnectionString>ConnectionString</see> erstellt wird.
' </remarks>
'**/
Public Event ErrorMissingCurrentConnection(ByVal Msg As String, ByRef NewCurrentConnection As ADODB.Connection)

'/**
' <summary>
' Ereignis f?r fehlenden OLEDB-Connectionstring
' (Weitergabe des AdodbHandler-Ereignisses: <see cref=AdodbHandler#ErrorEmptyOledbConnectionString>ErrorEmptyOledbConnectionString</see>)
' </summary>
' <param name="Msg">Meldung ?ber Ursache der Ereignisausl?sung. (Wird noch nicht genutzt)</param>
' <param name="NewConnectionString">OLEDB-Connectionstring</param>
' <remarks>?ber den ByRef-Parameter NewConnectionString kann der passende OLEDB-Connectionstring weitergegeben werden.</remarks>
'**/
Public Event ErrorEmptyOledbConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'/**
' <summary>
' Ereignis f?r fehlenden ODBC-Connectionstring
' (Weitergabe des OdbcHandler-Ereignisses: <see cref=OdbcHandler#ErrorEmptyConnectionString>ErrorEmptyConnectionString</see>)
' </summary>
' <param name="Msg">Meldung ?ber Ursache der Ereignisausl?sung. (Wird noch nicht genutzt)</param>
' <param name="NewConnectionString">ODBC-Connectionstring</param>
' <remarks>?ber den ByRef-Parameter NewConnectionString kann der passende ODBC-Connectionstring weitergegeben werden.</remarks>
'**/
Public Event ErrorEmptyOdbcConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'/**
' <summary>
' Ereignis-Meldung wenn Objektreferenzen entfernt wurden
' </summary>
' <remarks></remarks>
'**/
Public Event Disposed()

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
' Sub: Dispose (06.04.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Objektreferenzen entfernen
' </summary>
' <remarks></remarks>
'**/
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
'
' ADODB-Member
' ------------
'/** @name ADO
' *  ADO-Zugriff
' */
'//@{

'---------------------------------------------------------------------------------------
' Property: ADODB
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ADODB-Zugriffsmethoden
' </summary>
' <returns>AdodbHandler</returns>
' <remarks></remarks>
'**/
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
'/**
' <summary>
' ADODB-Connection zur akutellen DBSM-Verbindung = <see cref=AdodbHandler#CurrentConnection>Me.ADODB.CurrentConnection</see>
' </summary>
' <returns>ADODB.Connection</returns>
' <remarks></remarks>
'*/
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnection() As ADODB.Connection
   Set CurrentConnection = Me.ADODB.CurrentConnection
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentConnection = <see cref=AdodbHandler#CurrentConnection>Me.ADODB.CurrentConnection</see>
' </summary>
' <param name="NewConnection">Neue ADO-Connection</param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentConnection(ByRef NewConnection As ADODB.Connection)
   Set Me.ADODB.CurrentConnection = NewConnection
End Property

'---------------------------------------------------------------------------------------
' Property: OledbConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OledbConnectionString = <see cref=AdodbHandler#ConnectionString>Me.ADODB.ConnectionString</see>
' </summary>
' <returns>String</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get OledbConnectionString() As String
   OledbConnectionString = Me.ADODB.ConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: OleDBConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OledbConnectionString = <see cref=AdodbHandler#ConnectionString>Me.ADODB.ConnectionString</see>
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let OledbConnectionString(ByVal NewConnectionString As String)
   Me.ADODB.ConnectionString = NewConnectionString
End Property

'//@}

'#######################################################################################
'
' DAO-Member
' --------------
'/** @name DAO
' *  DAO-Zugriff
' */
'//@{

'---------------------------------------------------------------------------------------
' Property: DAO
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DAO-Zugriffsmethoden
' </summary>
' <returns>DaoHandler</returns>
' <remarks></remarks>
'**/
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
'/**
' <summary>
' CurrentDb = <see cref=DaoHandler#CurrentDb>Me.DAO.CurrentDb</see>
' </summary>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDb() As DAO.Database
   Set CurrentDb = Me.DAO.CurrentDb
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDb
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentDb = <see cref=DaoHandler#CurrentDb>Me.DAO.CurrentDb</see>
' </summary>
' <param name="NewCurrentDb">DAO.Database-Referenz die f?r CurrentDb verwendet werden soll</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDb(ByRef NewCurrentDb As DAO.Database)
   Set Me.DAO.CurrentDb = NewCurrentDb
End Property

'//@}


'#######################################################################################
'
' ODBC-Member
' --------------
'/** @name ODBC
' *  ODBC-Zugriff
' */
'//@{

'---------------------------------------------------------------------------------------
' Property: ODBC
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-Zugriffsmethoden
' </summary>
' <returns>OdbcHandler</returns>
' <remarks></remarks>
'**/
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
'/**
' <summary>
' ODBC-ConnectionString = <see cref=OdbcHandler#ConnectionString>Me.ODBC.ConnectionString</see>
' </summary>
' <returns>String</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get OdbcConnectionString() As String
   OdbcConnectionString = Me.ODBC.ConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: OdbcConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC-ConnectionString = <see cref=OdbcHandler#ConnectionString>Me.ODBC.ConnectionString</see>
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Let OdbcConnectionString(ByVal NewConnectionString As String)
   Me.ODBC.ConnectionString = NewConnectionString
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentDbBE = <see cref=OdbcHandler#CurrentDbBE>Me.ODBC.CurrentDbBE</see>
' </summary>
' <returns>DAO.Database auf ODBC-Datenbank</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDbBE() As DAO.Database
   Set CurrentDbBE = Me.ODBC.CurrentDbBE
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDbBE
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CurrentDbBE = <see cref=OdbcHandler#CurrentDbBE>Me.ODBC.CurrentDbBE</see>
' </summary>
' <param name="NewCurrentDbBE">DAO.Database-Referenz auf eine bereits ge?ffnete Datanbank</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set CurrentDbBE(ByRef NewDb As DAO.Database)
   If Not (NewDb Is Nothing) Then
      Set Me.ODBC.CurrentDbBE = NewDb
   End If
End Property

'//@}

'#######################################################################################
'
' Datenzugriff-Events
' -------------------

' ADODB-Handler-Events
' --------------------
Private Sub m_AdodbHandler_ErrorMissingCurrentConnection(ByVal Msg As String, ByRef NewCurrentConnection As ADODB.Connection)
   RaiseEvent ErrorMissingCurrentConnection(Msg, NewCurrentConnection)
End Sub

Private Sub m_AdodbHandler_ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)
   RaiseEvent ErrorEmptyOledbConnectionString(Msg, NewConnectionString)
End Sub


' DAO-Handler-Events
' ------------------
Private Sub m_DaoHandler_ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)
   RaiseEvent ErrorMissingCurrentDb(Msg, NewCurrentDbRef)
End Sub


' ODBC-Handler-Events
' ------------------
Private Sub m_OdbcHandler_ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)
   RaiseEvent ErrorEmptyOdbcConnectionString(Msg, NewConnectionString)
End Sub

Private Sub m_OdbcHandler_ErrorMissingCurrentDb(ByVal Msg As String, ByRef NewCurrentDbRef As DAO.Database)

   '1. Versuchen aus DAO-Klasse zu ?bergeben:
   Set NewCurrentDbRef = Me.DAO.CurrentDb
   
   '2. Event ausl?sen, falls auch Me.DAO.CurrentDb nichts liefert
   If NewCurrentDbRef Is Nothing Then
      RaiseEvent ErrorMissingCurrentDb(Msg, NewCurrentDbRef)
      If Not NewCurrentDbRef Is Nothing Then
         Set Me.DAO.CurrentDb = NewCurrentDbRef
      End If
   End If
   
End Sub
