Attribute VB_Name = "modDbConnection"
'---------------------------------------------------------------------------------------
' Package: data.modDbConnection
'---------------------------------------------------------------------------------------
'
' Collection of procedures for the DbConnectionManager class
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     Used to instantiate DbConnectionManager and access the main elements of DbConnectionManager.
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/modDbConnection.bas</file>
'  <license>_codelib/license.bas</license>
'  <use>data/DbConnectionManager.cls</use>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit
Option Private Module

Private m_DbConnectionManager As DbConnectionManager ' Hauptsteuerung

'#######################################################################################
' Group: Factories

'---------------------------------------------------------------------------------------
' Property: DbCon
'---------------------------------------------------------------------------------------
'
' Access to DbConnection of the main DbConnectionManager instance
'
' Returns:
'     DbConnectionHandler
'
' Remarks:
'     Simplifies the call, as the writing of the DbConnectionManager instance variables is omitted.
'     The first access creates the instance of the DbConnectionManager class if it does not already exist.
'
'---------------------------------------------------------------------------------------
Public Property Get DbCon() As DbConnectionHandler

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set DbCon = m_DbConnectionManager.DbConnection

End Property

'---------------------------------------------------------------------------------------
' Property: CurrentConnectionInfo
'---------------------------------------------------------------------------------------
'
' Access to DbConnectionInfo of the main DbConnectionManager instance
'
' Returns:
'     <data.DbConnectionInfo>
'
' Remarks:
'     Simplifies the call, as the writing of the DbConnectionManager instance variables is omitted.
'     The first access creates the instance of the DbConnectionManager class if it does not already exist.
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnectionInfo() As DbConnectionInfo

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set CurrentConnectionInfo = m_DbConnectionManager.ConnectionInfo

End Property

'#######################################################################################
' Group: Global Procedures

'---------------------------------------------------------------------------------------
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'
' Verbindung prüfen (call <data.DbConnectionManager.CheckConnectionStatus>)
'
' Return:
'     Boolean - True if success
'
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   CheckConnectionStatus = m_DbConnectionManager.CheckConnectionStatus

End Function

'---------------------------------------------------------------------------------------
' Sub: DisposeDbConnection
'---------------------------------------------------------------------------------------
'
' Dispose DbConnection objects (incl. DbConnectionManager)
'
'---------------------------------------------------------------------------------------
Public Sub DisposeDbConnection()

On Error Resume Next

   If Not (m_DbConnectionManager Is Nothing) Then
      m_DbConnectionManager.Dispose
   End If
   Set m_DbConnectionManager = Nothing
   
End Sub

'############################
'
' Hilfsfunktionen

'---------------------------------------------------------------------------------------
' Function: TableDefExists
'---------------------------------------------------------------------------------------
'
' Check if TableDef exists
'
' Parameters:
'     TableDefName   - (String) TableDef name
'     db             - (DAO.Database) Database to use (optional: if nothing CurrentDb will be used)
'
'---------------------------------------------------------------------------------------
Public Function TableDefExists(ByVal TableDefName As String, _
                      Optional ByVal db As DAO.Database = Nothing) As Boolean
'Schneller wäre der Zugriff auf MSysObject (select .. from MSysObject where Name = 'Tabellenname' AND Type IN (1, 4, 6)
'Eine weitere Alternative wäre das Auswerten über cnn.OpenSchema(adSchemaTables, ...) ... dann werden allerdings keine verknüpften Tabellen geprüft
   
   Dim tdf As DAO.TableDef

   If db Is Nothing Then
      Set db = CurrentDb
   End If
   
   db.TableDefs.Refresh
   For Each tdf In db.TableDefs
      If tdf.Name = TableDefName Then
         TableDefExists = True
         Exit For
      End If
   Next

End Function

'---------------------------------------------------------------------------------------
' Function: QueryDefExists
'---------------------------------------------------------------------------------------
'
' Check if QueryDef exists
'
' Parameters:
'     QueryDefName   - (String) QueryDef name
'     db             - (DAO.Database) Database to use (optional: if nothing CurrentDb will be used)
'
'---------------------------------------------------------------------------------------
Public Function QueryDefExists(ByVal QueryDefName As String, _
                      Optional ByVal db As DAO.Database = Nothing) As Boolean
   
   Dim qdf As DAO.QueryDef

   If db Is Nothing Then
      Set db = CurrentDb
   End If
   
   db.QueryDefs.Refresh
   For Each qdf In db.QueryDefs
      If qdf.Name = QueryDefName Then
         QueryDefExists = True
         Exit For
      End If
   Next

End Function
