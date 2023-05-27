Attribute VB_Name = "modDbConnection"
'---------------------------------------------------------------------------------------
' Modul: modDbConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Sammlung von Prozeduren für die DbConnectionManager-Klasse
' </summary>
' <remarks>Dient zum Instanzieren von DbConnectionManager und für den Zugriff auf die Hauptelemente von DbConnectionManager</remarks>
'\ingroup data
'**/
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

'
'---------------------------------------------------------------------------------------
' Property: DbCon
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Zugriff auf DbConnection der Haupt-DbConnectionManager-Instanz
' </summary>
' <returns>DbConnectionHandler</returns>
' <remarks>
' Erleichtert den Aufruf, da das Schreiben der DbConnectionManager-Instanz-Variablen entfällt.
' Beim ersten Zugriff wird die Instanz der DbConnectionManager-Klasse erstellt, falls diese noch nicht vorhanden ist.
' </remarks>
'**/
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
'/**
' <summary>
' Zugriff auf DbConnectionInfo der Haupt-DbConnectionManager-Instanz
' </summary>
' <returns>DbConnectionInfo</returns>
' <remarks>
' Erleichtert den Aufruf, da das Schreiben der DbConnectionManager-Instanz-Variablen entfällt.
' Beim ersten Zugriff wird die Instanz der DbConnectionManager-Klasse erstellt, falls diese noch nicht vorhanden ist.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnectionInfo() As DbConnectionInfo

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   Set CurrentConnectionInfo = m_DbConnectionManager.ConnectionInfo

End Property

'---------------------------------------------------------------------------------------
' Function: CheckConnectionStatus
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verbindung prüfen <see cref=DbConnectionManager#CheckConnectionStatus>DbConnectionManager.CheckConnectionStatus</see>
' </summary>
' <returns>Boolean: True = Verbindungsaufbau war erfolgreich</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckConnectionStatus() As Boolean

   If m_DbConnectionManager Is Nothing Then
      Set m_DbConnectionManager = New DbConnectionManager
   End If
   CheckConnectionStatus = m_DbConnectionManager.CheckConnectionStatus

End Function

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

Public Function TableDefExists(ByVal sTableDefName As String, Optional ByRef dbs As DAO.Database = Nothing) As Boolean
'Schneller wäre der Zugriff auf MSysObject (select .. from MSysObject where Name = 'Tabellenname' AND Type IN (1, 4, 6)
'Eine weitere Alternative wäre das Auswerten über cnn.OpenSchema(adSchemaTables, ...) ... dann werden allerdings keine verknüpften Tabellen geprüft
   
   Dim tdf As DAO.TableDef

   If dbs Is Nothing Then
      Set dbs = CurrentDb
   End If
   
   dbs.TableDefs.Refresh
   For Each tdf In dbs.TableDefs
      If tdf.Name = sTableDefName Then
         TableDefExists = True
         Exit For
      End If
   Next

End Function

Public Function QueryDefExists(ByVal sQueryDefName As String, Optional ByVal dbs As DAO.Database = Nothing) As Boolean
   
   Dim qdf As DAO.QueryDef

   If dbs Is Nothing Then
      Set dbs = CurrentDb
   End If
   
   dbs.QueryDefs.Refresh
   For Each qdf In dbs.QueryDefs
      If qdf.Name = sQueryDefName Then
         QueryDefExists = True
         Exit For
      End If
   Next

End Function
