VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempDbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: TempDbHandler
'---------------------------------------------------------------------------------------
'/**
' \author       Josef Poetzl
' <summary>
' Erzeugt und verwaltet eine mdb/accdb für temporaere Tabellen
' </summary>
' <remarks></remarks>
'\ingroup data_dao
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/TempDbHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoHandler.cls</use>
'  <test>_test/data/dao/TempDbHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ClassName As String = "TempDbHandler"

Private m_DaoHandler As DaoHandler

Private Const FILE_PREFIX As String = "TDB"
Private Const FILE_EXTENSION As String = vbNullString 'Dateinamenerweiterung: vbNullString ... Standard von GetTempFileNameA wird verwendet ("tmp")
Private Const MAX_PATH_LEN As Long = 255

Private m_TempDbFile As String
Private m_SharedMode As Boolean
Private m_DeleteFileOnTerminiate As Boolean
Private m_Disposed As Boolean

Private m_ClientDatabase As DAO.Database

Private Enum ErrorNumbers
      ErrTableCanNotBeCreatedTableExists = vbObjectError + 9001
End Enum
         
#If VBA7 Then

Private Declare PtrSafe Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare PtrSafe Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

#Else

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

#End If

'---------------------------------------------------------------------------------------
' Event: ErrorMissingDbConnection (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ereignis-Meldung wenn DbConnection fehlt
' </summary>
' <remarks>Ermöglicht Fehlerbehandlung vom FE aus</remarks>
'**/
Public Event ErrorMissingClientDatabase(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewClientDatabaseRef As DAO.Database)

'---------------------------------------------------------------------------------------
' Event: Disposed (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ereignis-Meldung wenn Objektreferenzen entfernt wurden
' </summary>
' <remarks></remarks>
'**/
Public Event Disposed()


Private Sub Class_Initialize()
   Set m_DaoHandler = New DaoHandler
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   If Not m_Disposed Then Dispose
End Sub

'---------------------------------------------------------------------------------------
' Sub: Dispose (06.04.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Dispose
' </summary>
' <param name="Param"></param>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Dispose()

   If m_Disposed Then Exit Sub

On Error Resume Next

   'aktuelle TempDb löschen
   If m_DeleteFileOnTerminiate Then DeleteCurrentDatabase

   ClearObjectReferences
   
   m_TempDbFile = vbNullString
   
   m_Disposed = True
   RaiseEvent Disposed
   
End Sub

Private Sub ClearObjectReferences()

On Error Resume Next

   Set m_ClientDatabase = Nothing

   If Not (m_DaoHandler Is Nothing) Then
      If Not (m_DaoHandler.CurrentDbReferenceOnly Is Nothing) Then
         m_DaoHandler.CurrentDbReferenceOnly.Close
      End If
      m_DaoHandler.Dispose
      Access.DBEngine.Idle dbRefreshCache
      Set m_DaoHandler = Nothing
   End If

End Sub

'---------------------------------------------------------------------------------------
' Property: ClientDatabase (2009-12-21)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die aktuelle Database-Instanz der Client-Anwendung (FE) zurück.
' </summary>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ClientDatabase() As DAO.Database

   If m_ClientDatabase Is Nothing Then
      RaiseEvent ErrorMissingClientDatabase("TempDbHandler.ClientDatabase: DAO.Database von Client-Anwendung fehlt.", vbObjectError, m_ClientDatabase)
      If m_ClientDatabase Is Nothing Then 'CurrentDb verwenden
         Set m_ClientDatabase = CurrentDb
      End If
   End If
   
   Set ClientDatabase = m_ClientDatabase

End Property

'---------------------------------------------------------------------------------------
' Property: ClientDatabase (2009-12-21)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die aktuelle Database-Instanz der Client-Anwendung (FE) zurück.
' </summary>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Set ClientDatabase(ByRef db As DAO.Database)
   Set m_ClientDatabase = db
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDatabase (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Gibt die aktuelle Database-Instanz der TempDB zurück.
' </summary>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get CurrentDatabase() As DAO.Database
   CheckCurrentDatabase
   Set CurrentDatabase = m_DaoHandler.CurrentDb
End Property

'---------------------------------------------------------------------------------------
' Sub: CloseCurrentDatabase (25.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Schließt Datenbank, löscht sie aber nicht ... DB wird beim nächsten Zugriffsversuch wieder geöffnet
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub CloseCurrentDatabase()

On Error Resume Next

   If Not (m_DaoHandler.CurrentDb Is Nothing) Then
      m_DaoHandler.CurrentDb.Close
      Set m_DaoHandler.CurrentDb = Nothing
   End If

End Sub


Private Sub CheckCurrentDatabase()

   Dim FilePath As String
   
On Error Resume Next

   If m_DaoHandler.CurrentDb Is Nothing Then
      If Len(m_TempDbFile) > 0 And Len(Dir$(m_TempDbFile)) > 0 Then
         Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(m_TempDbFile, Not m_SharedMode, False)
         FilePath = m_DaoHandler.CurrentDb.Name 'Zugriff testen
         If Err.Number <> 0 Then
            Err.Clear
            Me.CreateDatabase
         End If
      Else
         Me.CreateDatabase
      End If
   Else
      Err.Clear
      FilePath = m_DaoHandler.CurrentDb.Name 'Zugriff testen
      If Err.Number <> 0 Then
         Err.Clear
         If Len(m_TempDbFile) > 0 And Len(Dir$(m_TempDbFile)) > 0 Then
            Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(m_TempDbFile, Not m_SharedMode, False)
            FilePath = m_DaoHandler.CurrentDb.Name 'Zugriff testen
            If Err.Number <> 0 Then
               Err.Clear
               Me.CreateDatabase
            End If
         Else
            Me.CreateDatabase
         End If
      End If
   End If

End Sub

'---------------------------------------------------------------------------------------
' Function: CreateDatabase (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt neue mdb im %Temp%-Verzeichnis
' </summary>
' <param name="OpenSharedMode">True: öffnet die Datenbank im geteilten Modus, False: öffnet die Datenbank im Exklusivmodus. Standard: False</param>
' <param name="DeleteFileOnTerminate">True: Datei wird bei Class_Terminate gelöscht. Standard: True</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateDatabase(Optional ByVal OpenSharedMode As Boolean = False, _
                               Optional ByVal DeleteFileOnTerminate As Boolean = True) As DAO.Database

   m_TempDbFile = GetTempFileName(, FILE_PREFIX, FILE_EXTENSION)
   
   If CurrentDb.Version = "4.0" Then
      'Set m_tempDb = Access.DBEngine.CreateDatabase(m_TempDbFile, dbLangGeneral, dbVersion40)
      Set m_DaoHandler.CurrentDb = DBEngine.Workspaces(0).CreateDatabase(m_TempDbFile, dbLangGeneral, dbVersion40)
   Else
      Set m_DaoHandler.CurrentDb = Access.DBEngine.CreateDatabase(m_TempDbFile, dbLangGeneral)
   End If ' => DB wird exklusiv geöffnet!
   
   m_SharedMode = OpenSharedMode
   m_DeleteFileOnTerminiate = DeleteFileOnTerminate
   
   'schließen und wieder öffnen, um nicht mehr im Exclusiv-Modus zu sein.
   If m_SharedMode Then
      m_DaoHandler.CurrentDb.Close
      Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(m_TempDbFile, False, False)
   End If
   
   Set CreateDatabase = m_DaoHandler.CurrentDb

End Function

'---------------------------------------------------------------------------------------
' Function: OpenDatabase (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Öffnet bestehende Datenbank (mdb) als Temp-DB.
' </summary>
' <param name="DatabaseFile">Die zu öffnende Datei</param>
' <param name="Exclusive">True: öffnet die Datenbank im Exklusivmodus, False: öffnet die Datenbank im geteilten Modus. Standard: False</param>
' <param name="ReadOnly">öffnet die Datenbank schreibgeschützt</param>
' <returns>DAO.Database</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenDatabase(ByVal DatabaseFile As String, _
                  Optional ByVal Exclusive As Boolean = False, _
                  Optional ByVal ReadOnly As Boolean = False) As DAO.Database

   m_TempDbFile = DatabaseFile
   m_SharedMode = Not Exclusive
   Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(DatabaseFile, Exclusive, ReadOnly)
   m_DeleteFileOnTerminiate = False 'DB nicht löschen, da nicht von Klasse erstellt
   Set OpenDatabase = m_DaoHandler.CurrentDb

End Function

'---------------------------------------------------------------------------------------
' Function: DeleteCurrentDatabase (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Löscht mdb-Datei der aktuellen Database-Instanz.
' </summary>
' <remarks>Wird automatisch bei Class_Terminate ausgeführt, wenn Datei per TempDbHandler.CreateNewDatabase erzeugt wurde.</remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub DeleteCurrentDatabase()

   Dim DbFile As String
   
On Error Resume Next ' Da hier eine Fehlermeldung nutzlos ist.
   
   If Not (m_DaoHandler.CurrentDb Is Nothing) Then
      DbFile = m_DaoHandler.CurrentDb.Name
      m_DaoHandler.CurrentDb.Close
      Access.DBEngine.Idle dbRefreshCache
      DoEvents
      Set m_DaoHandler.CurrentDb = Nothing
   Else
      DbFile = m_TempDbFile
   End If
   
   If Len(DbFile) > 0 Then
      'Falls Datei noch im Zugriff ist, kann es sein, dass sie nicht gelöscht werden kann
      Kill DbFile
   End If
   
   m_TempDbFile = vbNullString
   
End Sub

'---------------------------------------------------------------------------------------
' Function: CheckTempTable (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' CheckTempTable
' </summary>
' <param name="SourceTableDefName">Tabellenname im TempDB (Backend)</param>
' <param name="CreateTableDDL">DDL-Anweisung zum Erstellen der Tabelle, falls Tabelle nicht vorhanden ist.</param>
' <param name="LinkedTableDefName">Tabellenname im Frontend</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckTable(ByVal SourceTableDefName As String, _
              Optional ByVal CreateTableSql As String, _
              Optional ByVal LinkedTableDefName As String) As Boolean

   If Not SourceTableDefExists(SourceTableDefName) Then
      If Len(CreateTableSql) > 0 Then
         CheckTable = Me.CreateTable(SourceTableDefName, CreateTableSql, LinkedTableDefName)
      Else
         CheckTable = False
      End If
      Exit Function
   End If
   
   If Len(LinkedTableDefName) > 0 Then
      If LinkedTableDefExists(LinkedTableDefName) Then
         ReconnectTempTable LinkedTableDefName, SourceTableDefName
      Else
         ConnectTempTable LinkedTableDefName, SourceTableDefName
      End If
   End If
   
   CheckTable = True

End Function


'---------------------------------------------------------------------------------------
' Function: CreateTable (03.05.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erstellt Tabelle per DDL-Anweisung
' </summary>
' <param name="SourceTableDefName">Tabellenname im TempDB (Backend)</param>
' <param name="CreateTableDDL">DDL-Anweisung zum Erstellen der Tabelle</param>
' <param name="LinkedTableDefName">Tabellenname im Frontend</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CreateTable(ByVal SourceTableDefName As String, _
                            ByVal CreateTableSql As String, _
                   Optional ByVal LinkedTableDefName As String = vbNullString _
                     ) As Boolean

   CheckCurrentDatabase
   If SourceTableDefExists(SourceTableDefName) Then 'zur Sicherheit löschen, falls unterschiedliche DDL-Anweisung
      Err.Raise ErrorNumbers.ErrTableCanNotBeCreatedTableExists, ClassName & CreateTable, "Table '" & SourceTableDefName & "' exists"
   End If
   
   m_DaoHandler.Execute CreateTableSql
   
   If Len(LinkedTableDefName) > 0 Then
      If LinkedTableDefExists(LinkedTableDefName) Then
         ReconnectTempTable LinkedTableDefName, SourceTableDefName
      Else
         ConnectTempTable LinkedTableDefName, SourceTableDefName
      End If
   End If
   
   CreateTable = True

End Function


'---------------------------------------------------------------------------------------
' Function: DropTable (03.05.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Löscht Tabelle
' </summary>
' <param name="Param"></param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function DropTable(ByVal SourceTableDefName As String, _
             Optional ByVal LinkedTableDefName As String = vbNullString) As Boolean
                     
   If Len(LinkedTableDefName) > 0 Then
      DropLinkedTableDef LinkedTableDefName
   End If
   
   If Len(SourceTableDefName) > 0 Then
      DropSourceTable SourceTableDefName
   End If
   
   DropTable = True

End Function

'---------------------------------------------------------------------------------------
' Sub: DropLinkedTableDefs (03.05.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Entfernt TableDef-Objekte, die mit Tabellen der TempDb verknüpft sind
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub DropLinkedTableDefs()

   Dim DbFile As String
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   Dim TdfCnt As Long, i As Long

   If Not (m_DaoHandler.CurrentDb Is Nothing) Then
      CheckCurrentDatabase
      DbFile = m_DaoHandler.CurrentDb.Name
   Else
      DbFile = m_TempDbFile
   End If
   
   If Len(DbFile) = 0 Then Exit Sub
   
   Set db = ClientDatabase
   With db
      TdfCnt = .TableDefs.Count
      For i = (TdfCnt - 1) To 0 Step -1
         Set tdf = .TableDefs(i)
         If tdf.Connect = ";Database=" & DbFile Then
            .TableDefs.Delete tdf.Name
         End If
      Next
   End With
   
End Sub

'---------------------------------------------------------------------------------------
' Property: LinkedTableDefExists (03.05.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft ob Tabelle im FE vorhanden ist
' </summary>
' <param name="TableDefName">Tabellenname</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get LinkedTableDefExists(ByVal TableDefName As String) As Boolean
   
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set db = ClientDatabase
   db.TableDefs.Refresh 'zur Sicherheit, falls im FE Tab gelöscht wurde
   
   LinkedTableDefExists = False
   
   For Each tdf In db.TableDefs
      If tdf.Name = TableDefName Then
         LinkedTableDefExists = True
         Exit For
      End If
   Next

End Property

'---------------------------------------------------------------------------------------
' Property: SourceTableDefExists (24.02.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Prüft ob Tabelle in Temp-Datenbank (BE) vorhanden ist
' </summary>
' <param name="TableDefName">Tabellenname</param>
' <returns>Boolean</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get SourceTableDefExists(ByVal TableDefName As String) As Boolean
   
   Dim tdf As DAO.TableDef

   SourceTableDefExists = False

   '1. Aufruf über CurrentDatabase, damit Existenz der Datei sichergestellt ist
   Me.CurrentDatabase.TableDefs.Refresh
   
   For Each tdf In m_DaoHandler.CurrentDb.TableDefs
      If tdf.Name = TableDefName Then
         SourceTableDefExists = True
         Exit For
      End If
   Next

End Property

Private Sub DropLinkedTableDef(ByVal TdfName As String)

On Error Resume Next '... damit muss nicht unbedingt auf Existenz geprüft werden
   ClientDatabase.TableDefs.Refresh
   ClientDatabase.TableDefs.Delete TdfName

End Sub

Private Sub DropSourceTable(ByVal SourceTableDefName As String)

On Error GoTo HandleErr

   CurrentDatabase.Execute "DROP TABLE [" & SourceTableDefName & "]", dbFailOnError

ExitHere:
   Exit Sub

HandleErr:
   Select Case Err.Number
   Case 3376 'Table does not exist.
      Err.Raise 123, Err.Source, Err.Description
      Resume ExitHere
   Case Else
      Err.Raise Err.Number, Err.Source, Err.Description
   End Select
   
End Sub

Private Sub ReconnectTempTable(ByVal LinkedTableDefName As String, ByVal SourceTableDefName As String)

   DropLinkedTableDef LinkedTableDefName
   ConnectTempTable LinkedTableDefName, SourceTableDefName
 
End Sub

Private Sub ConnectTempTable(ByVal LinkedTableDefName As String, ByVal SourceTableDefName As String)
   
   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef

   Set dbs = ClientDatabase
   dbs.TableDefs.Refresh
   
   Set tdf = dbs.CreateTableDef(LinkedTableDefName)
   With tdf
      tdf.SourceTableName = SourceTableDefName
      tdf.Connect = ";DATABASE=" & m_DaoHandler.CurrentDb.Name
   End With
   dbs.TableDefs.Append tdf
   
   dbs.TableDefs.Refresh

End Sub

Private Function GetTempFileName(Optional ByRef TempPath As String = "", _
                         Optional ByRef FilePrefix As String = "", _
                         Optional ByRef FileExtension As String = "") As String

   Dim TempFileName As String
   Dim TempFilePath As String

   If Len(TempPath) = 0 Then
      TempFileName = String$(MAX_PATH_LEN, 0)
      GetTempPath MAX_PATH_LEN, TempFileName
      TempFileName = Left$(TempFileName, InStr(TempFileName, Chr$(0)) - 1)
      TempFilePath = TempFileName
   Else
      TempFilePath = TempPath
   End If
   
   TempFileName = String$(MAX_PATH_LEN, 0)
   API_GetTempFilename TempFilePath, FilePrefix, 0&, TempFileName
   
   TempFileName = Left$(TempFileName, InStr(TempFileName, Chr$(0)) - 1)
   
   'Datei wieder löschen, da nur Name benötigt wird
   Call Kill(TempFileName)
   
   If Len(FileExtension) > 0 Then 'Fileextension umschreiben
     TempFileName = Left$(TempFileName, Len(TempFileName) - 3) & FileExtension
   End If
   GetTempFileName = TempFileName

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordset (03.05.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' openRecordset
' </summary>
' <param name="Quelle">SQL-Anweisung</param>
' <param name="RecordsetType">DAO.RecordsetTypeEnum (Standard: dbOpenForwardOnly)</param>
' <param name="RecordsetOptions">DAO.RecordsetOptionEnum (Standard: dbSeeChanges)</param>
' <param name="LockEdit">DAO.LockTypeEnum (Standard: dbOptimistic)</param>
' <returns>DAO.Recordset</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                       Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                       Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                       Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset
   Set OpenRecordset = m_DaoHandler.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)
End Function

'---------------------------------------------------------------------------------------
' Function: Execute (03.05.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Execute
' </summary>
' <param name="sSQL">SQL-Anweisung</param>
' <param name="Options"></param>
' <returns>Long (Anzahl der betroffenen Datensätze)</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal Query As String, _
          Optional ByVal Options As DAO.RecordsetOptionEnum) As Long
   Execute = m_DaoHandler.Execute(Query, Options)
End Function


'---------------------------------------------------------------------------------------
' Function: DLookupSQL (03.05.2007)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' DLookup-Recordset-Ersatzfunktion für die Übergabe einer kompletten SQL-Anweisung
' </summary>
' <param name="sSQL">SQL-Anweisung</param>
' <param name="Index">Feldauswahl, falls nicht das erste Feld des Recordsets zurückgelieert werden soll</param>
' <returns>Variant</returns>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function LookupSql(ByVal SqlText As String, Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant
   LookupSql = m_DaoHandler.LookupSql(SqlText, Index, ValueIfNull)
End Function
