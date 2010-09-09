VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TempDbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Erzeugt und verwaltet eine mdb/accdb für temporäre Tabellen"
'---------------------------------------------------------------------------------------
' Klassenmodul: TempDbHandler
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Erzeugt und verwaltet eine mdb/accdb für temporäre Tabellen
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/TempDbHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/modErrorHandler.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const m_conFilePrefix As String = "TDB"
Private Const m_conFileExtension As String = vbNullString 'Dateinamenerweiterung: vbNullString ... Standard von GetTempFileNameA wird verwendet ("tmp")
Private Const m_conMaxPathLen As Long = 255

Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" ( _
         ByVal nBufferLength As Long, _
         ByVal lpBuffer As String) As Long

Private Declare Function API_GetTempFilename Lib "kernel32" Alias "GetTempFileNameA" ( _
         ByVal lpszPath As String, _
         ByVal lpPrefixString As String, _
         ByVal wUnique As Long, _
         ByVal lpTempFileName As String) As Long

Private m_TempDb As DAO.Database
Private m_TempDbFile As String
Private m_SharedMode As Boolean
Private m_DeleteFileOnTerminiate As Boolean
Private m_Disposed As Boolean
         
Private m_ClientDatabase As DAO.Database
         
'---------------------------------------------------------------------------------------
' Event: TempDbHandlerError (03.05.2008)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Ereignis für Fehlermeldungen
' </summary>
' <param name="ErrMsg">Fehlermeldung</param>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="Response">ACLibErrorResumeMode-Wert</param>
' <remarks>Ermöglich die Behebung des Fehlers bzw. die Behandlung des Fehlers im Frontend</remarks>
'**/
Public Event TempDbHandlerError(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef Response As TempDbErrorResumeMode)

Public Enum TempDbErrorResumeMode
   aclibTempDbErrExit = 0       'Abbruch
   aclibTempDbErrResume = 1     'Resume, Problem wurde (von außen) behoben
   aclibTempDbErrResumeNext = 2 'Resume next, im Code weiterarbeiten
End Enum


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

   clearObjectReferences
   
   m_TempDbFile = vbNullString
   
   m_Disposed = True
   RaiseEvent Disposed
   
End Sub

Private Sub clearObjectReferences()

On Error Resume Next

   Set m_ClientDatabase = Nothing

   If Not (m_TempDb Is Nothing) Then
      m_TempDb.Close
      Access.DBEngine.Idle dbRefreshCache
      Set m_TempDb = Nothing
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
   
On Error GoTo HandleErr

   If m_ClientDatabase Is Nothing Then
      RaiseEvent ErrorMissingClientDatabase("TempDbHandler.ClientDatabase: DAO.Database von Client-Anwendung fehlt.", vbObjectError, m_ClientDatabase)
      If m_ClientDatabase Is Nothing Then 'CurrentDb verwenden
         Set m_ClientDatabase = CurrentDb
      End If
   End If
   
   Set ClientDatabase = m_ClientDatabase

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "ClientDatabase", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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

On Error GoTo HandleErr

   Set m_ClientDatabase = db

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "ClientDatabase", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr

   checkCurrentDatabase
   Set CurrentDatabase = m_TempDb

ExitHere:
On Error Resume Next
   Exit Property

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "CurrentDatabase", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
     
On Error GoTo HandleErr

   On Error Resume Next

   If Not (m_TempDb Is Nothing) Then
      m_TempDb.Close
      Set m_TempDb = Nothing
   End If

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "CloseCurrentDatabase", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub


Private Sub checkCurrentDatabase()

   Dim strFile As String
   
On Error Resume Next

   If m_TempDb Is Nothing Then
      If Len(m_TempDbFile) > 0 And Len(Dir$(m_TempDbFile)) > 0 Then
         Set m_TempDb = Access.DBEngine.OpenDatabase(m_TempDbFile, Not m_SharedMode, False)
         strFile = m_TempDb.Name 'Zugriff testen
         If Err.Number <> 0 Then
            Err.Clear
            Me.CreateDatabase
         End If
      Else
         Me.CreateDatabase
      End If
   Else
      Err.Clear
      strFile = m_TempDb.Name 'Zugriff testen
      If Err.Number <> 0 Then
         Err.Clear
         If Len(m_TempDbFile) > 0 And Len(Dir$(m_TempDbFile)) > 0 Then
            Set m_TempDb = Access.DBEngine.OpenDatabase(m_TempDbFile, Not m_SharedMode, False)
            strFile = m_TempDb.Name 'Zugriff testen
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
   
 
On Error GoTo HandleErr

   m_TempDbFile = getTempFileName(, m_conFilePrefix, m_conFileExtension)
   Set m_TempDb = Access.DBEngine.CreateDatabase(m_TempDbFile, dbLangGeneral) ' => DB wird exklusiv geöffnet!
   
   m_SharedMode = OpenSharedMode
   m_DeleteFileOnTerminiate = DeleteFileOnTerminate
   
   'schließen und wieder öffnen, um nicht mehr im Exclusiv-Modus zu sein.
   If m_SharedMode Then
      m_TempDb.Close
      Set m_TempDb = Access.DBEngine.OpenDatabase(m_TempDbFile, False, False)
   End If
   
   Set CreateDatabase = m_TempDb

ExitHere:
   On Error GoTo 0
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "CreateDatabase", Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

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
   
On Error GoTo HandleErr

   m_TempDbFile = DatabaseFile
   m_SharedMode = Not Exclusive
   Set m_TempDb = Access.DBEngine.OpenDatabase(DatabaseFile, Exclusive, ReadOnly)
   m_DeleteFileOnTerminiate = False 'DB nicht löschen, da nicht von Klasse erstellt
   Set OpenDatabase = m_TempDb

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "OpenDatabase : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
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

   Dim strFile As String
   
On Error Resume Next ' Da hier eine Fehlermeldung nutzlos ist.
   
   If Not (m_TempDb Is Nothing) Then
      strFile = m_TempDb.Name
      m_TempDb.Close
      Access.DBEngine.Idle dbRefreshCache
      DoEvents
      Set m_TempDb = Nothing
   Else
      strFile = m_TempDbFile
   End If
   
   If Len(strFile) > 0 Then
      'Falls Datei noch im Zugriff ist, kann es sein, dass sie nicht gelöscht werden kann
      Kill strFile
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
              Optional ByVal CreateTableDDL As String, _
              Optional ByVal LinkedTableDefName As String) As Boolean
   
On Error GoTo HandleErr

   If Not SourceTableDefExists(SourceTableDefName) Then
      If Len(CreateTableDDL) > 0 Then
         CheckTable = Me.CreateTable(SourceTableDefName, CreateTableDDL, LinkedTableDefName)
      Else
         CheckTable = False
      End If
      Exit Function
   End If
   
   If Len(LinkedTableDefName) > 0 Then
      If LinkedTableDefExists(LinkedTableDefName) Then
         reconnectTempTable LinkedTableDefName, SourceTableDefName
      Else
         connectTempTable LinkedTableDefName, SourceTableDefName
      End If
   End If
   
   CheckTable = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "CheckTable : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
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
                     ByVal CreateTableDDL As String, _
               Optional ByVal LinkedTableDefName As String = vbNullString _
                     ) As Boolean

On Error GoTo HandleErr

   checkCurrentDatabase
     
   If SourceTableDefExists(SourceTableDefName) Then 'zur Sicherheit löschen, falls unterschiedliche DDL-Anweisung
      dropSourceTable SourceTableDefName
   End If
   m_TempDb.Execute CreateTableDDL
   
   If Len(LinkedTableDefName) > 0 Then
      If LinkedTableDefExists(LinkedTableDefName) Then
         reconnectTempTable LinkedTableDefName, SourceTableDefName
      Else
         connectTempTable LinkedTableDefName, SourceTableDefName
      End If
   End If
   
   CreateTable = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "CreateTable : " & Err.Source, Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select

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
                     
On Error GoTo HandleErr

   If Len(LinkedTableDefName) > 0 Then
      dropLinkedTableDef LinkedTableDefName
   End If
   
   If Len(SourceTableDefName) > 0 Then
      dropSourceTable SourceTableDefName
   End If
   
   DropTable = True

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "DropTable : " & Err.Source, Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select


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

   Dim strFile As String

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Dim tdfCnt As Long, i As Long
   
On Error GoTo HandleErr

   If Not (m_TempDb Is Nothing) Then
      checkCurrentDatabase
      strFile = m_TempDb.Name
   Else
      strFile = m_TempDbFile
   End If
   
   If Len(strFile) = 0 Then Exit Sub
   
   Set db = ClientDatabase
   With db
      tdfCnt = .TableDefs.Count
      For i = (tdfCnt - 1) To 0 Step -1
         Set tdf = .TableDefs(i)
         If tdf.Connect = ";Database=" & strFile Then
            .TableDefs.Delete tdf.Name
         End If
      Next
   End With

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "DropLinkedTableDefs : " & Err.Source, Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select


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
   
On Error GoTo HandleErr

   Set db = ClientDatabase
   db.TableDefs.Refresh 'zur Sicherheit, falls im FE Tab gelöscht wurde
   
   LinkedTableDefExists = False
   
   For Each tdf In db.TableDefs
      If tdf.Name = TableDefName Then
         LinkedTableDefExists = True
         Exit For
      End If
   Next

ExitHere:
   Exit Property

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "LinkedTableDefExists", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
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
   
On Error GoTo HandleErr

   SourceTableDefExists = False

   '1. Aufruf über CurrentDatabase, damit Existenz der Datei sichergestellt ist
   Me.CurrentDatabase.TableDefs.Refresh
   
   For Each tdf In m_TempDb.TableDefs
      If tdf.Name = TableDefName Then
         SourceTableDefExists = True
         Exit For
      End If
   Next

ExitHere:
   Exit Property

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "SourceTableDefExists", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Property

Private Sub dropLinkedTableDef(ByVal TdfName As String)

On Error Resume Next '... damit muss nicht unbedingt auf Existenz geprüft werden
   ClientDatabase.TableDefs.Refresh
   ClientDatabase.TableDefs.Delete TdfName

End Sub

Private Sub dropSourceTable(ByVal SourceTableDefName As String)
On Error Resume Next
   CurrentDatabase.Execute "DROP TABLE [" & SourceTableDefName & "]"
End Sub

Private Sub reconnectTempTable(ByVal LinkedTableDefName As String, ByVal SourceTableDefName As String)
   
On Error GoTo HandleErr

   dropLinkedTableDef LinkedTableDefName
   connectTempTable LinkedTableDefName, SourceTableDefName

ExitHere:
   Exit Sub

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "reconnectTempTable", Err.Description)
   Case TempDbErrorResumeMode.aclibTempDbErrResume
      Resume
   Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
      Resume Next
   Case Else
      Resume ExitHere
   End Select
   
End Sub

Private Sub connectTempTable(ByVal LinkedTableDefName As String, ByVal SourceTableDefName As String)
   
   Dim dbs As DAO.Database
   Dim tdf As DAO.TableDef
   
On Error GoTo HandleErr

   Set dbs = ClientDatabase
   dbs.TableDefs.Refresh
   
   Set tdf = dbs.CreateTableDef(LinkedTableDefName)
   With tdf
      tdf.SourceTableName = SourceTableDefName
      tdf.Connect = ";DATABASE=" & m_TempDb.Name
   End With
   dbs.TableDefs.Append tdf
   
   dbs.TableDefs.Refresh

ExitHere:
   Exit Sub

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "connectTempTable : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

End Sub

Private Function getTempFileName(Optional ByRef sTempPath As String = "", _
                         Optional ByRef FilePrefix As String = "", _
                         Optional ByRef FileExtension As String = "") As String

   Dim strTempFileName As String
   Dim strTempPath As String
   Dim lngRet As Long
   
On Error GoTo HandleErr

   If Len(sTempPath) = 0 Then
      strTempFileName = String$(m_conMaxPathLen, 0)
      lngRet = GetTempPath(m_conMaxPathLen, strTempFileName)
      strTempFileName = Left$(strTempFileName, InStr(strTempFileName, Chr$(0)) - 1)
      strTempPath = strTempFileName
   Else
      strTempPath = sTempPath
   End If
   
   strTempFileName = String$(m_conMaxPathLen, 0)
   lngRet = API_GetTempFilename(strTempPath, FilePrefix, 0&, strTempFileName)
   
   strTempFileName = Left$(strTempFileName, InStr(strTempFileName, Chr$(0)) - 1)
   
   'Datei wieder löschen, da nur Name benötigt wird
   Call Kill(strTempFileName)
   
   If Len(FileExtension) > 0 Then 'Fileextension umschreiben
     strTempFileName = Left$(strTempFileName, Len(strTempFileName) - 3) & FileExtension
   End If
   getTempFileName = strTempFileName

ExitHere:
On Error GoTo 0
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "getTempFileName : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         getTempFileName = vbNullString
         Resume ExitHere
   End Select
  
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
   
On Error GoTo HandleErr

   If (RecordsetOptions And dbSeeChanges) = 0 Then
      RecordsetOptions = RecordsetOptions + dbSeeChanges
   End If
   Set OpenRecordset = CurrentDatabase.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "OpenRecordset : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

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
' <param name="RecordsAffected"></param>

' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub Execute(ByVal sSQL As String, _
               Optional ByVal Options As DAO.RecordsetOptionEnum, _
               Optional ByRef RecordsAffected As Long)

On Error GoTo HandleErr

   If (Options And dbSeeChanges) = 0 Then
      Options = Options + dbSeeChanges
   End If
   CurrentDatabase.Execute sSQL, Options
   RecordsAffected = CurrentDatabase.RecordsAffected

ExitHere:
On Error Resume Next
   Exit Sub

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "Execute : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select
   
End Sub


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
Public Function DLookupSQL(ByVal sSQL As String, Optional ByVal index As Variant = 0&) As Variant

   Dim rst As DAO.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(sSQL, dbOpenForwardOnly, dbSeeChanges, dbReadOnly)
   With rst
      If .EOF Then
         DLookupSQL = Null
      Else
         DLookupSQL = .Fields(index)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
On Error Resume Next
   Exit Function

HandleErr:
   Select Case handleTempDbHandlerError(Err.Number, "DLookupSQL : " & Err.Source, Err.Description)
      Case TempDbErrorResumeMode.aclibTempDbErrResume
         Resume
      Case TempDbErrorResumeMode.aclibTempDbErrResumeNext
         Resume Next
      Case Else
         Resume ExitHere
   End Select

End Function

Private Function handleTempDbHandlerError(ByVal lErrorNumber As Long, ByVal sSource As String, _
                                          ByVal sErrorDescription As String) As TempDbErrorResumeMode
'ACLibErrorResumeMode
'   ErrContinue = 0  'Abbruch ohne Meldung
'   ErrDisplay = 1   'Abbruch mit Meldung
'   Resume = 2       'Resume, Problem von außen behoben
'   ResumeNext = 3   'Resume next, im Code weiterarbeiten

   Dim strMsg As String
   Dim lResponse As TempDbErrorResumeMode
   
On Error Resume Next

   strMsg = sErrorDescription
   If Len(sSource) > 0 Then
      strMsg = sSource & ": " & vbNewLine & strMsg
   End If

   lResponse = -1
   RaiseEvent TempDbHandlerError(strMsg, lErrorNumber, lResponse)
   If lResponse = -1 Then 'Auf event wurde nicht reagiert bzw. lResponse nicht verändert
On Error GoTo 0
#If USE_GLOBAL_ERRORHANDLER Then
      lResponse = HandleError(lErrorNumber, "TempDbHandler." & sSource, sErrorDescription)
#Else
      Err.Raise lErrorNumber, "TempDbHandler." & sSource, sErrorDescription
#End If
   End If

   handleTempDbHandlerError = lResponse

End Function
