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
' Class: data.dao.TempDbHandler
'---------------------------------------------------------------------------------------
'
' Creates and manages a database file (mdb/accb) for temporary tables
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/dao/TempDbHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>data/dao/DaoHandler.cls</use>
'  <test>_test/data/dao/TempDbHandlerTests.cls</test>
'  <description>Creates and manages an database file for temporary tables</description>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Const ClassName As String = "TempDbHandler"

Private m_DaoHandler As DaoHandler

Private Const FILE_PREFIX As String = "TDB"
Private Const FILE_EXTENSION As String = vbNullString 'File name extension: vbNullString ... Default of GetTempFileNameA is used ("tmp")
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

'#############################################################
' Group: Events

'---------------------------------------------------------------------------------------
' Event: ErrorMissingDbConnection
'---------------------------------------------------------------------------------------
'
' Event message when DbConnection is missing
'
' Remarks:
'     Enables error handling from the frontend
'
Public Event ErrorMissingClientDatabase(ByVal ErrMsg As String, ByVal ErrNumber As Long, ByRef NewClientDatabaseRef As DAO.Database)

'---------------------------------------------------------------------------------------
' Event: Disposed
'---------------------------------------------------------------------------------------
'
' Event message when object references were removed
'
Public Event Disposed()

'#############################################################
' Group: TempDb configuration

Private Sub Class_Initialize()
   Set m_DaoHandler = New DaoHandler
End Sub

Private Sub Class_Terminate()
On Error Resume Next
   If Not m_Disposed Then Dispose
End Sub

'---------------------------------------------------------------------------------------
' Sub: Dispose
'---------------------------------------------------------------------------------------
'
' Dispose objects, reset members, delete file
'
'---------------------------------------------------------------------------------------
Public Sub Dispose()

   If m_Disposed Then Exit Sub

On Error Resume Next

   '
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
' Property: ClientDatabase
'---------------------------------------------------------------------------------------
'
' Database instance of the client application (frontend)
'
' Returns:
'     DAO.Database
'
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

Public Property Set ClientDatabase(ByRef db As DAO.Database)
   Set m_ClientDatabase = db
End Property

'---------------------------------------------------------------------------------------
' Property: CurrentDatabase
'---------------------------------------------------------------------------------------
'
' Database instance of TempDB
'
' Returns:
'     DAO.Database
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentDatabase() As DAO.Database
   CheckCurrentDatabase
   Set CurrentDatabase = m_DaoHandler.CurrentDb
End Property

'---------------------------------------------------------------------------------------
' Sub: CloseCurrentDatabase
'---------------------------------------------------------------------------------------
'
' Closes database but does not delete it ... DB is opened again at the next access attempt
'
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
      FilePath = m_DaoHandler.CurrentDb.Name 'Test access
      If Err.Number <> 0 Then
         Err.Clear
         If Len(m_TempDbFile) > 0 And Len(Dir$(m_TempDbFile)) > 0 Then
            Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(m_TempDbFile, Not m_SharedMode, False)
            FilePath = m_DaoHandler.CurrentDb.Name 'Test access
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
' Function: CreateDatabase
'---------------------------------------------------------------------------------------
'
' Erzeugt neue mdb im %Temp%-Verzeichnis
'
' Parameters:
'     OpenSharedMode          - True: opens the database in split mode, False: opens the database in exclusive mode. Default: False
'     DeleteFileOnTerminate   - True: File will be deleted at Class_Terminate. Default: True
'
' Returns:
'     DAO.Database
'
'---------------------------------------------------------------------------------------
Public Function CreateDatabase(Optional ByVal OpenSharedMode As Boolean = False, _
                               Optional ByVal DeleteFileOnTerminate As Boolean = True) As DAO.Database

   m_TempDbFile = GetTempFileName(, FILE_PREFIX, FILE_EXTENSION)
   
   If CurrentDb.Version = "4.0" Then
      Set m_DaoHandler.CurrentDb = DBEngine.Workspaces(0).CreateDatabase(m_TempDbFile, dbLangGeneral, dbVersion40)
   Else
      Set m_DaoHandler.CurrentDb = Access.DBEngine.CreateDatabase(m_TempDbFile, dbLangGeneral)
   End If ' => DB is opened exclusively!
   
   m_SharedMode = OpenSharedMode
   m_DeleteFileOnTerminiate = DeleteFileOnTerminate
   
   'close and reopen to no longer be in exclusive mode.
   If m_SharedMode Then
      m_DaoHandler.CurrentDb.Close
      Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(m_TempDbFile, False, False)
   End If
   
   Set CreateDatabase = m_DaoHandler.CurrentDb

End Function

'---------------------------------------------------------------------------------------
' Function: OpenDatabase
'---------------------------------------------------------------------------------------
'
' Opens existing database (mdb) as temp DB.
'
' Parameters:
'     DatabaseFile   - The file to open
'     Exclusive      - True: opens the database in exclusive mode, False: opens the database in split mode. Default: False
'     ReadOnly       - opens the database read-only
'
' Returns:
'     DAO.Database
'
'---------------------------------------------------------------------------------------
Public Function OpenDatabase(ByVal DatabaseFile As String, _
                  Optional ByVal Exclusive As Boolean = False, _
                  Optional ByVal ReadOnly As Boolean = False) As DAO.Database

   m_TempDbFile = DatabaseFile
   m_SharedMode = Not Exclusive
   Set m_DaoHandler.CurrentDb = Access.DBEngine.OpenDatabase(DatabaseFile, Exclusive, ReadOnly)
   m_DeleteFileOnTerminiate = False 'Do not delete DB because not created by class
   Set OpenDatabase = m_DaoHandler.CurrentDb

End Function

'---------------------------------------------------------------------------------------
' Function: DeleteCurrentDatabase
'---------------------------------------------------------------------------------------
'
' Deletes mdb file of the current database instance
'
' Remarks:
'     Will be executed automatically at Class_Terminate if file was created by TempDbHandler.CreateNewDatabase.
'
'---------------------------------------------------------------------------------------
Public Sub DeleteCurrentDatabase()

   Dim DbFile As String
   
On Error Resume Next ' Since an error message is useless here.
   
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
      'If file is still in access, it may not be possible to delete it
      Kill DbFile
   End If
   
   m_TempDbFile = vbNullString
   
End Sub

'---------------------------------------------------------------------------------------
' Function: CheckTempTable
'---------------------------------------------------------------------------------------
'
' CheckTempTable
'
' Parameters:
'     SourceTableDefName   - Table name in TempDB (backend)
'     CreateTableDDL       - (optional) DDL statement to create the table if table does not exist.
'     LinkedTableDefName   - (optional) Table name in frontend (no linked table if empty)
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function CheckTable(ByVal SourceTableDefName As String, _
              Optional ByVal CreateTableSql As String = vbNullString, _
              Optional ByVal LinkedTableDefName As String = vbNullString) As Boolean

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
' Function: CreateTable
'---------------------------------------------------------------------------------------
'
' Erstellt Tabelle per DDL-Anweisung
'
' Parameters:
'     SourceTableDefName   - Table name in TempDB (backend)
'     CreateTableDDL       - DDL statement to create the table if table does not exist.
'     LinkedTableDefName   - (optional) Table name in frontend (no linked table if empty)
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function CreateTable(ByVal SourceTableDefName As String, _
                            ByVal CreateTableSql As String, _
                   Optional ByVal LinkedTableDefName As String = vbNullString _
                     ) As Boolean

   CheckCurrentDatabase
   If SourceTableDefExists(SourceTableDefName) Then 'delete for safety, if different DDL statement
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
' Function: DropTable
'---------------------------------------------------------------------------------------
'
' Deletes table
'
' Parameters:
'     SourceTableDefName   - Table name in TempDB (backend)
'     LinkedTableDefName   - (optional) Table name in frontend (no linked table if empty)
'
' Returns:
'     Boolean
'
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
' Sub: DropLinkedTableDefs
'---------------------------------------------------------------------------------------
'
' Removes TableDef objects linked to tables of TempDb
'
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
' Property: LinkedTableDefExists
'---------------------------------------------------------------------------------------
'
' Checks if table is present in the frontend
'
' Parameters:
'     TableDefName - Table name in frontend
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Property Get LinkedTableDefExists(ByVal TableDefName As String) As Boolean
   
   Dim db As DAO.Database
   Dim tdf As DAO.TableDef
   
   Set db = ClientDatabase
   db.TableDefs.Refresh 'for safety, if table was deleted in frontend
   
   LinkedTableDefExists = False
   
   For Each tdf In db.TableDefs
      If tdf.Name = TableDefName Then
         LinkedTableDefExists = True
         Exit For
      End If
   Next

End Property

'---------------------------------------------------------------------------------------
' Property: SourceTableDefExists
'---------------------------------------------------------------------------------------
'
' Checks if table exists in temp database (backend)
'
' Parameters:
'     TableDefName - Table name in backend
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Property Get SourceTableDefExists(ByVal TableDefName As String) As Boolean
   
   Dim tdf As DAO.TableDef

   SourceTableDefExists = False

   '1. call via CurrentDatabase, so that existence of the file is ensured
   Me.CurrentDatabase.TableDefs.Refresh
   
   For Each tdf In m_DaoHandler.CurrentDb.TableDefs
      If tdf.Name = TableDefName Then
         SourceTableDefExists = True
         Exit For
      End If
   Next

End Property

Private Sub DropLinkedTableDef(ByVal TdfName As String)

On Error Resume Next '... with it does not necessarily have to be checked for existence
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
   
   'Delete file again, because only name is needed
   Call Kill(TempFileName)
   
   If Len(FileExtension) > 0 Then 'Rewrite file extension
     TempFileName = Left$(TempFileName, Len(TempFileName) - 3) & FileExtension
   End If
   GetTempFileName = TempFileName

End Function


'#############################################################
' Group: Data Operations

'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'
' Open DAO.Recordset
'
' Parameters:
'     Source            - SQL statement or table or query name
'     RecordsetType     - DAO.RecordsetTypeEnum (Standard: dbOpenDynaset)
'     RecordsetOptions  - DAO.RecordsetOptionEnum (Standard: dbSeeChanges)
'     LockEdit          - DAO.LockTypeEnum (Standard: dbOptimistic)
'
' Returns:
'     DAO.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                       Optional ByVal RecordsetType As DAO.RecordsetTypeEnum = dbOpenForwardOnly, _
                       Optional ByVal RecordsetOptions As DAO.RecordsetOptionEnum = DAO.RecordsetOptionEnum.dbSeeChanges, _
                       Optional ByVal LockEdit As DAO.LockTypeEnum = DAO.LockTypeEnum.dbOptimistic) As DAO.Recordset
   Set OpenRecordset = m_DaoHandler.OpenRecordset(Source, RecordsetType, RecordsetOptions, LockEdit)
End Function

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'
' Execute SQL statement
'
' Parameters:
'     CommandText - SQL statement (String)
'     Options     - DAO.RecordsetOptionEnum
'
' Returns:
'     RecordsAffected (Long)
'
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal Query As String, _
          Optional ByVal Options As DAO.RecordsetOptionEnum) As Long
   Execute = m_DaoHandler.Execute(Query, Options)
End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'
' Lookup DAO.Recordset replacement function for DLookup (passing a SQL statement)
'
' Parameters:
'     SqlText     - SQL statement
'     Index       - Field selection if the first field of the recordset is not to be returned.
'                   The index is passed in the same way as DAO.Recordset.Fields(index).
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant  - Null, if SQL statement does not return a record.
'
' Remarks:
'     The SQL statement must be written in the Jet/ACE SQL dialect.
'
'---------------------------------------------------------------------------------------
Public Function LookupSql(ByVal SqlText As String, Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant
   LookupSql = m_DaoHandler.LookupSql(SqlText, Index, ValueIfNull)
End Function
