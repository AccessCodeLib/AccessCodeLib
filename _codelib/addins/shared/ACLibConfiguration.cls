VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACLibConfiguration"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Konfigurationseinstellungen der CodeLib verwalten"
'---------------------------------------------------------------------------------------
' Klassenmodul: ACLibConfiguration (Josef Pötzl, 2009-12-11)
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Konfigurationseinstellungen der CodeLib verwalten
' </summary>
' <remarks>
' </remarks>
'\ingroup addins_shared
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>_codelib/addins/shared/ACLibConfiguration.cls</file>
'  <license>_codelib/license.bas</license>
'  <use>base/ApplicationHandler.cls</use>
'  <use>data/SqlTools.bas</use>
'  <use>file/modFiles.bas</use>
'  <use>data/modSQL_Tools.bas</use>
'  <use>data/dao/modDAO_Tools.bas</use>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Database
Option Explicit

Private m_DaoSqlTools As SqlTools

Private Const m_EXTENSION_KEY As String = "ACLibConfiguration"

#Const ADODB_EARLYBINDING = 0
'ADODB wird hier über Late binding eingesetzt, da es nur zum Erstellen der Tabelle genutzt wird


' Grundeinstellungen
Private Const m_ACLIB_CONFIG_ROOTFOLDERNAME As String = "AccessCodeLibrary"
Private Const m_ACLIB_CONFIG_DATABASENAME As String = "ACLib_Config"
Private Const m_ACLIB_CONFIG_TABLEDEFNAME As String = "ACLib_ConfigTable"

Private Const m_SQL_CONFIG_TABLE_FIELD_PropName As String = "PropName"
Private Const m_SQL_CONFIG_TABLE_FIELD_PropValue As String = "PropValue"
Private Const m_SQL_SELECT_PropertyValue As String = _
      "select " & m_SQL_CONFIG_TABLE_FIELD_PropName & ", " & m_SQL_CONFIG_TABLE_FIELD_PropValue & _
      " From " & m_ACLIB_CONFIG_TABLEDEFNAME & " where " & m_SQL_CONFIG_TABLE_FIELD_PropName & " = [?]"
      
Private Const m_SQL_CREATETABLE_ConfigTable As String = _
            "CREATE TABLE " & m_ACLIB_CONFIG_TABLEDEFNAME & _
            "([PropName] varchar(255) WITH COMPRESSION NOT NULL," & _
            " [PropValue] varchar(255) WITH COMPRESSION," & _
            " [PropRemarks] text WITH COMPRESSION," & _
            " CONSTRAINT PK_" & m_ACLIB_CONFIG_TABLEDEFNAME & " PRIMARY KEY ([PropName]))"

Private Const m_PROPNAME_LocalRepositoryRoot As String = "LocalRepositoryRoot"
Private Const m_PROPNAME_PrivateRepositoryRoot As String = "PrivateRepositoryRoot"
Private Const m_PROPNAME_ImportTestDefaultValue As String = "ImportTestDefaultValue"

' Hilfsvariablen
Private m_LocalRepositoryPath As String   ' allg. Repository-Verzeichnis
Private m_PrivateRepositoryPath As String ' privates Verzeichnis (nicht in CodeLib enthalten)
Private m_ImportTestDefaultValue As Long
Private m_ACLibPropertyDb As DAO.Database

'---------------------------------------------------------------------------------------
' Standard-Initialisierung von Erweiterungen
'---------------------------------------------------------------------------------------

Private WithEvents m_ApplicationHandler As ApplicationHandler
Attribute m_ApplicationHandler.VB_VarHelpID = -1

Public Property Set ApplicationHandlerRef(ByRef objRef As ApplicationHandler)
   Set m_ApplicationHandler = objRef
End Property

Public Property Get ExtensionKey() As String
   ExtensionKey = m_EXTENSION_KEY
End Property

'---------------------------------------------------------------------------------------
' Standard-Ereignisbehandlung von Erweiterungen
'---------------------------------------------------------------------------------------

' CheckExtension
Private Sub m_ApplicationHandler_CheckExtension(ByVal sExtensionKey As String, ByRef exists As Boolean)
   If sExtensionKey = m_EXTENSION_KEY Then exists = True
End Sub

' ExtensionLookup
Private Sub m_ApplicationHandler_ExtensionLookup(ByVal sExtensionKey As String, ByRef ExtensionReference As Object)
   If sExtensionKey = m_EXTENSION_KEY Then
      Set ExtensionReference = Me
   End If
End Sub

'ExtensionPropertyLookup
Private Sub m_ApplicationHandler_ExtensionPropertyLookup( _
                     ByVal sExtensionKey As String, ByVal PropertyName As String, _
                     ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
   If sExtensionKey = m_EXTENSION_KEY Then
      getExtensionPropertyLookup PropertyName, ResumeMode, ResumeMessage
   End If
End Sub

' AfterDispose
Private Sub m_ApplicationHandler_AfterDispose(ResumeMode As ApplicationHandlerResumeModes, ResumeMessage As Variant)
'=> Referenz in m_ApplicationHandler auf Nothing setzen
   Set m_ApplicationHandler = Nothing
End Sub


'---------------------------------------------------------------------------------------
' Ergänzungen für Ereiterung: ACLibConfiguration
'---------------------------------------------------------------------------------------

Private Sub getExtensionPropertyLookup(ByVal PropertyName As String, ByRef ResumeMode As ApplicationHandlerResumeModes, ByRef ResumeMessage As Variant)

   ResumeMode = AppResumeMode_Completed
   
   Select Case PropertyName
      Case m_PROPNAME_LocalRepositoryRoot
         ResumeMessage = LocalRepositoryPath
         
      Case Else 'Property wurde nicht erkannt
         ResumeMode = AppResumeMode_Error
         
   End Select

End Sub

Public Property Get LocalRepositoryPath() As String

   If Len(m_LocalRepositoryPath) = 0 Then
      m_LocalRepositoryPath = getACLibGlobalProperty(m_PROPNAME_LocalRepositoryRoot)
      If Len(m_LocalRepositoryPath) > 0 Then
         If Not DirExists(m_LocalRepositoryPath) Then
            Err.Raise vbObjectError, "ACLibConfiguration.LocalRepositoryPath", "Das Verzeichnis '" & m_LocalRepositoryPath & "' ist nicht vorhanden!"
            m_LocalRepositoryPath = vbNullString
         End If
         If Right$(m_LocalRepositoryPath, 1) <> "\" Then
            m_LocalRepositoryPath = m_LocalRepositoryPath & "\"
            setACLibGlobalProperty m_PROPNAME_LocalRepositoryRoot, m_LocalRepositoryPath
         End If
      End If
   End If
   
   LocalRepositoryPath = m_LocalRepositoryPath

End Property

Public Property Let LocalRepositoryPath(ByVal newPath As String)

   If Len(newPath) > 0 Then
      If Right$(newPath, 1) <> "\" Then
         newPath = newPath & "\"
      End If
   End If
   m_LocalRepositoryPath = newPath
   setACLibGlobalProperty m_PROPNAME_LocalRepositoryRoot, m_LocalRepositoryPath
  
End Property


Public Property Get PrivateRepositoryPath() As String

   If Len(m_PrivateRepositoryPath) = 0 Then
      m_PrivateRepositoryPath = getACLibGlobalProperty(m_PROPNAME_PrivateRepositoryRoot)
      If Len(m_PrivateRepositoryPath) > 0 Then
         If Not DirExists(m_PrivateRepositoryPath) Then
            Err.Raise vbObjectError, "ACLibConfiguration.PrivateRepositoryPath", "Das Verzeichnis '" & m_PrivateRepositoryPath & "' ist nicht vorhanden!"
            m_PrivateRepositoryPath = vbNullString
         End If
         If Right$(m_PrivateRepositoryPath, 1) <> "\" Then
            m_PrivateRepositoryPath = m_PrivateRepositoryPath & "\"
            setACLibGlobalProperty m_PROPNAME_PrivateRepositoryRoot, m_PrivateRepositoryPath
         End If
      End If
   End If
   
   PrivateRepositoryPath = m_PrivateRepositoryPath

End Property

Public Property Let PrivateRepositoryPath(ByVal newPath As String)

   If Len(newPath) > 0 Then
      If Right$(newPath, 1) <> "\" Then
         newPath = newPath & "\"
      End If
   End If
   m_PrivateRepositoryPath = newPath
   setACLibGlobalProperty m_PROPNAME_PrivateRepositoryRoot, m_PrivateRepositoryPath
 
End Property

Public Property Get ImportTestsDefaultValue() As Boolean
'm_ImportTestDefaultValue: 0 = noch nicht abgefragt
'                          1 = false
'                          2 = true

   If m_ImportTestDefaultValue = 0 Then
      m_ImportTestDefaultValue = Val(getACLibGlobalProperty(m_PROPNAME_ImportTestDefaultValue)) + 1
   End If
   ImportTestsDefaultValue = (m_ImportTestDefaultValue = 2)

End Property

Public Property Let ImportTestsDefaultValue(ByVal NewValue As Boolean)

   m_ImportTestDefaultValue = Abs(NewValue) + 1
   setACLibGlobalProperty m_PROPNAME_ImportTestDefaultValue, Abs(NewValue)

End Property

Private Function getACLibGlobalProperty(ByRef PropertyName As String) As String
   
   Dim rst As DAO.Recordset
   Dim strSQL As String

   strSQL = Replace(m_SQL_SELECT_PropertyValue, "[?]", DaoSqlTool.TextToSqlText(PropertyName))
   Set rst = ACLibPropertyDb.OpenRecordset(strSQL)
   If Not rst.EOF Then
      getACLibGlobalProperty = Nz(rst.Fields(m_SQL_CONFIG_TABLE_FIELD_PropValue), vbNullString)
   Else
      getACLibGlobalProperty = vbNullString
   End If
   rst.Close

End Function

Private Function setACLibGlobalProperty(ByRef PropertyName As String, ByRef NewValue As String) As String
   
   Dim rst As DAO.Recordset
   Dim strSQL As String

   strSQL = Replace(m_SQL_SELECT_PropertyValue, "[?]", DaoSqlTool.TextToSqlText(PropertyName))
   Set rst = ACLibPropertyDb.OpenRecordset(strSQL)
   If rst.EOF Then
      rst.AddNew
      rst.Fields(m_SQL_CONFIG_TABLE_FIELD_PropName).Value = PropertyName
   Else
      rst.Edit
   End If
   rst.Fields(m_SQL_CONFIG_TABLE_FIELD_PropValue).Value = NewValue
   rst.Update
   rst.Close
 
End Function

Private Property Get ACLibPropertyDb() As DAO.Database

   If m_ACLibPropertyDb Is Nothing Then
      If checkConfigTableDef Then
         Set m_ACLibPropertyDb = CodeDb
      End If
   End If
   Set ACLibPropertyDb = m_ACLibPropertyDb
 
End Property

#If ADODB_EARLYBINDING Then
Private Function CreateConfigTable(ByRef cnn As ADODB.Connection) As Boolean
#Else
Private Function CreateConfigTable(ByRef cnn As Object) As Boolean
#End If
      
   cnn.Execute m_SQL_CREATETABLE_ConfigTable
   CreateConfigTable = True

End Function


Private Function checkConfigTableDef() As Boolean

   Dim db As DAO.Database
   Dim tdf As DAO.TableDef

   Set db = CodeDb
   
   If Not TableDefExists(m_ACLIB_CONFIG_TABLEDEFNAME, db) Then
   
      Set tdf = db.CreateTableDef(m_ACLIB_CONFIG_TABLEDEFNAME)
      tdf.Connect = ";Database=" & ACLibConfigDatabaseFile
      tdf.SourceTableName = m_ACLIB_CONFIG_TABLEDEFNAME
      db.TableDefs.Append tdf
      
   ElseIf Len(Dir$(Mid$(db.TableDefs(m_ACLIB_CONFIG_TABLEDEFNAME).Connect, Len(";Database=") + 1))) = 0 Then
   
      With db.TableDefs(m_ACLIB_CONFIG_TABLEDEFNAME)
         .Connect = ";Database=" & ACLibConfigDatabaseFile
         .RefreshLink
      End With
      
   End If
   
   Set db = Nothing
   
   checkConfigTableDef = True

End Function

Public Property Get ACLibConfigDirectory() As String
   
   Dim strPath As String

   strPath = Environ("Appdata") & "\" & m_ACLIB_CONFIG_ROOTFOLDERNAME & "\"
   If Len(Dir$(strPath, vbDirectory)) = 0 Then
      MkDir strPath
   End If
   
   ACLibConfigDirectory = strPath

End Property

Private Property Get ACLibConfigDatabaseFile() As String
   
   Dim db As DAO.Database
   Dim strDbFile As String
   Dim bolCreateConfigTable As Boolean
   
#If ADODB_EARLYBINDING = 1 Then
   Dim cnn As ADODB.Connection
#Else
   Dim cnn As Object
#End If

   strDbFile = CodeDb.Name
   strDbFile = Mid$(strDbFile, InStrRev(strDbFile, "."))
   If Left$(strDbFile, 5) = ".accd" Then
      strDbFile = ".accdu"
   Else
      strDbFile = ".mdt"
   End If
   strDbFile = ACLibConfigDirectory & m_ACLIB_CONFIG_DATABASENAME & strDbFile
   
   If Len(Dir$(strDbFile)) = 0 Then

      'Datenbank anlegen
      If CodeDb.Version = "4.0" Then
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral, dbVersion40)
      Else
         Set db = DBEngine.CreateDatabase(strDbFile, dbLangGeneral)
      End If
      db.Close
      
      bolCreateConfigTable = True
   
   Else 'Prüfen ob Config-Tabelle vorhanden ist
      
      Set db = DBEngine.OpenDatabase(strDbFile)
      bolCreateConfigTable = Not TableDefExists(m_ACLIB_CONFIG_TABLEDEFNAME, db)
      db.Close
   
   End If
   
   If bolCreateConfigTable Then
      'Tabelle erstellen
#If ADODB_EARLYBINDING = 1 Then
      Set cnn = New ADODB.Connection
#Else
      Set cnn = CreateObject("ADODB.Connection")
#End If
      cnn.ConnectionString = Replace(CodeProject.Connection.ConnectionString, CodeDb.Name, strDbFile)
      cnn.Open
      CreateConfigTable cnn
      cnn.Close
      Set cnn = Nothing
   End If

   ACLibConfigDatabaseFile = strDbFile

End Property

Private Property Get DaoSqlTool()
   If m_DaoSqlTools Is Nothing Then
      Set m_DaoSqlTools = SqlTools.Clone("\#yyyy-mm-dd\#", "True", "*")
   End If
   Set DaoSqlTool = m_DaoSqlTools
End Property

Private Sub Class_Terminate()
   Set m_DaoSqlTools = Nothing
End Sub
