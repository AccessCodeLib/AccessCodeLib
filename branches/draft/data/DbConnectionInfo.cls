VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbConnectionInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Klassenmodul: DbConnectionInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsklasse zur Einstellung der Verbindungsparameter
' </summary>
' <remarks></remarks>
'**/
'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionInfo.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

Private Enum DbConnectionInfoErrorNumbers
   ERRNR_NODBMS = vbObjectError + 1
   ERRNR_NOCONFIG = vbObjectError + 2
   ERRNR_INACTIVE = vbObjectError + 3
End Enum

Private Const m_LoginFormName As String = g_ApplicationLoginFormName

Private Const m_CheckConnectionTimeout As Long = 5 'Connectiontimeout für Verbindungsprüfung (Standard wäre 15 sec)

'Verbindungskennung
Private m_CID As String 'kann als Konstante verwendet werden, wenn nur eine Verbindungseinstellung genutzt werden soll

'DBMS-Bezeichnung
Private m_DBMSNAME As String 'Variable, damit mit mehreren DBMS gearbeitet werden kann

Private Const m_DBMSNAME_MSJET As String = "MSJET"
Private Const m_DBMSNAME_MSSQL As String = "MSSQL"
Private Const m_DBMSNAME_MYSQL As String = "MySQL"
Private Const m_DBMSNAME_ORACLE As String = "ORACLE"
Private Const m_DBMSNAME_ODBC As String = "ODBC"
                             
'DBMS-Provider/Driver
Private m_DBODBCDRIVER As String
Private m_DBOLEDBPROVIDER As String

'DB-Server (Anm.: inkl. Instanz beim SQL-Server):
Private m_DBSERVER As String

'Port (z. B. zur Unterscheidung vorhandener Server-Instanzen)
Private m_PORT As String

'DB-Name
Private m_DBNAME As String  'muss vorhanden sein!

'ConnectionMode (1 .. DNS less, 2 .. mit DSN, 128 .. benutzerdefinierter Connectionstring)
Private m_CONNECTIONMODE As Long 'Enum DbmsConnectionModes

'Benutzer-Kennung
Private m_DBUSER As String
Private m_DBUSERPWD As String
Private m_TRUSTEDCONNECTION As Boolean

'Weitere Optionen
Private m_DBOPTIONS_ODBC As String
Private m_DBOPTIONS_OLEDB As String

'ConnectionString
Private m_OledbConnectionString As String
Private m_OdbcConnectionString As String
Private m_DatabaseFile As String 'für dateibasierte DBS

'Login-Formular
Private m_UseLoginForm As Boolean
Private m_SAVEPWD As Boolean

Private WithEvents m_LoginForm As Form_frmLogin
Attribute m_LoginForm.VB_VarHelpID = -1
Private m_LoginCancelled As Boolean

'DSN
Private m_DSN As String

Private m_DbmsInfoReaded As Boolean

'Events
'/**
' <summary>
' Ereignis wenn sich die Verbindung ändert
' </summary>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="Source">Prozedur(en) in der der Fehler auftrat</param>
' <param name="ErrDescription">Fehlerbeschreibung</param>
' <param name="Response">AcUtilErrorResumeMode-Wert (ermöglich Fehlersteuerung als Reaktion auf das Ereignis)</param>
' <remarks>Kann nützlich sein, wenn in Abhängigkeit der Verbindung Anzeigen aktualisiert werden usw.</remarks>
'**/
Public Event DbmsConnectionChanged(ByVal newCID As String)

'Events
'/**
' <summary>
' Ereignis wenn sich die Verbindung ändert
' </summary>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="Source">Prozedur(en) in der der Fehler auftrat</param>
' <param name="ErrDescription">Fehlerbeschreibung</param>
' <param name="Response">AcUtilErrorResumeMode-Wert (ermöglich Fehlersteuerung als Reaktion auf das Ereignis)</param>
' <remarks>Kann nützlich sein, wenn in Abhängigkeit der Verbindung Anzeigen aktualisiert werden usw.</remarks>
'**/
Public Event DbmsConnectionInfoCleared()

'/**
' <summary>
' Ereignis zur Weitergabe von Informationen
' </summary>
' <param name="ErrNumber">Fehlernummer</param>
' <param name="Source">Prozedur(en) in der der Fehler auftrat</param>
' <param name="ErrDescription">Fehlerbeschreibung</param>
' <param name="Response">AcUtilErrorResumeMode-Wert (ermöglich Fehlersteuerung als Reaktion auf das Ereignis)</param>
' <remarks>Dient zur Weitergabe von Infos (ähnlich debug.print)</remarks>
'**/
Public Event DbmsInfoMessage(ByVal Source As String, ByVal Msg As String, ByVal MsgStyle As VbMsgBoxStyle)

''/**
'' <summary>
'' Ereignis für Fehlermeldungen
'' </summary>
'' <param name="ErrNumber">Fehlernummer</param>
'' <param name="Source">Prozedur(en) in der der Fehler auftrat</param>
'' <param name="ErrDescription">Fehlerbeschreibung</param>
'' <param name="Response">AcUtilErrorResumeMode-Wert (ermöglich Fehlersteuerung als Reaktion auf das Ereignis)</param>
'' <remarks>Ermöglich die Behebung bzw. die Behandlung des Fehlers</remarks>
''**/
'Public Event DbConnectionInfoError(ByVal ErrNumber As Long, ByVal Source As String, ByVal ErrDescription As String, ByRef Response As AcUtilErrorResumeMode)

Public Event PasswordChanged()
Public Event Disposed()


'API
Private Declare PtrSafe Function GetPrivateProfileString Lib "kernel32" _
  Alias "GetPrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpDefault As String, _
  ByVal lpReturnedString As String, _
  ByVal nSize As Long, _
  ByVal lpFileName As String) As Long

Private Declare PtrSafe Function WritePrivateProfileString Lib "kernel32" _
  Alias "WritePrivateProfileStringA" ( _
  ByVal lpApplicationName As String, _
  ByVal lpKeyName As Any, _
  ByVal lpString As Any, _
  ByVal lpFileName As String) As Long

Public Property Get CID() As String
On Error Resume Next
   CID = m_CID
End Property

Public Property Let CID(ByVal selectCID As String)
On Error Resume Next
   
   ClearConnectionInfo
   m_CID = selectCID
   
   If Len(m_CID) > 0 Then
      readDBMSprop RaiseDbmsConnectionChangedEvent:=True
   Else
      RaiseEvent DbmsConnectionChanged(vbNullString)
   End If
   
End Property

Public Property Get DbmsName() As String
On Error Resume Next
   If StrPtr(m_DBMSNAME) = 0 Then
      readDBMSprop
   End If
   DbmsName = m_DBMSNAME
End Property

Public Property Get DbServer() As String
On Error Resume Next
   DbServer = m_DBSERVER
End Property

Public Property Get DbName() As String
On Error Resume Next
   DbName = m_DBNAME
End Property

Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Private Sub Dispose()
On Error Resume Next
   If Not (m_LoginForm Is Nothing) Then Set m_LoginForm = Nothing
   m_DbmsInfoReaded = False
   RaiseEvent Disposed
End Sub

Private Sub readDBMSprop(Optional ByVal withoutLoginForm As Boolean = False, Optional ByVal RaiseDbmsConnectionChangedEvent As Boolean = False)

   Dim strSQL As String
   Dim rst As DAO.Recordset
   Dim bolCancel As Boolean

   Dim tempCID As String

   tempCID = m_CID

   ClearConnectionInfo
   
   If Len(m_CID) = 0 Then 'jenen DS auswählen, der als "Aktiv" gekennzeichnet wurde
      strSQL = "SELECT CID, DBMS, dbmsConnectionMode, dbmsDatabase, dbmsUseTrustedConnection, dbmsUser, dbmsPwd, dbmsUseLoginForm, " & _
                     " dbmsServer, dbmsPort, dbmsOleDbProvider, dbmsOdbcDriver, dbmsOptionsOLEDB, dbmsOptionsODBC," & _
                     " dbmsDSN, dbmsConStrOLEDB, dbmsConStrODBC FROM usys_DbmsConnection WHERE ActiveConnection=-1"
   Else
      strSQL = "SELECT CID, DBMS, dbmsConnectionMode, dbmsDatabase, dbmsUseTrustedConnection, dbmsUser, dbmsPwd, dbmsUseLoginForm, " & _
                     " dbmsServer, dbmsPort, dbmsOleDbProvider, dbmsOdbcDriver, dbmsOptionsOLEDB, dbmsOptionsODBC," & _
                     " dbmsDSN, dbmsConStrOLEDB, dbmsConStrODBC FROM usys_DbmsConnection WHERE CID=" & SqlTools.TextToSqlText(m_CID)
   End If
             
   Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
   If rst.EOF Then
      
      rst.Close
      
      If Len(m_CID) = 0 Then 'es wurde noch keine Verbindung aktiviert
         bolCancel = True
         If MsgBox("Es wurde noch keine Verbindung aktiviert." & vbNewLine & _
                     "Bitte überprüfen Sie die Verbindungsparameter und aktivieren Sie die gewünschte Verbindung.", _
                     vbInformation + vbOKCancel) = vbOK Then
            
            If Len(DbmsConfigFormName) > 0 Then
               DoCmd.OpenForm DbmsConfigFormName, , , , , acDialog
            End If
            
            'Abbruch, falls über Formular bereits initialisiert wurde
            If m_DbmsInfoReaded Then
               Exit Sub
            End If
            
            Set rst = CurrentDb.OpenRecordset(strSQL, dbOpenSnapshot)
            If Not rst.EOF Then
               m_CID = Nz(rst.Fields("CID"), vbNullString)
               bolCancel = (Len(m_CID) = 0)
            End If
         End If
         
         If bolCancel Then
            rst.Close
            Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.readDBMSprop", "Es wurde keine Verbindung aktiviert."
            Exit Sub
         End If
      Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.readDBMSprop", "Kein Eintrag für '" & m_CID & "' gefunden"
         Exit Sub
      End If
   End If
   
   
   'Einstellungen auslesen
   With rst
   
      If Len(m_CID) = 0 Then
         m_CID = Nz(rst.Fields("CID"), vbNullString)
         If Len(m_CID) = 0 Then
            Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.readDBMSprop", "Es wurde keine Verbindung übergeben."
            Exit Sub
         End If
      End If
      
      m_DBMSNAME = Nz(rst.Fields("DBMS"), vbNullString)
      m_CONNECTIONMODE = Nz(.Fields("dbmsConnectionMode"), 0) '1 = DNS less, 2 = DSN, 4 = benutzerdefiniert
   
      m_DBNAME = .Fields("dbmsDatabase") & vbNullString
      
      m_DBSERVER = .Fields("dbmsServer") & vbNullString
      m_PORT = .Fields("dbmsPort") & vbNullString
      
      m_TRUSTEDCONNECTION = Nz(.Fields("dbmsUseTrustedConnection"), False)
      m_UseLoginForm = Nz(.Fields("dbmsUseLoginForm"), False)
      m_DBUSER = .Fields("dbmsUser") & vbNullString
      m_DBUSERPWD = .Fields("dbmsPwd") & vbNullString
      
      'Falls kein Name angegeben ist, Loginform aktivieren
      If Not (m_UseLoginForm Or m_TRUSTEDCONNECTION) And Len(m_DBUSER) = 0 Then
         m_UseLoginForm = (m_DBMSNAME <> "MSJET")
      End If
      
      m_DSN = .Fields("dbmsDSN") & vbNullString
      
      m_DBOLEDBPROVIDER = .Fields("dbmsOleDbProvider") & vbNullString
      m_DBOPTIONS_OLEDB = .Fields("dbmsOptionsOLEDB") & vbNullString
      
      m_DBODBCDRIVER = .Fields("dbmsOdbcDriver") & vbNullString
      m_DBOPTIONS_ODBC = .Fields("dbmsOptionsODBC") & vbNullString
      
      m_OledbConnectionString = .Fields("dbmsConStrOLEDB") & vbNullString
      m_OdbcConnectionString = .Fields("dbmsConStrODBC") & vbNullString

      .Close
   End With
   Set rst = Nothing

   m_DbmsInfoReaded = True
   m_SAVEPWD = (Len(m_DBUSERPWD) > 0) And (m_UseLoginForm = False)
   
   If m_DBMSNAME = "MSJET" Then
      m_DatabaseFile = m_DBNAME
   End If

   If Not withoutLoginForm Then
   If m_UseLoginForm Then 'ergibt Anmeldung, daher nicht für Vorabeinstellungen nutzen! (=>withoutLoginForm=true)
   
      Dim LoginUserName As String
      LoginUserName = GetUserNameFromConfigFile
      If Len(LoginUserName) = 0 Then
         LoginUserName = m_DBUSER
      End If
   
      'm_LoginForm über Property LoginForm zugänglich machen und im LoginForm selbst setzen:
      Set g_TempRef = Me
      DoCmd.OpenForm m_LoginFormName, acNormal, , , acFormEdit, acDialog, m_CID & "@" & LoginUserName
      Set g_TempRef = Nothing
   End If
   End If
   
   If (tempCID <> m_CID) Or RaiseDbmsConnectionChangedEvent Then
      RaiseEvent DbmsConnectionChanged(m_CID)
   End If

End Sub

Private Function GetUserNameFromConfigFile() As String

   Dim ReturnStr As String * 1280
   Dim RetVal As Long
   
On Error Resume Next

   RetVal = GetPrivateProfileString("Login", "UserName", vbNullString & vbNullChar, ReturnStr, Len(ReturnStr), CurrentProject.Path & "\" & m_CID & ".ini")
   GetUserNameFromConfigFile = Trim(Left(ReturnStr, RetVal))

End Function

Private Sub SaveUserNameToConfigFile(ByVal UserName As String)
On Error Resume Next
   WritePrivateProfileString "Login", "UserName", UserName, CurrentProject.Path & "\" & m_CID & ".ini"
End Sub

Public Property Get LoginForm() As Form
On Error Resume Next
   Set LoginForm = m_LoginForm
End Property

Public Property Set LoginForm(ByRef ref As Form)
On Error Resume Next
   Set m_LoginForm = ref
End Property

Private Sub m_LoginForm_Cancelled()

   RaiseEvent DbmsInfoMessage(m_DBMSNAME, "Login wurde abgebrochen", vbExclamation)
   Set m_LoginForm = Nothing
   ClearConnectionInfo
   m_LoginCancelled = True

End Sub

Private Sub m_LoginForm_Logon(ByVal LoginName As String, ByVal Password As String, Cancel As Boolean)
   
   Dim ErrMsg As String

   m_LoginCancelled = False
   m_DBUSER = LoginName
   m_DBUSERPWD = Password
   
   Call setConnectionStrings(False)
   
   If CheckAdoConnection(m_OledbConnectionString, ErrMsg) = False Then
      If Len(ErrMsg) > 0 Then ErrMsg = vbNewLine & vbNewLine & ErrMsg
      MsgBox "Verbindung konnte nicht hergestellt werden" & ErrMsg
      m_LoginCancelled = True
      Cancel = True
      Exit Sub
   End If

   If Len(m_DBUSER) > 0 Then
      CurrentDb.Execute "UPDATE (select * from usys_DbmsConnectionWHERE DBMS='" & m_DBMSNAME & "') SET dbmsUser = '" & m_DBUSER & "'" ' thanks to MSFT (https://support.office.com/en-us/article/access-error-query-is-corrupt-fad205a5-9fd4-49f1-be83-f21636caedec)
      SaveUserNameToConfigFile m_DBUSER
   End If

   Set m_LoginForm = Nothing

End Sub

Private Sub setConnectionStrings(Optional ByVal bWithConnectionCheck As Boolean = True)

   Dim myConnectionStrings As DbmsConnectionStrings
   Dim ErrMsg As String

   If Not m_DbmsInfoReaded Then
      readDBMSprop
      If m_LoginCancelled Or ((Len(m_DBSERVER) + Len(m_DBNAME)) = 0) Then
         Exit Sub
      End If
   ElseIf m_LoginCancelled Then
      Exit Sub
   End If
   
   If m_CONNECTIONMODE = DbmsConnectionModes.DBMS_USERDEF Then
      'benutzerdefinierter String => kein Ermitteln über Dbms-Klasse erforderlich
      Exit Sub
   End If
   
   If Not m_TRUSTEDCONNECTION Then 'Verbindung mit User + Pwd

      m_TRUSTEDCONNECTION = False

      If Not m_DbmsInfoReaded Then
      If Len(m_DBUSER) = 0 Then
         readDBMSprop
      End If
      End If

   End If

   myConnectionStrings = Me.ConnectionStrings

   m_OledbConnectionString = myConnectionStrings.OledbConnectionString
   m_OdbcConnectionString = myConnectionStrings.OdbcConnectionString
   m_DatabaseFile = myConnectionStrings.DatabaseFile
   
   If bWithConnectionCheck Then
   If CheckConnection(ErrMsg) = False Then
      If Len(ErrMsg) > 0 Then ErrMsg = vbNewLine & vbNewLine & ErrMsg
      MsgBox "Verbindung konnte nicht hergestellt werden" & ErrMsg, vbExclamation
      ClearConnectionInfo
   End If
   End If

End Sub

'---------------------------------------------------------------------------------------
' Function: CheckConnection
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Verbindung zum Server bzw. Datenbank prüfen
' </summary>
' <param name="ErrMsg">Fehlermeldung vom ADODB- bzw. ODBC-Verbindungsaufbau</param>
' <returns>Boolean: True = Ado- und ODBC-Verbindung konnten aufgebaut werden</returns>
' <remarks>
' Als erster Test erfolgt der ADO-Verbindungsaufbau. Sobald dieser erfolgreich war, wird der ODBC-Verbindungsaufbau getestet.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function CheckConnection(Optional ByRef ErrMsg As String = vbNullString) As Boolean
   
   Dim checkOk As Boolean

   checkOk = CheckAdoConnection(OledbConnectionString, ErrMsg)
   If checkOk Then ' ODBC testen (außer ADO-Verbindung schlug bereits fehl)
      If m_DBMSNAME <> "MSJET" Then
         checkOk = checkOk And CheckOdbcConnection(OdbcConnectionString, ErrMsg)
      End If
   End If
   
   CheckConnection = checkOk

End Function

Private Function CheckAdoConnection(ByVal sOleDbConnectionstring As String, ByRef Msg As String) As Boolean

   Dim cnn As ADODB.Connection
   Dim i As Long

   If Len(sOleDbConnectionstring) = 0 Then
      CheckAdoConnection = False
      Exit Function
   End If
   
   Set cnn = New ADODB.Connection
   With cnn
      .ConnectionString = sOleDbConnectionstring
      .ConnectionTimeout = m_CheckConnectionTimeout
      
On Error Resume Next
      .Open
      If Err.Number <> 0 Then
         Msg = Msg & "Ursache (OLEDB):"
         For i = 0 To cnn.Errors.Count - 1
            Msg = Msg & vbNewLine & cnn.Errors(i).Description
         Next
         cnn.Errors.Clear
         Err.Clear
         CheckAdoConnection = False
         Exit Function
      End If
      
      If .State = adStateOpen Then
         .Close
         CheckAdoConnection = True
      Else
         CheckAdoConnection = False
      End If
   End With
   Set cnn = Nothing

End Function

Private Function CheckOdbcConnection(ByVal sOdbcConnectionstring As String, ByRef Msg As String) As Boolean

   Dim db As DAO.Database
   Dim checkOk As Boolean
   Dim i As Long

   If Len(sOdbcConnectionstring) = 0 Then
      CheckOdbcConnection = False
      Exit Function
   End If
   
On Error Resume Next

   Set db = DBEngine.OpenDatabase("", dbDriverComplete, True, sOdbcConnectionstring)
   If Err.Number <> 0 Then
      Msg = Msg & "Ursache (ODBC):"
      For i = 0 To DBEngine.Errors.Count - 1
         Msg = Msg & vbNewLine & DBEngine.Errors(i).Description
      Next
      Err.Clear
      checkOk = False
   Else
      checkOk = True
      db.Close
   End If
   Set db = Nothing
   
   CheckOdbcConnection = checkOk

End Function

'---------------------------------------------------------------------------------------
' Property: OledbConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OledbConnectionString
' </summary>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get OledbConnectionString() As String

   If Len(m_OledbConnectionString) = 0 Then
      setConnectionStrings
   End If
   OledbConnectionString = m_OledbConnectionString
   
End Property

'---------------------------------------------------------------------------------------
' Property: OdbcConnectionString
'---------------------------------------------------------------------------------------
'/**
' <summary>
' OdbcConnectionString
' </summary>
' <param name="bWithoutUserInfo">False = Benutzerkennung ist im ODBC-Connectionstring nethalten, True = ohne Benutzerkennung</param>
' <returns>String</returns>
' <remarks>Das Einstellen des Parameters bWithoutUserInfo ist normalerweise nicht erforderlich, kann jedoch hilfreich sein, falls der ODBC-Connectionstring ohne Passwort angezeigt werden soll.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get OdbcConnectionString(Optional ByVal bWithoutUserInfo As Boolean = False) As String

   If Len(m_OdbcConnectionString) = 0 Then
      setConnectionStrings
   End If
   If bWithoutUserInfo Then
      OdbcConnectionString = Replace(m_OdbcConnectionString, getOdbcUserInfo(), "")
   Else
      OdbcConnectionString = m_OdbcConnectionString
   End If

End Property

'---------------------------------------------------------------------------------------
' Property: DatabaseFile
'---------------------------------------------------------------------------------------
'/** 'nicht exportieren (2. * für Export fehlt)
' <summary>
' Dateiname inkl. Pfad eines dateibasierten Backend
' </summary>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DatabaseFile() As String

   If Len(m_DatabaseFile) = 0 Then
      setConnectionStrings
   End If
   DatabaseFile = m_DatabaseFile

End Property

'---------------------------------------------------------------------------------------
' Property: DbUser
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Name des Datenbankbenutzers
' </summary>
' <param name="withoutServerCheck">True = Keine Prüfung durchführen, ob Verbidnungsdaten bereits eingelesen wurden</param>
' <returns>String</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DbUser(Optional ByVal withoutServerCheck As Boolean = False) As String

   If Not withoutServerCheck Then
   If Not m_DbmsInfoReaded Then
      setConnectionStrings
   End If
   End If
   
   If m_TRUSTEDCONNECTION Then
      DbUser = getDbUserName
   Else
      DbUser = m_DBUSER
   End If

End Property

'---------------------------------------------------------------------------------------
' Property: DbUserPassword
'---------------------------------------------------------------------------------------
'/** 'nicht exportieren (2. * für Export fehlt)
' <summary>
' Passwort des Datenbankbenutzers, das bei Anmeldung eingegeben wurde
' </summary>
' <returns>String</returns>
' <remarks>
' Das Passwort sollte normalerweise nie abgefragt werden müssen und nur zur Weitergabe innerhalb der Klassenstruktur benutzt werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get DbUserPassword() As String

   If Not m_DbmsInfoReaded Then
      setConnectionStrings
   End If
   DbUserPassword = m_DBUSERPWD
 
End Property

'---------------------------------------------------------------------------------------
' Property: SavePassword
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Einstellung aus den Verbindungsparametern, ob Passwort bei verknüpften Tabellen gespeichert werden soll
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get SavePassword() As Boolean

   If Not m_DbmsInfoReaded Then
      setConnectionStrings
   End If
   SavePassword = m_SAVEPWD

End Property

'---------------------------------------------------------------------------------------
' Property: UseLoginForm
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Einstellung aus den Verbindungsparametern, ob beim Anmelden das Login-Formular verwendet werden soll
' </summary>
' <returns>Boolean</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get UseLoginForm() As Boolean

   If Not m_DbmsInfoReaded Then
      setConnectionStrings
   End If
   UseLoginForm = m_UseLoginForm

End Property

'---------------------------------------------------------------------------------------
' Sub: ClearConnectionInfo
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Eingestellte Verbindungsparameter löschen
' </summary>
' <remarks>
' Erfordert erneutes Einlesen aus Tabelle, wenn die Verbidnungsdaten angefordert werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Sub ClearConnectionInfo()

   Dim bolRaiseEvent As Boolean

On Error Resume Next

   bolRaiseEvent = (Len(m_DBMSNAME & m_OledbConnectionString & m_OdbcConnectionString) > 0)
   
   m_DBMSNAME = vbNullString
   m_OledbConnectionString = vbNullString
   m_OdbcConnectionString = vbNullString
   m_DatabaseFile = vbNullString
   m_DBOLEDBPROVIDER = vbNullString
   m_DBODBCDRIVER = vbNullString
   m_DBSERVER = vbNullString
   m_PORT = vbNullString
   m_DBNAME = vbNullString
   m_DBUSER = vbNullString
   m_DBUSERPWD = vbNullString
   m_DBOPTIONS_ODBC = vbNullString
   m_DBOPTIONS_OLEDB = vbNullString
   m_TRUSTEDCONNECTION = False
   m_UseLoginForm = False
   m_LoginCancelled = False
   m_DbmsInfoReaded = False
   
   If bolRaiseEvent Then 'Ereignis nur beim ersten Leeren auslösen
      RaiseEvent DbmsConnectionInfoCleared
   End If
   
End Sub

'---------------------------------------------------------------------------------------
' Function: ChangeDbUserPassword
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Datenbankpasswort des Benutzers ändern
' </summary>
' <param name="NewPassword">Neues Passwort</param>
' <param name="OldPassword">Altes Passwort</param>
' <param name="LoginName">Benutzername (Falls kein Benutzername übergeben wird, wird der aktuelle Benutzer eingesetzt)</param>
' <param name="Msg">ByRef-Parameter zur Rückgabe einer Fehlermeldung, falls das Passwort nicht geändert werden könnte</param>
' <returns>Boolean: True = Passwort wurde geändert</returns>
' <remarks>
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ChangeDbUserPassword( _
               ByVal NewPassword As String, _
               ByVal OldPassword As String, _
               Optional ByVal LoginName As String = vbNullString, _
               Optional ByRef Msg As String) As Boolean
   
   Dim bolPwdChanged As Boolean
   Dim strSQL As String
   Dim strLoginName As String
   Dim cnn As ADODB.Connection

   Set cnn = New ADODB.Connection
   cnn.Open OledbConnectionString
   
   If Len(LoginName) > 0 Then
      strLoginName = LoginName
   Else
      strLoginName = m_DBUSER
   End If

   strSQL = "ALTER LOGIN " & strLoginName & " WITH PASSWORD = '" & NewPassword & "'"
   
   If Len(OldPassword) > 0 Then
      strSQL = strSQL & " OLD_PASSWORD = '" & OldPassword & "'"
   End If
   
   On Error Resume Next
   With cnn
      .Errors.Clear
      .Execute strSQL
      If .Errors.Count > 0 Then
         bolPwdChanged = False
         Msg = .Errors(0).Description
         .Errors.Clear
      Else
         bolPwdChanged = True
      End If
      .Close
   End With
   Set cnn = Nothing
   
   If bolPwdChanged Then
   
      If strLoginName = m_DBUSER Then
         m_DBUSERPWD = NewPassword
         setConnectionStrings
         RaiseEvent PasswordChanged
      End If
      
   End If
   
   ChangeDbUserPassword = bolPwdChanged

End Function

'#################################
'
' DMBS-spezifische Prozeduren
'

'---------------------------------------------------------------------------------------
' Property: PermanentBackendRstSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Select-Anweisung für ODBC-Recordsetverbindung zum Aufrechterhalten der Verbindung
' </summary>
' <returns>String</returns>
' <remarks>
' Diese Prozedur muss erweitert werden, wenn weitere DBMS (wie Oracle) eingesetzt werden sollen.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get PermanentBackendRstSqlText() As String
On Error Resume Next
   PermanentBackendRstSqlText = "select 1"
End Property

Private Function getDbUserName() As String

   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim strSQL As String

   Set cnn = New ADODB.Connection
   cnn.Open Me.OledbConnectionString

   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSSQL
         strSQL = "select current_user"
      Case m_DBMSNAME_MYSQL
         strSQL = "select current_user"
      Case m_DBMSNAME_ORACLE
         strSQL = "select user from dual"
      Case Else
         On Error GoTo 0
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.DbUserName", "Für '" & m_DBMSNAME & "' kann der UserName nicht ermittelt werden. (Auswahl nicht möglich.)"
   End Select
   
   Set rst = cnn.Execute(strSQL)
   
   getDbUserName = Nz(rst.Fields(0), "N/A")
   rst.Close
   cnn.Close

End Function

'---------------------------------------------------------------------------------------
' Property: ConnectionStrings
'---------------------------------------------------------------------------------------
'/**
' <summary>
' ODBC- und OLEDB-Connectionstring
' </summary>
' <returns>DbmsConnectionStrings</returns>
' <remarks>
' Diese Prozedur muss erweitert werden, wenn weitere DBMS (wie Oracle) eingesetzt werden sollen.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get ConnectionStrings() As DbmsConnectionStrings

   If m_DbmsInfoReaded = False Then
      readDBMSprop
   End If

   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSSQL
         ConnectionStrings = ConnectionStrings_MSSQL
      Case m_DBMSNAME_MSJET
         ConnectionStrings = ConnectionStrings_MSJET
      Case m_DBMSNAME_MYSQL
         ConnectionStrings = ConnectionStrings_MySQL
      Case m_DBMSNAME_ORACLE
         ConnectionStrings = ConnectionStrings_ORACLE
      Case m_DBMSNAME_ODBC
         ConnectionStrings = ConnectionStrings_ODBC
      Case Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "ConnectionStrings", "Für '" & m_DBMSNAME & "' können keine Connectionstrings erzeugt werden. (Auswahl nicht möglich.)"
   End Select

End Property

Private Function getOdbcUserInfo(Optional ByVal sUID As String = vbNullString, _
                                 Optional ByVal sPWD As String = vbNullString) As String
   
   If Len(sUID) = 0 Then
      If m_DbmsInfoReaded = False Then
         readDBMSprop
      End If
      sUID = m_DBUSER
      sPWD = m_DBUSERPWD
   End If
   
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSSQL
         getOdbcUserInfo = getOdbcUserInfo_MSSQL(sUID, sPWD)
      Case m_DBMSNAME_MYSQL
         getOdbcUserInfo = getOdbcUserInfo_MySQL(sUID, sPWD)
      Case m_DBMSNAME_ORACLE
         getOdbcUserInfo = getOdbcUserInfo_ORACLE(sUID, sPWD)
      Case Else
         On Error GoTo 0
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.getOdbcUserInfo", "Für '" & m_DBMSNAME & "' kann keine OdbcUserInfo erzeugt werden. (Auswahl nicht möglich.)"
   End Select

End Function

Private Function ConnectionStrings_MSJET() As DbmsConnectionStrings

   Dim tempConnectionStrings As DbmsConnectionStrings
   Dim strODBC As String, strOLEDB As String
   Dim strDbFile As String

   strDbFile = m_DBNAME

   If Len(m_DBOPTIONS_OLEDB) > 0 Then
      strOLEDB = strOLEDB & ";" & m_DBOPTIONS_OLEDB
   End If
   If Len(m_DBOPTIONS_ODBC) > 0 Then
      strODBC = strODBC & ";" & m_DBOPTIONS_ODBC
   End If

   If Len(m_DBOLEDBPROVIDER) = 0 Then
      m_DBOLEDBPROVIDER = CurrentProject.Connection.Provider
   End If

   strOLEDB = "Provider=" & m_DBOLEDBPROVIDER & ";Data Source=" & strDbFile & ";Mode=Share Deny None;" & strOLEDB

   strODBC = ";DATABASE=" & strDbFile & strODBC
   
   tempConnectionStrings.OledbConnectionString = strOLEDB
   tempConnectionStrings.OdbcConnectionString = strODBC
   tempConnectionStrings.DatabaseFile = strDbFile
   
   ConnectionStrings_MSJET = tempConnectionStrings

End Function

Private Function ConnectionStrings_MSSQL() As DbmsConnectionStrings

   Dim tempConnectionStrings As DbmsConnectionStrings
   Dim strODBC As String, strOLEDB As String
   Dim strSrvCon As String

   strSrvCon = m_DBSERVER
   If Len(m_PORT) > 0 Then
      strSrvCon = strSrvCon & "," & m_PORT
   End If
   
   If m_TRUSTEDCONNECTION Then
      strODBC = ";Trusted_Connection=Yes"
      strOLEDB = ";Integrated Security = SSPI"
   ElseIf Len(m_DBUSER) > 0 Then 'Verbindung mit User + Pwd
      strODBC = ";Trusted_Connection=No;" & getOdbcUserInfo_MSSQL(m_DBUSER, m_DBUSERPWD)
      strOLEDB = ";User ID=" & m_DBUSER & ";Password=""" & m_DBUSERPWD & """"
   End If

   If Len(m_DBOPTIONS_OLEDB) > 0 Then
      strOLEDB = strOLEDB & ";" & m_DBOPTIONS_OLEDB
   End If
   If Len(m_DBOPTIONS_ODBC) > 0 Then
      strODBC = strODBC & ";" & m_DBOPTIONS_ODBC
   End If

   If Len(m_DBOLEDBPROVIDER) = 0 Then
      m_DBOLEDBPROVIDER = "SQLOLEDB"
   End If

   strOLEDB = "Provider=" & m_DBOLEDBPROVIDER & ";Data Source=" & strSrvCon & ";Initial Catalog=" & m_DBNAME & strOLEDB
   'strOLEDB = GetDataProviderOledbString(strOLEDB)

   If Len(m_DSN) > 0 Then
      strODBC = "DSN=" & m_DSN & strODBC
   Else
      If Len(m_DBODBCDRIVER) = 0 Then
        m_DBODBCDRIVER = "SQL Server"
      End If
      strODBC = "DRIVER={" & m_DBODBCDRIVER & "};DATABASE=" & m_DBNAME & ";SERVER=" & strSrvCon & strODBC
   End If
   
   tempConnectionStrings.OledbConnectionString = strOLEDB
   tempConnectionStrings.OdbcConnectionString = "ODBC;" & strODBC
   
   ConnectionStrings_MSSQL = tempConnectionStrings

End Function

Private Function GetDataProviderOledbString(ByVal BaseOledbString As String) As String
'Provider=Microsoft.Access.OLEDB.10.0;Data Provider=SQLOLEDB
   
   GetDataProviderOledbString = Replace(BaseOledbString, "Provider=", "Provider=Microsoft.Access.OLEDB.10.0;Data Provider=")

End Function

Private Function getOdbcUserInfo_MSSQL(sUID As String, sPWD As String) As String
On Error GoTo 0
   getOdbcUserInfo_MSSQL = "UID=" & sUID & ";PWD={" & sPWD & "}"
End Function

Private Function ConnectionStrings_MySQL() As DbmsConnectionStrings

   Dim tempConnectionStrings As DbmsConnectionStrings
   Dim strODBC As String
   Dim strSrvCon As String

   If m_TRUSTEDCONNECTION Then 'Wie würde das bei MySQL funktionieren?
      m_TRUSTEDCONNECTION = False
      If setLoginData Then
         strODBC = getOdbcUserInfo_MySQL(m_DBUSER, m_DBUSERPWD) & ";"
      Else
         On Error GoTo 0
         Err.Raise ERRNR_INACTIVE, "ConnectionStrings_MySQL", "Der Windowsbenutzer wurde für die Datenbank nicht freigegeben!"
      End If
   Else 'Verbindung mit User + Pwd
      strODBC = getOdbcUserInfo_MySQL(m_DBUSER, m_DBUSERPWD) & ";"
   End If

   strSrvCon = m_DBSERVER
   If Len(m_PORT) > 0 Then
     strSrvCon = strSrvCon & ";PORT=" & m_PORT
   End If

   If Len(m_DBOPTIONS_ODBC) > 0 Then
     strODBC = strODBC & m_DBOPTIONS_ODBC & ";"
   End If
   
   If Len(m_DSN) > 0 Then
      strODBC = "DSN=" & m_DSN & ";" & strODBC
   Else
      If Len(m_DBODBCDRIVER) = 0 Then
        m_DBODBCDRIVER = "MySQL ODBC 5.1 Driver"
      End If
      strODBC = "DRIVER={" & m_DBODBCDRIVER & "};DATABASE=" & m_DBNAME & ";SERVER=" & strSrvCon & ";" & strODBC
   End If
   
   tempConnectionStrings.OdbcConnectionString = "ODBC;" & strODBC
   tempConnectionStrings.OledbConnectionString = "Provider=MSDASQL;" & strODBC

   ConnectionStrings_MySQL = tempConnectionStrings

End Function

Private Function getOdbcUserInfo_MySQL(sUID As String, sPWD As String) As String
On Error GoTo 0
   getOdbcUserInfo_MySQL = "UID=" & sUID & ";PWD=" & sPWD
End Function

Private Function ConnectionStrings_ORACLE() As DbmsConnectionStrings

   Dim tempConnectionStrings As DbmsConnectionStrings
   Dim strODBC As String, strOLEDB As String
   Dim strSrvCon As String

   If m_TRUSTEDCONNECTION Then 'Wie funktioiniert das bei Oracle?
     strODBC = "Trusted_Connection=Yes;"
     strOLEDB = "Integrated Security = SSPI;"
   Else 'Verbindung mit User + Pwd
     strODBC = getOdbcUserInfo_ORACLE(m_DBUSER, m_DBUSERPWD) & ";"
     strOLEDB = "User ID=" & m_DBUSER & ";Password=" & m_DBUSERPWD & ";"
   End If
   
   If Len(m_DBOPTIONS_OLEDB) > 0 Then
     strOLEDB = strOLEDB & m_DBOPTIONS_OLEDB & ";"
   End If
   If Len(m_DBOPTIONS_ODBC) > 0 Then
     strODBC = strODBC & m_DBOPTIONS_ODBC & ";"
   End If
   
   'Microsoft ODBC for Oracle:
   'strSrvCon = "(DESCRIPTION=(ADDRESS=(COMMUNITY=192.168.0)(PROTOCOL=TCP)(Host=" & m_DBSERVER & ")(Port=" & m_PORT & "))(CONNECT_DATA=(SID=" & m_DBNAME & ")))"
   strSrvCon = "(DESCRIPTION=(ADDRESS=(COMMUNITY=)(PROTOCOL=TCP)(Host=" & m_DBSERVER & ")(Port=" & m_PORT & "))(CONNECT_DATA=(SID=" & m_DBNAME & ")))"
      
   If Len(m_DBOLEDBPROVIDER) = 0 Then
      m_DBOLEDBPROVIDER = "OraOLEDB.Oracle"
   End If
   strOLEDB = "Provider=" & m_DBOLEDBPROVIDER & ";Data Source=" & strSrvCon & ";" & strOLEDB

   If Len(m_DSN) > 0 Then
      strODBC = "DSN=" & m_DSN & ";" & strODBC
   Else
      If Len(m_DBODBCDRIVER) = 0 Then
        m_DBODBCDRIVER = "Microsoft ODBC for Oracle" '"Oracle in OraDb10g_home1" oder "Microsoft ODBC for Oracle"
        'sDBODBCDRIVER = "Oracle in OraDb10g_home1"
      End If
      strODBC = "DRIVER={" & m_DBODBCDRIVER & "};SERVER=" & strSrvCon & ";" & strODBC
   End If
   
   tempConnectionStrings.OledbConnectionString = strOLEDB
   tempConnectionStrings.OdbcConnectionString = "ODBC;" & strODBC
   
   ConnectionStrings_ORACLE = tempConnectionStrings

End Function

Private Function getOdbcUserInfo_ORACLE(sUID As String, sPWD As String) As String
On Error GoTo 0
   getOdbcUserInfo_ORACLE = "UID=" & sUID & ";PWD=" & sPWD
End Function

Private Function ConnectionStrings_ODBC() As DbmsConnectionStrings

   Dim tempConnectionStrings As DbmsConnectionStrings
   Dim strODBC As String
   Dim strSrvCon As String
   
   If Len(m_DBODBCDRIVER) = 0 Then
      Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NOCONFIG, "ConnectionStrings_ODBC", "Es wurde kein ODBC-Treiber angegeben"
   End If

   If m_TRUSTEDCONNECTION Then
     strODBC = "Trusted_Connection=Yes;"
   Else 'Verbindung mit User + Pwd
      strODBC = getOdbcUserInfo_ODBC(m_DBUSER, m_DBUSERPWD) & ";"
   End If

   strSrvCon = m_DBSERVER
   If Len(m_PORT) > 0 Then
     strSrvCon = strSrvCon & ";PORT=" & m_PORT
   End If

   If Len(m_DBOPTIONS_ODBC) > 0 Then
     strODBC = strODBC & m_DBOPTIONS_ODBC & ";"
   End If
   
   If Len(m_DSN) > 0 Then
      strODBC = "DSN=" & m_DSN & ";" & strODBC
   Else
      strODBC = "DRIVER={" & m_DBODBCDRIVER & "};DATABASE=" & m_DBNAME & ";SERVER=" & strSrvCon & ";" & strODBC
   End If
   
   tempConnectionStrings.OdbcConnectionString = "ODBC;" & strODBC
   tempConnectionStrings.OledbConnectionString = "Provider=MSDASQL;" & strODBC
   
   ConnectionStrings_ODBC = tempConnectionStrings
End Function

Private Function getOdbcUserInfo_ODBC(sUID As String, sPWD As String) As String
   getOdbcUserInfo_ODBC = "UID=" & sUID & ";PWD=" & sPWD
End Function

'---------------------------------------------------------------------------------------
' Property: TableListSqlText
'---------------------------------------------------------------------------------------
'/**
' <summary>
' SQL-Anweisung im SQL-Dialtekt des DBMS zur Auflistung der im DBS enthaltenen Tabellen
' </summary>
' <returns>String</returns>
' <remarks>
' Diese Prozedur muss erweitert werden, wenn weitere DBMS (wie Oracle) eingesetzt werden sollen.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Property Get TableListSqlText() As String
   
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSSQL
         TableListSqlText = "select T.TABLE_NAME, T.TABLE_SCHEMA, cast(T.TABLE_SCHEMA + '.' + T.TABLE_NAME as varchar(255)) AS TabKey from information_schema.TABLES T Order By T.TABLE_NAME, T.TABLE_SCHEMA"
      Case m_DBMSNAME_MYSQL
         TableListSqlText = "SELECT T.TABLE_NAME, NULL AS TABLE_SCHEMA, T.TABLE_NAME AS TabKey FROM information_schema.`TABLES` T WHERE T.TABLE_SCHEMA='" & DbName & "' Order By T.TABLE_NAME"
      Case m_DBMSNAME_ORACLE
         TableListSqlText = "SELECT T.TNAME AS TABLE_NAME, NULL AS TABLE_SCHEMA, T.TNAME AS TabKey FROM TAB T Order By T.TNAME"
      Case m_DBMSNAME_MSJET
         TableListSqlText = "select T.NAME AS TABLE_NAME, NULL AS TABLE_SCHEMA, T.NAME AS TabKey from [" & Me.DatabaseFile & "].MSysObjects T WHERE Type=1 AND (T.Name < 'Msys' OR T.Name >= 'Msyt') Order By T.NAME"
      Case m_DBMSNAME_ODBC
         TableListSqlText = vbNullString
      Case vbNullString
         TableListSqlText = vbNullString
      Case Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.TableListSQL", "Für '" & m_DBMSNAME & "' können keine Connectionstrings erzeugt werden. (Auswahl nicht möglich.)"
   End Select

End Property

Private Function setLoginData() As Boolean
   
   Dim cnn As ADODB.Connection
   Dim cmd As ADODB.Command
   Dim rst As ADODB.Recordset
   
   Dim strWindowUser As String

   strWindowUser = Environ("username")
   
   Set cnn = New ADODB.Connection
   cnn.Open ConnectionStrings_MySQL.OledbConnectionString
   
   Set cmd = New ADODB.Command
   Set cmd.ActiveConnection = cnn
   cmd.CommandText = "GetLoginData"
   cmd.CommandType = adCmdStoredProc
   cmd.Parameters.Append cmd.CreateParameter("windowsuser", adVarChar, adParamInput, 255)
   cmd.Parameters(0) = strWindowUser
   Set rst = cmd.Execute
   If rst.EOF Then
      m_DBUSER = strWindowUser
      m_DBUSERPWD = vbNullString
      setLoginData = False
   Else
      m_DBUSER = rst.Fields(0)
      m_DBUSERPWD = rst.Fields(1)
      setLoginData = True
   End If
   
   rst.Close
   Set cmd = Nothing
   cnn.Close

End Function

'---------------------------------------------------------------------------------------
' Function: ParseTableName
'---------------------------------------------------------------------------------------
'/**
' <summary>
' Hilfsprozedur zum Anpassen des Tabellennamens
' </summary>
' <param name="sTabName">Tabellename</param>
' <returns>String</returns>
' <remarks>
' Bei manchen DBMS muss Groß- o. Kleinschreibung beachtet bzw. der Name mit Sonderzeichen eingeklammert werden.
' </remarks>
'**/
'---------------------------------------------------------------------------------------
Public Function ParseTableName(ByVal sTabName As String) As String
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSJET
         ParseTableName = "[" & sTabName & "]"
      Case m_DBMSNAME_MSSQL 'keine Änderung erforderlich - zur Sicherheint in Anführugnszeichen einklammern, damit auch  Tabellennamen mit Leerzeichen verwendet werden können
         ParseTableName = """" & sTabName & """"
      Case m_DBMSNAME_MYSQL 'zur Sicherheit in "`" einklammern.
         ParseTableName = "`" & sTabName & "`"
      Case m_DBMSNAME_ORACLE
         ParseTableName = """" & sTabName & """"
      Case Else
         ParseTableName = sTabName
   End Select
End Function

Public Property Get DateFormat() As String
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSJET
         DateFormat = "\#yyyy-mm-dd hh:nn:ss\#"
      Case m_DBMSNAME_MSSQL
         DateFormat = "'yyyymmdd hh:nn:ss'"
      Case m_DBMSNAME_MYSQL
         DateFormat = "'yyyy-mm-dd hh:nn:ss'"
      Case m_DBMSNAME_ORACLE
         DateFormat = "'yyyymmdd hh:nn:ss'"
      Case m_DBMSNAME_ODBC
         DateFormat = "'yyyymmdd hh:nn:ss'"
      Case Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.DateFormat", "Für '" & m_DBMSNAME & "' kann keine Datumsfromat ermittelt werden."
   End Select
End Property

Public Property Get WildCardString() As String
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSJET
         WildCardString = "*"
      Case m_DBMSNAME_MSSQL
         WildCardString = "%"
      Case m_DBMSNAME_MYSQL
         WildCardString = "%"
      Case m_DBMSNAME_ORACLE
         WildCardString = "%"
      Case m_DBMSNAME_ODBC
         WildCardString = "%"
      Case Else
         WildCardString = "%"
   End Select
End Property

Public Property Get BooleanTrueString() As String
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSJET
         BooleanTrueString = "True"
      Case m_DBMSNAME_MSSQL
         BooleanTrueString = "1"
      Case m_DBMSNAME_MYSQL
         BooleanTrueString = "1"
      Case m_DBMSNAME_ORACLE
         BooleanTrueString = "1"
      Case m_DBMSNAME_ODBC
         BooleanTrueString = "True"
      Case Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.BooleanTrueString", "Für '" & m_DBMSNAME & "' kann kein True-String ermittelt werden."
   End Select
End Property
