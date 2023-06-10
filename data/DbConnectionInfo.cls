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
' Class: data.DbConnectionInfo
'---------------------------------------------------------------------------------------
' Auxiliary class for setting the connection parameters
'
' Author:
'     Josef Poetzl
'
' Remarks:
'     Uses table usys_DbmsConnection (create table with DBMS Connection Wizard)
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/DbConnectionInfo.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>DAO</name><major>5</major><minor>0</minor><guid>{00025E01-0000-0000-C000-000000000046}</guid></ref>
'  <use>data/defDbConnection.bas</use>
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

Private Const ConnDataSelectSql As String = _
                     "SELECT CID, DBMS, dbmsConnectionMode, dbmsDatabase, dbmsUseTrustedConnection, dbmsUser, dbmsPwd, dbmsUseLoginForm, " & _
                     " dbmsServer, dbmsPort, dbmsOleDbProvider, dbmsOdbcDriver, dbmsOptionsOLEDB, dbmsOptionsODBC," & _
                     " dbmsDSN, dbmsConStrOLEDB, dbmsConStrODBC FROM usys_DbmsConnection"

Private Const ActivatedConnDataSelectSql As String = _
                     ConnDataSelectSql & " WHERE ActiveConnection=-1"

Private Const SelectConnectionIdentifierConnDataSelectSql As String = _
                     ConnDataSelectSql & " WHERE CID={ConnectionIdentifier}"


Private Const m_DefaultLoginFormName As String = "LoginForm"
Private m_LoginFormName As String
Private m_DbmsConfigFormName As String

Private Const m_CheckConnectionTimeoutSeconds As Long = 5 'Connection timeout [sec] for connection check (default would be 15 sec)

'Connection identifier
Private m_ConnectionIdentifier As String

'DBMS Name
Private m_DBMSNAME As String

Private Const m_DBMSNAME_MSJET As String = "MSJET"
Private Const m_DBMSNAME_MSSQL As String = "MSSQL"
Private Const m_DBMSNAME_MYSQL As String = "MySQL"
Private Const m_DBMSNAME_ORACLE As String = "ORACLE"
Private Const m_DBMSNAME_ODBC As String = "ODBC"
                             
'DBMS Provider/Driver
Private m_DBODBCDRIVER As String
Private m_DBOLEDBPROVIDER As String

'Database Server (Note: incl. instance of SQL server & co.)
Private m_DBSERVER As String

'Port (e.g. to identify existing server instances)
Private m_PORT As String

'Database Name
Private m_DBNAME As String  'required!

'Connection modes (1 .. DNS less, 2 .. with DSN, 128 .. usedefined connection string)
Private m_CONNECTIONMODE As DbmsConnectionModes

'User identifier
Private m_DBUSER As String
Private m_DBUSERPWD As String
Private m_TRUSTEDCONNECTION As Boolean

'Other Options
Private m_DBOPTIONS_ODBC As String
Private m_DBOPTIONS_OLEDB As String

'Connection string
Private m_OledbConnectionString As String
Private m_OdbcConnectionString As String
Private m_DatabaseFile As String 'für dateibasierte DBS

'Login form
Private m_UseLoginForm As Boolean
Private m_SAVEPWD As Boolean

#If CONFIG_USELOGINFORM = 1 Then
Private WithEvents m_LoginForm As Form_LoginForm
Attribute m_LoginForm.VB_VarHelpID = -1
#End If

Private m_LoginCancelled As Boolean

'DSN
Private m_DSN As String

'Other auxiliary variables:
Private m_DbmsInfoReaded As Boolean

'###################################################
'
' Group: Events
'

'---------------------------------------------------------------------------------------
'Event: DbmsConnectionChanged
'---------------------------------------------------------------------------------------
'
' Event when the connection changes
'
' Parameters:
'     NewConnectionIdentifier - new used connection identifier (from table data field CID)
'
' Remarks:
'     Can be useful if displays are updated depending on the connection, etc.
'
'---------------------------------------------------------------------------------------
Public Event DbmsConnectionChanged(ByVal NewConnectionIdentifier As String)

'---------------------------------------------------------------------------------------
'Event: DbmsConnectionInfoCleared
'---------------------------------------------------------------------------------------
'
' Event raised when the connection parameters have been reset
'
'---------------------------------------------------------------------------------------
Public Event DbmsConnectionInfoCleared()

'---------------------------------------------------------------------------------------
'Event: DbmsInfoMessage
'---------------------------------------------------------------------------------------
'
' Information sharing event
'
' Parameters:
'     Source         - Source of the message
'     Msg            - Message
'     MsgStyle       - VbMsgBoxStyle
'
' Remarks:
'     Used to pass on info (similar to debug.print)
'
'---------------------------------------------------------------------------------------
Public Event DbmsInfoMessage(ByVal Source As String, ByVal Msg As String, ByVal MsgStyle As VbMsgBoxStyle)

'---------------------------------------------------------------------------------------
'Event: PasswordChanged
'---------------------------------------------------------------------------------------
'
' Event raised when the password has been changed
'
'---------------------------------------------------------------------------------------
Public Event PasswordChanged()

'---------------------------------------------------------------------------------------
'Event: Disposed
'---------------------------------------------------------------------------------------
'
' Event raised when disposed
'
'---------------------------------------------------------------------------------------
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
  
Private Sub Class_Terminate()
On Error Resume Next
   Dispose
End Sub

Private Sub Dispose()
On Error Resume Next
#If CONFIG_USELOGINFORM = 1 Then
   If Not (m_LoginForm Is Nothing) Then Set m_LoginForm = Nothing
#End If
   m_DbmsInfoReaded = False
   RaiseEvent Disposed
End Sub

  
'###################################################
'
' Group: DBMS Connection Parameters
'
  
'---------------------------------------------------------------------------------------
'Property: ConnectionIdentifier
'---------------------------------------------------------------------------------------
'
' Connection identifier - value from data field CID
'
'---------------------------------------------------------------------------------------
Public Property Get ConnectionIdentifier() As String
On Error Resume Next
   ConnectionIdentifier = m_ConnectionIdentifier
End Property

Public Property Let ConnectionIdentifier(ByVal ConnectionIdentifierToUse As String)
On Error Resume Next
   
   ClearConnectionInfo
   m_ConnectionIdentifier = ConnectionIdentifierToUse
   
   If Len(m_ConnectionIdentifier) > 0 Then
      LoadConnectionData RaiseDbmsConnectionChangedEvent:=True
   Else
      RaiseEvent DbmsConnectionChanged(vbNullString)
   End If
   
End Property

Public Property Get CID() As String
   Stop 'Deprecated: use ConnectionIdentifier
   CID = ConnectionIdentifier
End Property

Public Property Let CID(ByVal SelectCID As String)
   Stop 'Deprecated: use ConnectionIdentifier
   ConnectionIdentifier = SelectCID
End Property


'---------------------------------------------------------------------------------------
'Property: DbmsName
'---------------------------------------------------------------------------------------
'
' Name of DBMS (MSJET, MSSQL, MySQL, ORACLE, ODBC)
'
'---------------------------------------------------------------------------------------
Public Property Get DbmsName() As String
On Error Resume Next
   If StrPtr(m_DBMSNAME) = 0 Then
      LoadConnectionData
   End If
   DbmsName = m_DBMSNAME
End Property

'---------------------------------------------------------------------------------------
'Property: DbServer
'---------------------------------------------------------------------------------------
'
' Name of Database Server
'
' Remarks:
'     Incl. instance of SQL Server & co.
'
'---------------------------------------------------------------------------------------
Public Property Get DbServer() As String
On Error Resume Next
   DbServer = m_DBSERVER
End Property

'---------------------------------------------------------------------------------------
'Property: DbName
'---------------------------------------------------------------------------------------
'
' Name of Database
'
'---------------------------------------------------------------------------------------
Public Property Get DbName() As String
On Error Resume Next
   DbName = m_DBNAME
End Property

'---------------------------------------------------------------------------------------
' Property: DatabaseFile
'---------------------------------------------------------------------------------------
'
' File name incl. path of a file-based backend
'
'---------------------------------------------------------------------------------------
Public Property Get DatabaseFile() As String

   If Len(m_DatabaseFile) = 0 Then
      SetConnectionStrings
   End If
   DatabaseFile = m_DatabaseFile

End Property

'---------------------------------------------------------------------------------------
' Property: UseLoginForm
'---------------------------------------------------------------------------------------
'
' Setting from the connection parameters whether to use the login form when logging in.
'
'---------------------------------------------------------------------------------------
Public Property Get UseLoginForm() As Boolean

   If Not m_DbmsInfoReaded Then
      SetConnectionStrings
   End If
   UseLoginForm = m_UseLoginForm

End Property


#If CONFIG_USELOGINFORM = 1 Then

'---------------------------------------------------------------------------------------
'Property: LoginForm
'---------------------------------------------------------------------------------------
'
' Form reference to LoginForm
'
'---------------------------------------------------------------------------------------
Public Property Get LoginForm() As Form
On Error Resume Next
   Set LoginForm = m_LoginForm
End Property

Public Property Set LoginForm(ByVal NewRef As Form)
On Error Resume Next
   Set m_LoginForm = NewRef
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
   
   Call SetConnectionStrings(False)
   
   If CheckAdoConnection(m_OledbConnectionString, ErrMsg) = False Then
      If Len(ErrMsg) > 0 Then ErrMsg = vbNewLine & vbNewLine & ErrMsg
      MsgBox "Verbindung konnte nicht hergestellt werden" & ErrMsg
      m_LoginCancelled = True
      Cancel = True
      Exit Sub
   End If

   If Len(m_DBUSER) > 0 Then
      CurrentDb.Execute "UPDATE (select * from usys_DbmsConnection WHERE DBMS='" & m_DBMSNAME & "') SET dbmsUser = '" & m_DBUSER & "'" ' thanks to MSFT (https://support.office.com/en-us/article/access-error-query-is-corrupt-fad205a5-9fd4-49f1-be83-f21636caedec)
      SaveUserNameToConfigFile m_DBUSER
   End If

   Set m_LoginForm = Nothing

End Sub

#End If

Private Sub SetConnectionStrings(Optional ByVal bWithConnectionCheck As Boolean = True)

   Dim myConnectionStrings As DbmsConnectionStrings
   Dim ErrMsg As String

   If Not m_DbmsInfoReaded Then
      LoadConnectionData
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
         LoadConnectionData
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
' Property: DbUser
'---------------------------------------------------------------------------------------
'
' Name des Datenbankbenutzers
'
' Parameters:
'     WithoutServerCheck   -  True = Do not check whether connection data has already been read in.
'
'---------------------------------------------------------------------------------------
Public Property Get DbUser(Optional ByVal WithoutServerCheck As Boolean = False) As String

   If Not WithoutServerCheck Then
   If Not m_DbmsInfoReaded Then
      SetConnectionStrings
   End If
   End If
   
   If m_TRUSTEDCONNECTION Then
      DbUser = GetDbUserName
   Else
      DbUser = m_DBUSER
   End If

End Property

'---------------------------------------------------------------------------------------
' Property: DbUserPassword
'---------------------------------------------------------------------------------------
'
' Password of the database user
'
' Remarks:
'     The password should normally never need to be requested and should only be used for sharing within the class structure.
'
'---------------------------------------------------------------------------------------
Private Property Get DbUserPassword() As String

   If Not m_DbmsInfoReaded Then
      SetConnectionStrings
   End If
   DbUserPassword = m_DBUSERPWD
 
End Property

Friend Property Let DbUserPassword(ByVal NewValue As String)

   m_DBUSERPWD = NewValue
 
End Property

'---------------------------------------------------------------------------------------
' Property: SavePassword
'---------------------------------------------------------------------------------------
'
' Setting from the connection parameters whether password should be saved for linked tables
'
'---------------------------------------------------------------------------------------
Public Property Get SavePassword() As Boolean

   If Not m_DbmsInfoReaded Then
      SetConnectionStrings
   End If
   SavePassword = m_SAVEPWD

End Property

'###################################################
'
' Group: DBMS Connection
'

'---------------------------------------------------------------------------------------
' Property: ConnectionStrings
'---------------------------------------------------------------------------------------
'
' ODBC and OLEDB Connectionstring
'
' Returns:
'     <data.defDbConnection::DbmsConnectionStrings>
'
' Remarks:
'     This procedure must be extended if other DBMSs (such as Oracle) are to be used.
'
'---------------------------------------------------------------------------------------
Friend Property Get ConnectionStrings() As DbmsConnectionStrings

   If m_DbmsInfoReaded = False Then
      LoadConnectionData
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


'---------------------------------------------------------------------------------------
' Property: OledbConnectionString
'---------------------------------------------------------------------------------------
'
' OledbConnectionString
'
' Returns:
'     String
'
'---------------------------------------------------------------------------------------
Friend Property Get OledbConnectionString() As String

   If Len(m_OledbConnectionString) = 0 Then
      SetConnectionStrings
   End If
   OledbConnectionString = m_OledbConnectionString
   
End Property

'---------------------------------------------------------------------------------------
' Property: OdbcConnectionString
'---------------------------------------------------------------------------------------
'
' OdbcConnectionString
'
' Parameters:
'     WithoutUserInfo - False = user ID incl. password is contained in the ODBC connection string,
'                       True = without user data
'
' Returns:
'     String
'
' Remarks:
'     Setting the parameter WithoutUserInfo is normally not necessary,
'     but can be helpful if the ODBC connection string is required without a password.
'
'---------------------------------------------------------------------------------------
Friend Property Get OdbcConnectionString(Optional ByVal WithoutUserInfo As Boolean = False) As String

   If Len(m_OdbcConnectionString) = 0 Then
      SetConnectionStrings
   End If
   If WithoutUserInfo Then
      OdbcConnectionString = Replace(m_OdbcConnectionString, GetOdbcUserInfo(), "")
   Else
      OdbcConnectionString = m_OdbcConnectionString
   End If

End Property

'---------------------------------------------------------------------------------------
' Function: CheckConnection
'---------------------------------------------------------------------------------------
'
' Check connection to server and database
'
' Parameters:
'     ErrMsg      - Error message from ADODB or ODBC connection
'
' Returns:
'     Boolean     - True = Ado- and ODBC-connection could be established
'
' Remarks:
'     The first test is to establish the ADO connection.
'     As soon as this is successful, the ODBC connection is tested.
'
'---------------------------------------------------------------------------------------
Public Function CheckConnection(Optional ByRef ErrMsg As String = vbNullString) As Boolean
   
   Dim checkOk As Boolean

   checkOk = CheckAdoConnection(OledbConnectionString, ErrMsg)
   If checkOk Then ' test ODBC testen (if ADO connection don't failed)
      If m_DBMSNAME <> "MSJET" Then
         checkOk = checkOk And CheckOdbcConnection(OdbcConnectionString, ErrMsg)
      End If
   End If
   
   CheckConnection = checkOk

End Function

Private Function CheckAdoConnection(ByVal sOleDbConnectionstring As String, ByRef Msg As String) As Boolean

   Dim cnn As Object 'ADODB.Connection
   Dim i As Long

   If Len(sOleDbConnectionstring) = 0 Then
      CheckAdoConnection = False
      Exit Function
   End If
   
   Set cnn = CreateObject("ADODB.Connection")
   With cnn
      .ConnectionString = sOleDbConnectionstring
      .ConnectionTimeout = m_CheckConnectionTimeoutSeconds
      
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
' Sub: ClearConnectionInfo
'---------------------------------------------------------------------------------------
'
' Delete set connection parameters
'
' Remarks:
'     Requires re-reading from table when connection data is requested.
'
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
   
   If bolRaiseEvent Then
      RaiseEvent DbmsConnectionInfoCleared
   End If
   
End Sub

'---------------------------------------------------------------------------------------
' Function: ChangeDbUserPassword
'---------------------------------------------------------------------------------------
'
' Datenbankpasswort des Benutzers ändern
'
' Parameters:
'     NewPassword - New passwort
'     OldPassword - old passwort
'     LoginName   - Login name (If no login name is passed, the current login name is used).
'     Msg         - ByRef parameter to return an error message if the password could not be changed
'
' Returns:
'     Boolean     - True = Password has been changed
'
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
         SetConnectionStrings
         RaiseEvent PasswordChanged
      End If
      
   End If
   
   ChangeDbUserPassword = bolPwdChanged

End Function

'#################################
'
' Group: DMBS specific
'

'---------------------------------------------------------------------------------------
' Property: PermanentBackendRstSqlText
'---------------------------------------------------------------------------------------
'
' Select statement for ODBC recordset connection to maintain connection
'
' Remarks:
'     This procedure must be extended if other DBMSs (such as Oracle) are to be used.
'
'---------------------------------------------------------------------------------------
Public Property Get PermanentBackendRstSqlText() As String
On Error Resume Next
   PermanentBackendRstSqlText = "select 1"
End Property

Private Function GetDbUserName() As String

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
   
   GetDbUserName = Nz(rst.Fields(0), "N/A")
   rst.Close
   cnn.Close

End Function

Private Function GetOdbcUserInfo(Optional ByVal sUID As String = vbNullString, _
                                 Optional ByVal sPWD As String = vbNullString) As String
   
   If Len(sUID) = 0 Then
      If m_DbmsInfoReaded = False Then
         LoadConnectionData
      End If
      sUID = m_DBUSER
      sPWD = m_DBUSERPWD
   End If
   
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSSQL
         GetOdbcUserInfo = getOdbcUserInfo_MSSQL(sUID, sPWD)
      Case m_DBMSNAME_MYSQL
         GetOdbcUserInfo = getOdbcUserInfo_MySQL(sUID, sPWD)
      Case m_DBMSNAME_ORACLE
         GetOdbcUserInfo = getOdbcUserInfo_ORACLE(sUID, sPWD)
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
      If SetLoginData Then
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
'
' SQL-Anweisung im SQL-Dialtekt des DBMS zur Auflistung der im DBS enthaltenen Tabellen
'
' Remarks:
'     This procedure must be extended if further DBMS are to be used.
'
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

Private Function SetLoginData() As Boolean
   
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
      SetLoginData = False
   Else
      m_DBUSER = rst.Fields(0)
      m_DBUSERPWD = rst.Fields(1)
      SetLoginData = True
   End If
   
   rst.Close
   Set cmd = Nothing
   cnn.Close

End Function

'---------------------------------------------------------------------------------------
' Function: ParseTableName
'---------------------------------------------------------------------------------------
'
' Auxiliary procedure for adjusting the table name
'
' Parameters:
'     TableName  -  Table name to parse
'
' Returns:
'     Table name for the specific DBMS
'
' Remarks:
'     With some DBMS, the name must be case-sensitive or bracketed with special characters.
'
'---------------------------------------------------------------------------------------
Public Function ParseTableName(ByVal TableName As String) As String
   Select Case m_DBMSNAME
      Case m_DBMSNAME_MSJET
         ParseTableName = "[" & TableName & "]"
      Case m_DBMSNAME_MSSQL ' No change required - for safety, enclose in inverted commas so that table names with spaces can also be used.
         ParseTableName = """" & TableName & """"
      Case m_DBMSNAME_MYSQL 'bracket in "`" for safety.
         ParseTableName = "`" & TableName & "`"
      Case m_DBMSNAME_ORACLE
         ParseTableName = """" & TableName & """"
      Case Else
         ParseTableName = TableName
   End Select
End Function

'---------------------------------------------------------------------------------------
' Property: DateFormat
'---------------------------------------------------------------------------------------
'
' DBMS specific date format
'
' Returns:
'     Format string
'
'---------------------------------------------------------------------------------------
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

'---------------------------------------------------------------------------------------
' Property: WildCardString
'---------------------------------------------------------------------------------------
'
' DBMS specific wildcard (like *)
'
'---------------------------------------------------------------------------------------
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

'---------------------------------------------------------------------------------------
' Property: BooleanTrueString
'---------------------------------------------------------------------------------------
'
' DBMS specific string for true
'
'---------------------------------------------------------------------------------------
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

#If CONFIG_USELOGINFORM = 1 Then
'---------------------------------------------------------------------------------------
' Property: LoginFormName
'---------------------------------------------------------------------------------------
'
' Name of the login form
'
' Remarks:
'     Example from ACLib: user/login/LoginForm.frm
'
'---------------------------------------------------------------------------------------
Public Property Get LoginFormName() As String
   If Len(m_LoginFormName) = 0 Then
      LoginFormName = m_DefaultLoginFormName
   Else
      LoginFormName = m_LoginFormName
   End If
End Property

Public Property Let LoginFormName(ByVal NewValue As String)
    m_LoginFormName = NewValue
End Property

#End If

'---------------------------------------------------------------------------------------
' Property: DbmsConfigFormName
'---------------------------------------------------------------------------------------
'
' Name of the DBMS configuration form
'
'---------------------------------------------------------------------------------------
Public Property Get DbmsConfigFormName() As String
   DbmsConfigFormName = m_DbmsConfigFormName
End Property

Public Property Let DbmsConfigFormName(ByVal NewValue As String)
    m_DbmsConfigFormName = NewValue
End Property


Private Sub LoadConnectionData(Optional ByVal WithoutLoginForm As Boolean = False, Optional ByVal RaiseDbmsConnectionChangedEvent As Boolean = False)
   
   Dim tempConnID As String

   tempConnID = m_ConnectionIdentifier

   ClearConnectionInfo
   
   LoadConnectionDataFromTable
   
   m_DbmsInfoReaded = True
   m_SAVEPWD = (Len(m_DBUSERPWD) > 0) And (m_UseLoginForm = False)
   
   If m_DBMSNAME = "MSJET" Then
      m_DatabaseFile = m_DBNAME
   End If

   If Not WithoutLoginForm Then
   If m_UseLoginForm Then 'ergibt Anmeldung, daher nicht für Vorabeinstellungen nutzen! (=>withoutLoginForm=true)
#If CONFIG_USELOGINFORM = 1 Then
      Dim LoginUserName As String
      LoginUserName = GetUserNameFromConfigFile
      If Len(LoginUserName) = 0 Then
         LoginUserName = m_DBUSER
      End If
   
      'm_LoginForm über Property LoginForm zugänglich machen und im LoginForm selbst setzen:
      DoCmd.OpenForm LoginFormName, acNormal, , , acFormEdit, acDialog, ObjPtr(Me) & "@" & m_ConnectionIdentifier & "@" & LoginUserName
#Else
   Err.Raise vbObjectError, "DbConnectionInfo", "CONFIG_USELOGINFORM in Projekteigenschaften nicht gesetzt!"
#End If
   End If
   End If
   
   If (tempConnID <> m_ConnectionIdentifier) Or RaiseDbmsConnectionChangedEvent Then
      RaiseEvent DbmsConnectionChanged(m_ConnectionIdentifier)
   End If

End Sub

Private Sub LoadConnectionDataFromTable()

   Dim rst As DAO.Recordset
   Dim SelectSql As String
   Dim Loa
   
   If Len(m_ConnectionIdentifier) = 0 Then 'jenen DS auswählen, der als "Aktiv" gekennzeichnet wurde
      SelectSql = ActivatedConnDataSelectSql
   Else
      SelectSql = Replace(SelectConnectionIdentifierConnDataSelectSql, "{ConnectionIdentifier}", SqlTools.TextToSqlText(m_ConnectionIdentifier))
   End If

   Set rst = CurrentDb.OpenRecordset(SelectSql, dbOpenSnapshot)
   If rst.EOF Then
      
      rst.Close
      Set rst = Nothing
      
      If Len(m_ConnectionIdentifier) = 0 Then 'es wurde noch keine Verbindung aktiviert
         If Not TryGetConnectionDataWithConfigForm(rst) Then
            Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.LoadDbmsConnectionDataFromTable", "Es wurde keine Verbindung aktiviert."
            Exit Sub
         ElseIf m_DbmsInfoReaded Then
            Exit Sub
         End If
      Else
         Err.Raise DbConnectionInfoErrorNumbers.ERRNR_NODBMS, "DbConnectionInfo.LoadDbmsConnectionDataFromTable", "Kein Eintrag für '" & m_ConnectionIdentifier & "' gefunden"
         Exit Sub
      End If
      
   End If
   
   'Read settings
   LoadConnectionDataFromRecordset rst
   rst.Close
   Set rst = Nothing

End Sub

Private Function TryGetConnectionDataWithConfigForm(ByRef rst As DAO.Recordset) As Boolean

   If Len(DbmsConfigFormName) > 0 Then
   
      If MsgBox("Es wurde noch keine Verbindung aktiviert." & vbNewLine & _
                     "Bitte überprüfen Sie die Verbindungsparameter und aktivieren Sie die gewünschte Verbindung.", _
                     vbInformation + vbOKCancel) = vbOK Then
                     
         DoCmd.OpenForm DbmsConfigFormName, , , , , acDialog
         'Abbruch, falls über Formular bereits initialisiert wurde
         If m_DbmsInfoReaded Then
            TryGetConnectionDataWithConfigForm = True
            Exit Function
         End If
         
      End If
      
      Set rst = CurrentDb.OpenRecordset(ActivatedConnDataSelectSql, dbOpenSnapshot)
      If Not rst.EOF Then
         m_ConnectionIdentifier = Nz(rst.Fields("CID"), vbNullString)
         TryGetConnectionDataWithConfigForm = True
      Else
         rst.Close
         Set rst = Nothing
         TryGetConnectionDataWithConfigForm = False
      End If
      
   End If
   
End Function

Private Sub LoadConnectionDataFromRecordset(ByVal rst As DAO.Recordset)

   With rst
      If Len(m_ConnectionIdentifier) = 0 Then
         m_ConnectionIdentifier = Nz(rst.Fields("CID"), vbNullString)
         If Len(m_ConnectionIdentifier) = 0 Then
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

   End With
   
End Sub

Private Function GetUserNameFromConfigFile() As String

   Dim ReturnStr As String * 1280
   Dim RetVal As Long
   
On Error Resume Next

   RetVal = GetPrivateProfileString("Login", "UserName", vbNullString & vbNullChar, ReturnStr, Len(ReturnStr), CurrentProject.Path & "\" & m_ConnectionIdentifier & ".ini")
   GetUserNameFromConfigFile = Trim(Left(ReturnStr, RetVal))

End Function

Private Sub SaveUserNameToConfigFile(ByVal UserName As String)
On Error Resume Next
   WritePrivateProfileString "Login", "UserName", UserName, CurrentProject.Path & "\" & m_ConnectionIdentifier & ".ini"
End Sub
