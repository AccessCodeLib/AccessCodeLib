VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdodbHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Class: data.adodb.AdodbHandler
'---------------------------------------------------------------------------------------
'
' ADODB data connection methods
'
' Author:
'     Josef Poetzl
'
'---------------------------------------------------------------------------------------

'---------------------------------------------------------------------------------------
'<codelib>
'  <file>data/ado/AdodbHandler.cls</file>
'  <license>_codelib/license.bas</license>
'  <ref><name>ADODB</name><major>2</major><minor>1</minor><guid>{00000201-0000-0010-8000-00AA006D2EA4}</guid></ref>
'  <test>_test/data/ado/AdodbHandlerTests.cls</test>
'</codelib>
'---------------------------------------------------------------------------------------
'
Option Compare Text
Option Explicit

'ADODB-Connection
Private m_AdoConnection As ADODB.Connection

'ADO-Connectionstring
Private m_AdoConnectionString As String

'#############################################################
' Group: Events

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
'
' Parameters:
'     Msg                  - Message about the cause of the event triggering. (Not yet used)
'     NewConnectionString  - OLEDB Connectionstring
'
' Remarks:
'     The appropriate OLEDB connection string can be passed on via the ByRef parameter NewConnectionString.
'
'---------------------------------------------------------------------------------------
Public Event ErrorEmptyConnectionString(ByVal Msg As String, ByRef NewConnectionString As String)

'---------------------------------------------------------------------------------------
' Event: HookInsertIdentityReturn
'---------------------------------------------------------------------------------------
'
' Event for missing OLEDB connection string
'
' Parameters:
'     InsertSql            - SQL text of insert statement
'     IdentityTable        - Identity table
'     ActiveConnection     - ADODB Connection to use
'     IdentitySelectString - Input string to use
'     HookResponseMode       - resume action after event call: 0 ... code resume after event, 1 exit function after event
'     HookResponseMessage    - return value of function (e.g. InsertIdentityReturn) when HookResumeMode <> 0
'
' Remarks:
'     New values can be passed on via the ByRef parameters.
'
'---------------------------------------------------------------------------------------
Public Event HookInsertIdentityReturn(ByRef InsertSQL As String, ByRef IdentityTable As String, _
                                      ByRef ActiveConnection As ADODB.Connection, ByRef IdentitySelectString As String, _
                                      ByRef HookResponseMode As Long, ByRef HookResponseMessage As Variant)

Private m_HooksEnabled As Boolean

Private m_DefaultCursorLocation As ADODB.CursorLocationEnum

'#############################################################
' Group: Class control

'
' Init / Terminate
' ----------------
Private Sub Class_Initialize()
   m_DefaultCursorLocation = adUseServer
End Sub

Private Sub Class_Terminate()
   Dispose
End Sub

Public Sub Dispose()
   Set m_AdoConnection = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Property: HooksEnabled
'---------------------------------------------------------------------------------------
' Active Hooks
'
' Remarks:
'     Used in InsertIdentityReturn
'
'---------------------------------------------------------------------------------------
Public Property Get HooksEnabled() As Boolean
   HooksEnabled = m_HooksEnabled
End Property

Public Property Let HooksEnabled(ByVal NewValue As Boolean)
   m_HooksEnabled = NewValue
End Property

Public Property Get DefaultCursorLocation() As ADODB.CursorLocationEnum
   DefaultCursorLocation = m_DefaultCursorLocation
End Property

Public Property Let DefaultCursorLocation(ByVal NewValue As ADODB.CursorLocationEnum)
   m_DefaultCursorLocation = NewValue
End Property

'#############################################################
' Group: Data Connection

'---------------------------------------------------------------------------------------
' Property: CurrentConnection
'---------------------------------------------------------------------------------------
'
' Get:
'     Currently set ADODB connection
'
' Set:
'     Set new ADODB connection
'
' Remarks:
'     If there is no ADODB connection yet, a new ADODB connection will be created.
'     Instantiation flow:
'        1. raiseEvent #ErrorMissingCurrentConnection ... if no connection was passed:
'        2. create new ADODB connection based on the OLEDB connection string (use property <data.adodb.AdodbHandler::ConnectionString>)
'
'     If a reference to an ADODB connection already exists, it will be checked if it is closed and opened if necessary.
'
'---------------------------------------------------------------------------------------
Public Property Get CurrentConnection() As ADODB.Connection

   Dim strConnString As String

   If (m_AdoConnection Is Nothing) Then
      'Request connection via event.
      RaiseEvent ErrorMissingCurrentConnection("Unavailable CurrentConnection", m_AdoConnection)
      If (m_AdoConnection Is Nothing) Then
         'Set new connection if no connection was passed by event
         strConnString = ConnectionString
         If Len(strConnString) > 0 Then
            Set m_AdoConnection = New ADODB.Connection
            m_AdoConnection.Open strConnString
         End If
      End If
   ElseIf (m_AdoConnection.State <> ObjectStateEnum.adStateOpen) Then
      If Len(m_AdoConnection.ConnectionString) > 0 Then
         m_AdoConnection.Open
      Else '(re)open connection
         m_AdoConnection.Open ConnectionString
      End If
   End If
   Set CurrentConnection = m_AdoConnection

End Property

Public Property Set CurrentConnection(ByRef NewConnection As ADODB.Connection)
   If Not (m_AdoConnection Is Nothing) Then
      If m_AdoConnection.State <> adStateClosed Then
         m_AdoConnection.Close
      End If
      Set m_AdoConnection = Nothing
   End If
   Set m_AdoConnection = NewConnection
End Property

'---------------------------------------------------------------------------------------
' Property: ConnectionString
'---------------------------------------------------------------------------------------
'
' OLEDB connection string
'
'---------------------------------------------------------------------------------------
Public Property Get ConnectionString() As String
   If Len(m_AdoConnectionString) = 0 Then
      If Not (m_AdoConnection Is Nothing) Then
         m_AdoConnectionString = m_AdoConnection.ConnectionString
      End If
      If Len(m_AdoConnectionString) = 0 Then
         RaiseEvent ErrorEmptyConnectionString("AdodbHandler.ConnectionString", m_AdoConnectionString)
      End If
   End If
   ConnectionString = m_AdoConnectionString
End Property

Public Property Let ConnectionString(ByVal Value As String)
   Set CurrentConnection = Nothing
   m_AdoConnectionString = Value
End Property

'#############################################################
' Group: Data Operations

'---------------------------------------------------------------------------------------
' Function: Execute
'---------------------------------------------------------------------------------------
'
' Execute SQL statement
'
' Parameters:
'     CommandText       - SQL statement
'     RecordsAffected   - ByRef parameter to return the number of records affected by the Execute statement.
'     Options           - additional options of the ADODB.Connection.Execute method
'
' Returns:
'     ADODB.Recordset
'
' Remarks:
'     Encapsulation of the ADODB.Connection.Execute method (parameters are passed on 1:1)
'
'---------------------------------------------------------------------------------------
Public Function Execute(ByVal CommandText As String, _
               Optional ByRef RecordsAffected As Long, _
               Optional ByVal Options As ADODB.ExecuteOptionEnum = -1) As ADODB.Recordset
   Set Execute = Me.CurrentConnection.Execute(CommandText, RecordsAffected, Options)
End Function

'---------------------------------------------------------------------------------------
' Function: ExecuteCommand
'---------------------------------------------------------------------------------------
'
' Execute sql statement using ADODB.Command
'
' Parameters:
'     CmdText           - ADODB.CommandText value
'     CmdType           - ADODB.CommandType value
'     CmdParamDefs      - 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value)
'     RecordsAffected   - ByRef parameter to return the number of records affected by the Execute statement.
'     Options           - additional options of the ADODB.Command.Execute method
'
' Returns:
'     ADODB.Recordset
'
' Remarks:
'     With this variant, SQL injection can be avoided very easily,
'     since the parameters are not built directly into the SQL statement
'     and thus a parameter value (which comes from a user input, for example) cannot inject an SQL statement.
'
'---------------------------------------------------------------------------------------
Public Function ExecuteCommand(ByVal CmdText As String, _
                               ByVal CmdType As ADODB.CommandTypeEnum, _
                      Optional ByVal CmdParamDefs As Variant, _
                      Optional ByRef RecordsAffected As Long, _
                      Optional ByVal Options As Long = -1) As ADODB.Recordset
                      
   Dim cmd As ADODB.Command
   Dim ParamDefsArraySize As Long
   Dim i As Long
   Dim ParamName As String

   Set cmd = New ADODB.Command
   With cmd
   
      Set .ActiveConnection = Me.CurrentConnection
      .CommandText = CmdText
      .CommandType = CmdType
   
      If Not IsMissing(CmdParamDefs) Then
         ParamDefsArraySize = UBound(CmdParamDefs, 1)
         For i = 0 To ParamDefsArraySize
            ParamName = Nz(CmdParamDefs(i, 0), vbNullString)
            If Len(ParamName) > 0 Then
               .Parameters(ParamName).Value = CmdParamDefs(i, 1)
            End If
         Next i
      End If
      
      Set ExecuteCommand = .Execute(RecordsAffected, , Options)
      
   End With
   Set cmd = Nothing

End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordset
'---------------------------------------------------------------------------------------
'
' Optimizes ADODB.Recordset.Open method
'
' Parameters:
'     Source                  - Data source (e.g. SQL statement)
'     CursorType              - ADODB.CursorTypeEnum value
'     LockType                - ADODB.LockTypeEnum value
'     CursorLocation          - ADODB.CursorLocationEnum value
'     DisconnectedRecordset   - True: disconnect connection from recordset after opening
'
' Returns:
'     ADODB.Recordset
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordset(ByVal Source As String, _
                     Optional ByVal CursorType As ADODB.CursorTypeEnum = ADODB.CursorTypeEnum.adOpenForwardOnly, _
                     Optional ByVal LockType As ADODB.LockTypeEnum = ADODB.LockTypeEnum.adLockReadOnly, _
                     Optional ByVal CursorLocation As ADODB.CursorLocationEnum = -1, _
                     Optional ByVal DisconnectedRecordset As Boolean = False) As ADODB.Recordset

   Dim rst As ADODB.Recordset
   
   If CursorLocation = -1 Then
      CursorLocation = m_DefaultCursorLocation
   End If
   
   If DisconnectedRecordset Then
      If CursorLocation = adUseServer Then
         CursorLocation = adUseClient
      End If
      If CursorType = adOpenForwardOnly Then
         CursorType = adOpenKeyset
      End If
   End If

   Set rst = New ADODB.Recordset
   With rst
      .CursorLocation = CursorLocation
      .Open Source, Me.CurrentConnection, CursorType, LockType
      
      If DisconnectedRecordset Then
         Set .ActiveConnection = Nothing
      End If
      
   End With
   Set OpenRecordset = rst
   
End Function

'---------------------------------------------------------------------------------------
' Function: OpenRecordsetCommandParam
'---------------------------------------------------------------------------------------
'
' Open recordset using ADODB.Command
'
' Parameters:
'     CmdText                 - ADODB.CommandText value
'     CmdType                 - ADODB.CommandType value
'     CmdParamDefs            - 2-dimensional parameter array (Array(n,1) ... x(n,0) = parameter name, x(n,1) = parameter value)
'     RecordsAffected         - ByRef parameter to return the number of records affected by the Execute statement.
'     Options                 - additional options of the ADODB.Command.Execute method
'     DisconnectedRecordset   - True: disconnect connection from recordset after opening
'
' Returns:
'     ADODB.Recordset
'
' Remarks:
'     With this variant, SQL injection can be avoided very easily,
'     since the parameters are not built directly into the SQL statement
'     and thus a parameter value (which comes from a user input, for example) cannot inject an SQL statement.
'
'---------------------------------------------------------------------------------------
Public Function OpenRecordsetCommandParam(ByVal CmdText As String, _
                                          ByVal CmdType As ADODB.CommandTypeEnum, _
                                 Optional ByVal CmdParamDefs As Variant, _
                                 Optional ByRef RecordsAffected As Long, _
                                 Optional ByVal Options As Long = -1, _
                                 Optional ByVal DisconnectedRecordset As Boolean = False) As ADODB.Recordset

   Dim rst As ADODB.Recordset
   
   Set rst = ExecuteCommand(CmdText, CmdType, CmdParamDefs, RecordsAffected, Options)
   If DisconnectedRecordset Then
      Set rst.ActiveConnection = Nothing
   End If
   
   Set OpenRecordsetCommandParam = rst

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSQL
'---------------------------------------------------------------------------------------
'
' Lookup ADODB.Recordset replacement function for DLookup (passing a SQL statement)
'
' Parameters:
'     SqlText     - SQL statement
'     Index       - Field selection if the first field of the recordset is not to be returned.
'                   The index is passed in the same way as ADODB.Recordset.Fields(index).
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant  - Null, if SQL statement does not return a record.
'
' Remarks:
'     The SQL statement must be written in the SQL dialect of the DBMS concerned.
'
'---------------------------------------------------------------------------------------
Public Function LookupSql(ByVal SqlText As String, _
                 Optional ByVal Index As Variant = 0&, _
                 Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim rst As ADODB.Recordset

On Error GoTo HandleErr

   Set rst = Me.OpenRecordset(SqlText, adOpenForwardOnly, adLockReadOnly, adUseServer)
   With rst
      If .EOF Then
         LookupSql = ValueIfNull
      Else
         LookupSql = Nz(.Fields(Index), ValueIfNull)
      End If
      .Close
   End With
   Set rst = Nothing

ExitHere:
   Exit Function
   
HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "LookupSQL:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: Lookup
'---------------------------------------------------------------------------------------
'
' ADODB.Recordset replacement function for DLookup
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'     ValueIfNull - Return value if record is missing or data field value is Null (default: Null).
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Lookup(ByVal Expr As String, ByVal Domain As String, _
              Optional ByVal Criteria As Variant, _
              Optional ByVal ValueIfNull As Variant = Null) As Variant

   Dim SelectSql As String

   SelectSql = BuildSelectSql(Expr, Domain, Criteria, False)
   Lookup = Me.LookupSql(SelectSql, , ValueIfNull)

End Function

Private Function BuildSelectSql(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant, _
                       Optional ByVal Distinct As Boolean = False)

   Dim SelectSql As String

   SelectSql = "SELECT "
   If Distinct Then
      SelectSql = SelectSql & "Distinct "
   End If
   SelectSql = SelectSql & Expr & " FROM (" & Domain & ")"
   If Not (VarType(Criteria) = vbError) Then
      If Len(Criteria) > 0 Then
         SelectSql = SelectSql & " WHERE " & Criteria
      End If
   End If
   
   BuildSelectSql = SelectSql

End Function

'---------------------------------------------------------------------------------------
' Function: Count
'---------------------------------------------------------------------------------------
'
' ADODB.Recordset replacement function for DCount
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Long
'
'---------------------------------------------------------------------------------------
Public Function Count(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant, _
             Optional ByVal Distinct As Boolean = False) As Long
             
   If Distinct Then
      If Expr <> "*" Then
         Domain = "(" & BuildSelectSql(Expr, Domain, Criteria, True) & ")"
         Criteria = vbNullString
      End If
   End If
   
   Count = Nz(Me.Lookup("Count(" & Expr & ")", Domain, Criteria), 0)
   
End Function

'---------------------------------------------------------------------------------------
' Function: Max
'---------------------------------------------------------------------------------------
'
' ADODB.Recordset replacement function for DMax
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Max(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Max = Me.Lookup("Max(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Min
'---------------------------------------------------------------------------------------
'
' ADODB.Recordset replacement function for DMin
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Min(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Min = Me.Lookup("Min(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Sum
'---------------------------------------------------------------------------------------
'
' ADODB.Recordset replacement function for DSum
'
' Parameters:
'     Expr        - Field name or calculated expression (select part)
'     Domain      - Table, query or SQL expression for sub-query (FROM part)
'     Criteria    - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Variant
'
'---------------------------------------------------------------------------------------
Public Function Sum(ByVal Expr As String, ByVal Domain As String, Optional ByVal Criteria As Variant) As Variant
   Sum = Me.Lookup("Sum(" & Expr & ")", Domain, Criteria)
End Function

'---------------------------------------------------------------------------------------
' Function: Exists
'---------------------------------------------------------------------------------------
'
' Check if record exists
'
' Parameters:
'     Domain   - data source (FROM part)
'     Criteria - (Optional) Criteria string (WHERE part)
'
' Returns:
'     Boolean
'
'---------------------------------------------------------------------------------------
Public Function Exists(ByVal Domain As String, _
              Optional ByVal Criteria As String = vbNullString) As Boolean

   Dim strSQL As String

   strSQL = "SELECT 1 FROM " & Domain
   If Len(Criteria) > 0 Then
      strSQL = strSQL & " WHERE " & Criteria
   End If
   Exists = (LookupSql(strSQL, , 0) = 1)

End Function

'---------------------------------------------------------------------------------------
' Function: InsertIdentityReturn
'---------------------------------------------------------------------------------------
'
' Execute insert SQL statement and return last identity value (auto value)
'
' Parameters:
'     InsertSQL   - Insert SQL statement
'
' Returns:
'     Identity value or NULL if no record was appended
'
' Remarks:
'     Works for Jet only from Jet 4.0 (Access 2000), IdentityTable is only applicable for MSSQL, under MySQL the IDENT_CURRENT function does not exist.
'
'---------------------------------------------------------------------------------------
Public Function InsertIdentityReturn(ByVal InsertSQL As String, Optional ByVal IdentityTable As String = vbNullString) As Variant

   Dim cnn As ADODB.Connection
   Dim rst As ADODB.Recordset
   Dim RecordsAffected As Long
   Dim IdentitySql As String
   
   Dim CurrentHookResumeMode As Long
   Dim CurrentHookResumeMessage As Variant

   Set cnn = Me.CurrentConnection
   
   If m_HooksEnabled Then
      RaiseEvent HookInsertIdentityReturn(InsertSQL, IdentityTable, cnn, IdentitySql, CurrentHookResumeMode, CurrentHookResumeMessage)
      If CurrentHookResumeMode <> 0 Then 'Hook ermittelte bereits das Ergebnis
         InsertIdentityReturn = CurrentHookResumeMessage
         Exit Function
      End If
   End If
   
   cnn.Execute InsertSQL, RecordsAffected
   
   If RecordsAffected = 0 Then
      InsertIdentityReturn = Null 'es wurde kein neuer DS angefügt
      Exit Function
   End If

   If Len(IdentitySql) = 0 Then
      If Len(IdentityTable) > 0 Then
         IdentitySql = "SELECT IDENT_CURRENT('" & IdentityTable & "')" 'wegen trigger .. funktioniert nicht bei allen DBMS!!!
      Else
         IdentitySql = "SELECT @@Identity" 'zuletzt zugewiesener Autowert
      End If
   End If

On Error GoTo HandleErr

   Set rst = cnn.Execute(IdentitySql)
   With rst
      If .EOF Then
         InsertIdentityReturn = Null
      Else
         InsertIdentityReturn = .Fields(0)
      End If
      .Close
   End With
   Set rst = Nothing
   Set cnn = Nothing

ExitHere:
   Exit Function

HandleErr:
   If Not (rst Is Nothing) Then
      rst.Close
      Set rst = Nothing
   End If

   Err.Raise Err.Number, "InsertIdentityReturn:" & Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

'---------------------------------------------------------------------------------------
' Function: InsertValuesIdentityReturn
'---------------------------------------------------------------------------------------
'
' Create and execute insert SQL statement and return last identity value (auto value)
'
' Parameters:
'     Source         - Table order query name
'     InsertFields   - ParamArray: "Field1", Value1, "Field2", Value2, ....
'
' Returns:
'     Identity value or NULL if no record was appended
'
' Remarks:
'     Works for Jet only from Jet 4.0 (Access 2000), IdentityTable is only applicable for MSSQL, under MySQL the IDENT_CURRENT function does not exist.
'
'---------------------------------------------------------------------------------------
Public Function InsertValuesIdentityReturn(ByVal Source As String, ParamArray InsertFields() As Variant) As Variant

   Dim strInsert As String
   Dim strValues As String
   Dim i As Long

   For i = 0 To (UBound(InsertFields) - 1) Step 2
      strInsert = strInsert & ", " & InsertFields(i)
      strValues = strValues & ", " & InsertFields(i + 1)
   Next i
   strInsert = "INSERT INTO " & Source & " (" & Mid(strInsert, 3) & ") VALUES (" & Mid(strValues, 3) & ")"

   InsertValuesIdentityReturn = InsertIdentityReturn(strInsert, Source)

End Function

'---------------------------------------------------------------------------------------
' Function: ValueList
'---------------------------------------------------------------------------------------
'
' Generate concat string from recordset result
'
' Parameters:
'     SqlSource         - Select statement
'     ListConcatString  - (optional) Character to join results, Default = ", "
'     ValueIfNull       - (optional) Value if no results, Default = Null
'
' Returns:
'     Variant           - result exists: string, no result exists: ValueIfNull
'
'---------------------------------------------------------------------------------------
Public Function ValueList(ByVal SqlSource As String, Optional ByVal ListConcatString As String = ", ", Optional ValueIfNull As Variant = Null) As Variant

   Dim ListString As String

   With OpenRecordset(SqlSource, adOpenKeyset, adLockOptimistic, adUseClient, True)
      Do While Not .EOF
         ListString = ListString & ListConcatString & .Fields(0)
         .MoveNext
      Loop
      .Close
   End With

   If Len(ListString) > 0 Then
      ValueList = Mid(ListString, 3)
   Else
      ValueList = ValueIfNull
   End If

End Function

'---------------------------------------------------------------------------------------
' Function: LookupSqlValueCollection
'---------------------------------------------------------------------------------------
'
' Generate collection from recordset result
'
' Parameters:
'     SqlSource         - Select statement
'     ValueIfNull       - (optional) Value if no results, Default = Null
'
' Returns:
'     Collection
'
'---------------------------------------------------------------------------------------
Public Function LookupSqlValueCollection(ByVal SqlSource As String, Optional ByVal ValueIfNull As Variant = Null) As Collection

   Dim Results As Collection
   Dim rst As ADODB.Recordset
   Dim fld As ADODB.Field

   Set rst = OpenRecordset(SqlSource, adOpenKeyset, adLockOptimistic, adUseClient, True)

   Set Results = New Collection
   If Not rst.EOF Then
      For Each fld In rst.Fields
         Results.Add Nz(fld.Value, ValueIfNull), fld.Name
      Next
   End If
   rst.Close

   Set LookupSqlValueCollection = Results

End Function
