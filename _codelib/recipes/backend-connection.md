# Backend connection (DAO, ODBC, ADODB)

Keywords: backend, SQL Server, ODBC, OLEDB, ADODB, pass-through query, linked tables, relink, connection string, login,
Backend, Verbindung, Pass-Through-Abfrage, verknüpfte Tabellen neu verknüpfen, Verbindungszeichenfolge, Anmeldung

## Import

- `data/modDbConnection.bas` (dependencies are added by the import: DbConnectionManager, DbConnectionHandler,
  DbConnectionInfo, DbConnectionSettings, DaoHandler, AdodbHandler, OdbcHandler, SqlTools, ...)

With the ApplicationHandler framework additionally `base/ApplicationHandler_DbConnection.cls`
(see [application-framework.md](application-framework.md)).

## Prerequisite

The connection data is read from the table `usys_DbmsConnection` (DBMS, server, database, user, connection strings).
The table is created and filled with the ACLib DBMS Connection Wizard.

## Usage

```vba
' Check connection (login data via login form or event LoginDataRequired of DbConnectionInfo, if required)
If Not CheckConnectionStatus() Then
   Err.Raise vbObjectError, "StartApp", "No connection to the backend"
End If

' DAO on the frontend / linked tables
Debug.Print DbCon.DAO.Count("*", "tblCustomer")

' ODBC pass-through to the server (T-SQL)
Dim rst As DAO.Recordset
Set rst = DbCon.ODBC.OpenRecordsetPT("SELECT CustomerID, Name FROM dbo.Customer WHERE " & _
                                     SqlTools.TSql.BuildCriteria("Name", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, SearchText))
' ... use rst
rst.Close

Debug.Print DbCon.ODBC.LookupSqlPT("SELECT COUNT(*) FROM dbo.Customer")

' Relink ODBC tables and pass-through queries to the current connection
DbCon.ODBC.RelinkTablesAndQueries

' ADODB command with named parameters
Dim Params(0, 1) As Variant
Params(0, 0) = "FilterId": Params(0, 1) = CustomerId
Dim rsAdo As ADODB.Recordset
Set rsAdo = DbCon.ADODB.ExecuteCommand("PARAMETERS FilterId Long; SELECT * FROM tblCustomer WHERE CustomerID = [FilterId]", _
                                       adCmdText, Params)

' On application exit
DisposeDbConnection
```

## Notes

- `DbCon` creates the `DbConnectionManager` on first access; `CurrentConnectionInfo` gives server, database, DBMS
  and SQL formats of the current connection.
- Pass-through SQL is sent unchanged to the server: build values only with `SqlTools.TSql` (or the matching dialect),
  never by string concatenation of user input.
- `ExecuteCommand`/`OpenRecordsetCommandParam` set parameters by name (`Params(i, 0)` = name, `Params(i, 1)` = value);
  the names must be known to the provider.
- The event `BeforeInitDbConnection` of `DbConnectionManager` allows to adjust connection strings and settings.

## See also

- [dao-data-access.md](dao-data-access.md) - DaoHandler methods (also available via `DbCon.DAO`)
- [sql-criteria.md](sql-criteria.md) - SQL literals for T-SQL
- tests `_test/data/DbConnectionHandlerTests.cls`, `_test/data/odbc/OdbcHandlerTests.cls`, `_test/data/ado/AdodbHandlerTests.cls`
