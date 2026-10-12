# DAO data access: parameter queries, lookups, transactions

Keywords: DAO, execute SQL, parameter query, QueryDef parameters, insert, update, delete, recordset, DLookup replacement,
DCount, transaction, rollback, Parameterabfrage, Datensatz anfügen, Transaktion, Nachschlagen

## Import

- `data/dao/DaoHandler.cls`
- `data/dao/DaoTransactionHandler.cls` (only for transactions)

## Usage

`DaoHandler` is a predeclared class: use it directly (`DaoHandler.Xxx`, works on `CurrentDb`) or create an instance
for another database (`Set Hdl = New DaoHandler: Set Hdl.CurrentDb = OpenDatabase(...)`).

```vba
' Execute with parameters (values are not concatenated into the SQL text)
Dim RecordsAffected As Long
RecordsAffected = DaoHandler.ExecuteParamSql( _
      "PARAMETERS [pName] Text(255), [pAmount] Long; " & _
      "INSERT INTO tblOrder (CustomerName, Amount) VALUES ([pName], [pAmount])", _
      Me.txtName.Value, Me.txtAmount.Value)

' Saved query with named parameters
DaoHandler.ExecuteQueryDefByName "qryInsertOrder", DaoHandler.GetNamedParamDefArray("pName", "abc", "pAmount", 123)

' Recordset with parameters
Dim rst As DAO.Recordset
Set rst = DaoHandler.OpenRecordsetParamSql( _
      "PARAMETERS [pId] Long; SELECT * FROM tblOrder WHERE OrderID = [pId]", _
      dbOpenDynaset, dbSeeChanges, dbOptimistic, OrderId)
' ... use rst
rst.Close
Set rst = Nothing

' Lookups (replacement for DLookup/DCount, with value for Null)
Dim Amount As Variant
Amount = DaoHandler.LookupSql("SELECT Amount FROM tblOrder WHERE OrderID = 5", , 0)
Amount = DaoHandler.Lookup("Amount", "tblOrder", SqlTools.DAO.BuildCriteria("OrderID", SQL_Numeric, SQL_Equal, OrderId), 0)
Debug.Print DaoHandler.Count("*", "tblOrder"), DaoHandler.Exists("tblOrder", "Amount > 1000")
```

Transaction (all statements or none):

```vba
Public Sub BookOrder(ByVal OrderId As Long)

   Dim Trans As DaoTransactionHandler
   Dim ErrNumber As Long, ErrSource As String, ErrDescription As String

   Set Trans = New DaoTransactionHandler

On Error GoTo HandleErr

   Trans.BeginTrans
   DaoHandler.ExecuteParamSql "PARAMETERS [pId] Long; UPDATE tblOrder SET Booked = True WHERE OrderID = [pId]", OrderId
   DaoHandler.ExecuteParamSql "PARAMETERS [pId] Long; INSERT INTO tblLog (OrderID) VALUES ([pId])", OrderId
   Trans.CommitTrans
   Exit Sub

HandleErr:
   ErrNumber = Err.Number: ErrSource = Err.Source: ErrDescription = Err.Description
   If Trans.TransIsStarted Then Trans.Rollback
   Err.Raise ErrNumber, ErrSource, ErrDescription

End Sub
```

## Notes

- `ExecuteParamSql`, `ExecuteQueryDefByName` and `ExecuteQueryDef` execute with `dbFailOnError`.
  `DaoHandler.Execute` only adds `dbSeeChanges`: pass `dbFailOnError` yourself (`DaoHandler.Execute Sql, dbFailOnError`).
- Parameters are passed in the order of the `PARAMETERS` declaration, or by name with `GetNamedParamDefArray`.
- Criteria strings for `Lookup`, `Count`, `Exists` etc. should be built with SqlTools (see [sql-criteria.md](sql-criteria.md)).
- `DaoTransactionHandler` works on `DBEngine.Workspaces(0)` (property `WorkspaceID`); an open transaction is rolled back
  when the object is terminated.
- `InsertIdentityReturn` (new AutoNumber), `InsertValues`, `CopyRecord(s)` and `NewRecord` are available for inserts.

## See also

- [backend-connection.md](backend-connection.md) - DAO/ODBC/ADODB access to a backend
- `_examples/data/dao/DaoHandler_Beispiele.bas`, tests `_test/data/dao/DaoHandlerTests.cls`,
  `_test/data/dao/DaoTransactionHandlerTests.cls`
