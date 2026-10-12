# Combine several filter conditions

Keywords: filter string, WHERE clause, several conditions, AND/OR groups, IN list, sub-select, EXISTS, search criteria,
Filterstring, mehrere Bedingungen, Suchkriterien, Bedingungsgruppen, Unterabfrage

## Import

- `data/FilterStringBuilder.cls` (uses `data/SqlTools.cls` and `text/StringCollection.cls`)

## Usage

```vba
Dim FilterString As String

With New FilterStringBuilder
   Set .SqlTool = SqlTools.DAO            ' SQL dialect (DAO/Jet/ACE); SqlTools.TSql for SQL Server

   .Add "StatusID", SQL_Numeric, SQL_Equal, Me.cboStatus.Value                        ' Null => ignored
   .Add "OrderDate", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, Me.txtFrom.Value, Me.txtTo.Value

   With .NewConditionGroup(SQL_Or)                                                    ' (... Or ...)
      .Add "LastName", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.txtSearch.Value
      .Add "FirstName", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.txtSearch.Value
   End With

   FilterString = .ToString               ' conditions combined with And (default)
End With
```

Further options:

```vba
.Add "Country", SQL_Text, SQL_In, Array("AT", "DE", "CH")                             ' Country In ('AT','DE','CH')
.Add "Tag", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix + SQL_SplitValueToArray, "a;b;c"  ' value list from a string
.AddCriteria "Deleted = False"                                                        ' own criteria string

' CustomerID In (Select CustomerID From tblOrder Where ...); True = ignore if the sub-select has no criteria
With .AddSubSelectCriteria("CustomerID", SQL_In, "Select CustomerID From tblOrder", True)
   .Add "Amount", SQL_Numeric, SQL_Equal + SQL_GreaterThan, 1000                      ' WHERE of the sub-select
End With

With .AddExistsCriteria("Select * From tblOrder AS T2")
   .AddCriteria "T2.CustomerID = tblCustomer.CustomerID"
End With

FilterString = .ToString(SQL_Or)          ' or combine with Or
```

## Notes

- Empty values (`Null`) do not create a condition, so the builder can be filled directly from unbound search controls.
- Set `.SqlTool` explicitly. Without it a clone of the global `SqlTools` default instance is used, which has no SQL
  dialect configured.
- Parameters of `.Add` are the same as `SqlTools.BuildCriteria` (see [sql-criteria.md](sql-criteria.md)).
- Field names and the SQL text of sub-selects are inserted as they are: never build them from user input.

## See also

- [filter-form.md](filter-form.md) - the same with form controls and automatic filtering
- `_examples/data/FilterStringBuilder_Examples.bas`, tests `_test/data/FilterStringBuilderTests.cls`
