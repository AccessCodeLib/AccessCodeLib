# Build SQL criteria and SQL literals safely

Keywords: dynamic SQL, WHERE, criteria, filter string, SQL injection, date literal, text literal, quotes,
Kriterium, Filter, dynamisches SQL, Datum im SQL, Hochkomma, SQL-Injection

## Import

- `data/SqlTools.cls`

## Usage

```vba
' Criteria for a form filter or a WHERE condition (DAO/Jet/ACE syntax)
Dim Criteria As String
Criteria = SqlTools.DAO.BuildCriteria("LastName", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Me.txtLastName.Value)
' "abc" => LastName Like 'abc*'   |   Null => "" (no condition)

With Me
   .Filter = Criteria
   .FilterOn = (Len(Criteria) > 0)
End With
```

More variants (`With SqlTools.DAO`):

```vba
.BuildCriteria("F", SQL_Text, SQL_Equal, "abc")                              ' F = 'abc'
.BuildCriteria("F", SQL_Numeric, SQL_Equal + SQL_LessThan, 123)              ' F <= 123
.BuildCriteria("F", SQL_Numeric, SQL_Between, 123, 456)                      ' F Between 123 And 456
.BuildCriteria("F", SQL_Numeric, SQL_In, Array(123, 456, 789))               ' F In (123,456,789)
.BuildCriteria("F", SQL_Text, SQL_Like + SQL_Add_WildCardSuffix, Array("a", "c"))  ' F Like 'a*' Or F Like 'c*'
.BuildCriteria("F", SQL_Date, SQL_Between + SQL_Add_WildCardSuffix, #1/1/2013#, #12/31/2013#)
                                                     ' whole last day: F >= #2013-01-01# And F < #2014-01-01#
```

Single values as SQL literals:

```vba
SqlTools.DAO.TextToSqlText("O'Neil")     ' 'O''Neil'
SqlTools.DAO.DateToSqlText(Now)          ' #2023-12-24 13:45:00#  (locale independent)
SqlTools.DAO.NumberToSqlText(1.5)        ' 1.5  (decimal point, also with German locale)
SqlTools.DAO.ConvertToSqlText(Value, SQL_Boolean)
```

## Notes

- Use a configured instance: `SqlTools.DAO` (Jet/ACE, `#yyyy-mm-dd hh:nn:ss#`, `True`, `*`) or `SqlTools.TSql`
  (SQL Server, `'yyyymmdd hh:nn:ss'`, `1`, `%`), or `SqlTools.NewInstance(DateFormat, BooleanTrueString, WildCard)`.
  The default instance `SqlTools` without configuration raises an error for date and boolean values.
- `Null` as filter value creates no condition (empty string). With `SQL_Between` and one `Null` value the result is
  `>=` or `<=`. Use the parameter `IgnoreValue` for further values that should not filter (e.g. 0 or "*").
- `SQL_Like` without a wildcard in the value creates `=`.
- `SQL_Add_WildCardSuffix` with dates and `<=`/`Between` includes the whole day of the end date.
- Field names are inserted as they are: never pass user input as field name.
- For executing insert/update/delete statements prefer parameter queries (see [dao-data-access.md](dao-data-access.md)).
  SqlTools is meant for dynamic criteria (form/report filter, `WhereCondition` of `DoCmd.OpenForm`/`OpenReport`,
  domain functions, pass-through SQL).

## See also

- [filter-string.md](filter-string.md) - several conditions and condition groups
- [filter-form.md](filter-form.md) - search form with filter controls
- `_examples/data/SqlTools_BuildCriteria_Examples.bas`, tests `_test/data/SqlToolsTests.cls`,
  `_test/data/SqlToolsBuildCriteriaTests.cls`
