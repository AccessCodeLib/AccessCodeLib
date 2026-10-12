# AccessCodeLib recipes

Task-oriented guides: which modules to import for a task, minimal usage code and pitfalls.
The import paths are the repository paths for the ACLib Import Wizard (`aclib_resolve` / `aclib_import`);
dependencies are added by the import. Module overview: [../catalog/catalog.md](../catalog/catalog.md).

| Task | Recipe | Main module |
|---|---|---|
| Build WHERE criteria and SQL literals safely (DAO, T-SQL) | [sql-criteria.md](sql-criteria.md) | `data/SqlTools.cls` |
| Combine several filter conditions, groups, sub-selects | [filter-string.md](filter-string.md) | `data/FilterStringBuilder.cls` |
| Search form with filter controls | [filter-form.md](filter-form.md) | `form/filter/FilterControlManager.cls` |
| DAO: parameter queries, lookups, transactions | [dao-data-access.md](dao-data-access.md) | `data/dao/DaoHandler.cls` |
| Backend connection (DAO, ODBC pass-through, ADODB) | [backend-connection.md](backend-connection.md) | `data/modDbConnection.bas` |
| File/folder dialogs, paths, temp files, zip | [file-dialogs-and-paths.md](file-dialogs-and-paths.md) | `file/FileTools.bas` |

Recipe structure: `Keywords` (also German), `Import`, `Usage`, `Notes`, `See also`.
