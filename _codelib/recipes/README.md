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
| Open a file or URL, send mail, start a program | [open-file-or-url.md](open-file-or-url.md) | `api/winapi/WinApiShellTools.cls` |
| Text functions (format, pad, trim, placeholders) | [text-functions.md](text-functions.md) | `text/StringTools.cls` |
| Non-modal info message with auto close | [info-message.md](info-message.md) | `form/InfoMessageBox.cls` |
| Sort a continuous form by clicking on the column header | [sort-continuous-form.md](sort-continuous-form.md) | `form/sort/FormSortingExtension.cls` |
| Export data to Excel, report to PDF | [export-to-excel-or-pdf.md](export-to-excel-or-pdf.md) | `api/office/excel/modExcel_Tools.bas` |
| Application start with the ApplicationHandler framework | [application-framework.md](application-framework.md) | `base/_initApplication.bas` |

Recipe structure: `Keywords` (also German), `Import`, `Usage`, `Notes`, `See also`.
