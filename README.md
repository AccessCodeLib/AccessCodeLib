# Access Code Library

Reusable VBA modules and classes for _Microsoft Access_ applications.

Every module starts with a `<codelib>` header block that describes the file, its dependencies,
license, tests and examples. Tools such as the ACLib Import Wizard read this block to import a
module together with everything it needs.

## Importing modules

Use the [ACLib Import Wizard](https://github.com/AccessCodeLib/ACLibImportWizard) to import,
export and update library code in an Access database. Select a module by its repository path
(e.g. `data/SqlTools.cls`); dependencies are imported automatically.

Example of a `<codelib>` block (`data/SqlTools.cls`):

```vb
'<codelib>
'  <file>data/SqlTools.cls</file>
'  <license>_codelib/license.bas</license>
'  <test>_test/data/SqlToolsTests.cls</test>
'  <description>Converts values to SQL literals (text, number, date, boolean) for DAO or T-SQL and builds criteria safely</description>
'  <example>_examples/data/SqlTools_BuildCriteria_Examples.bas</example>
'</codelib>
```

## Finding the right module

* [Recipes](_codelib/recipes/README.md) - task-oriented guides: which modules to import, minimal usage code, pitfalls
* [Module catalog](_codelib/catalog/catalog.md) - all modules with description and dependencies
  (generated; details in [catalog.json](_codelib/catalog/catalog.json))
* [Code reference](https://accesscodelib.github.io/AccessCodeLib/) - API documentation (Natural Docs)

## Repository layout

| Folder | Content |
|---|---|
| `api` | Wrappers for Windows API, Office automation, Graphviz, PDFCreator |
| `base` | ApplicationHandler framework (application start, settings, extensions) |
| `com`, `vba` | COM helpers, VBA error trapping |
| `crypt` | MD5 |
| `data` | SQL tools, filter strings, DAO/ADODB/ODBC handlers, backend connections |
| `datetime`, `math`, `text` | Data types and functions for dates, numbers and strings |
| `file` | File and folder tools, file search |
| `form` | Form extensions: filter controls, sorting, undo, audit trail, info messages |
| `localization` | Translation dictionary |
| `ribbon` | Ribbon callbacks and control state |
| `usability`, `utils` | UI helpers: option manager, calculation text box, progress and status bar |
| `user` | User handling, roles, login form |
| `misc` | Stopwatch |
| `_codelib` | Catalog, recipes, license module, shared add-in code, configuration samples |
| `_examples` | Usage examples (referenced via `<example>` tags) |
| `_test` | AccUnit test classes (referenced via `<test>` tags) |
| `_dotnetlib` | .NET libraries (SqlTools) and their VBA integration |
| `_packages` | Packages, e.g. the Access add-in template |
| `_tools` | Catalog generator, MZ-Tools settings |

## Using with AI agents

* [ACLibImportWizard-MCP](https://github.com/AccessCodeLib/ACLibImportWizard-MCP) - MCP server that lets
  an AI agent resolve and import modules (`aclib_resolve`, `aclib_import`) through the ACLib Import Wizard
* [aclib-import skill](https://github.com/AccessCodeLib/ACLibImportWizard-MCP/tree/main/skills/aclib-import) -
  agent workflow: find modules in the catalog, resolve dependencies, import after approval, compile
* [catalog.json](_codelib/catalog/catalog.json) - machine-readable module list (dependencies, used-by,
  references, tests, public members)
* [Recipes](_codelib/recipes/README.md) - import paths and usage code per task

## Tests

Test classes for [AccUnit](https://github.com/AccessCodeLib/AccUnit) are in `_test/`. Each tested module
links its test classes with `<test>` tags in its `<codelib>` block.

## Contributing

Contributions are welcome - new modules, bug reports, recipes or documentation.
See [CONTRIBUTING.md](CONTRIBUTING.md).

## License

BSD 3-Clause License, see [LICENSE](LICENSE).
