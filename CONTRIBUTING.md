# Contributing

Contributions are welcome: new modules, bug fixes, bug reports, recipes, tests and documentation.

## Bug reports and ideas

Open an [issue](https://github.com/AccessCodeLib/AccessCodeLib/issues). For bugs, name the module
(repository path, e.g. `data/SqlTools.cls`), the Access version (32/64-bit) and a short code sample.

## Branches

| Branch | Purpose |
|---|---|
| `master` | Released state, published to the code reference and the import wizard |
| `draft` | Integration branch - target of all pull requests |

1. Fork the repository (or create a branch from `draft`).
2. Open a pull request into `draft`. The `catalog-check` workflow must be green.
3. `draft` is merged into `master` with **Create a merge commit**; afterwards a workflow
   fast-forwards `draft` to `master`.

## Modules

* Export the code from Access as text files (`.bas`, `.cls`, `.frm`) - binary `.accdb`/`.mdb` files are not
  part of the repository.
* Keep the encoding of existing files (mostly Windows-1252) and CRLF line endings.
* Every module needs `Option Compare Text` and `Option Explicit`; declare variables with explicit types.
* API declarations use `PtrSafe` and `LongPtr` for handles and pointers.
* Code, identifiers, comments and commit messages in English; commit subject in imperative mood,
  max. ~72 characters.

### `<codelib>` block

Every module starts with a header comment followed by a `<codelib>` block. The ACLib Import Wizard reads
this block to import a module with its dependencies, so it must stay valid:

| Tag | Content |
|---|---|
| `<file>` | Repository path of the module itself |
| `<license>` | License module, usually `_codelib/license.bas` |
| `<replace>` | Repository path of a module that this module replaces (e.g. an older variant or `base/_config_Application.bas`) |
| `<use>` | Repository path of a required module (one tag per dependency) |
| `<ref>` | Required library reference |
| `<test>` | AccUnit test class in `_test/` |
| `<example>` | Usage example in `_examples/` |
| `<description>` | One-line description, shown in the catalog |
| `<execute>` | Procedure call run after the import, e.g. a setup routine (`OptionManagerSetup_SetupTable()`) |

See `data/SqlTools.cls` for an example.

### Documentation

Document modules and public members in [Natural Docs](https://www.naturaldocs.org/) style, see
`data/SqlTools.cls`. The [code reference](https://accesscodelib.github.io/AccessCodeLib/) is built
from these comments on every push to `master`.

### Catalog

After adding or changing modules, regenerate the catalog and commit the result:

```powershell
.\_tools\catalog\Build-Catalog.ps1
```

The `catalog-check` workflow fails if `_codelib/catalog` is out of date.

## Recipes

Recipes in `_codelib/recipes/` are task-oriented guides. Use the structure `Keywords`, `Import`, `Usage`,
`Notes`, `See also` and add the recipe to the table in [_codelib/recipes/README.md](_codelib/recipes/README.md).

## Tests

Add [AccUnit](https://github.com/AccessCodeLib/AccUnit) test classes to `_test/` (same folder structure as
the module) and reference them with a `<test>` tag in the module's `<codelib>` block.
