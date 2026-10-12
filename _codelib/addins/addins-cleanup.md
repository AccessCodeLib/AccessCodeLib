# Clean up `_codelib/addins` (postponed)

Status: analysis of 2026-10-12, nothing changed yet. Suggested branch: `feature/addins-cleanup` from `master`.

## Findings

`_codelib/addins` contains source modules of the add-ins and modules shared by them.

1. **Outdated add-in copies:** `ImportWizard`, `FilterFormWizard`, `AccUnitLoader`, `StructViewer` were last changed
   in 2022. The add-ins are developed in their own repositories, e.g. `ACLibFileManager.cls` has 2028 lines in
   ACLibImportWizard and 1695 lines here. These copies also contain outdated `<use>` targets:
   `file/modFiles.bas` (now `file/FileTools.bas`), `data/dao/modDAO_Tools.bas` (now `data/dao/DaoTools.bas`?),
   `data/modSQL_Tools.bas` (now `data/SqlTools.cls`?), `defGlobal_ACLibFilterFormWizard.bas`,
   `shared/AccUnitConfiguration.cls`, and a malformed line in `frmFilterFormWizard.frm`
   (`<use>text/StringTools.bas</code>`).
2. **`shared/` is in use, but outdated here:** the add-in repositories import these modules from this repository
   (`<file>_codelib/addins/shared/...</file>`), but their own copies are newer. Removing `shared/` would break add-in
   updates (`missing`).

| Module | Newer version in | Differing lines |
|---|---|---|
| `ACLibConfiguration.cls` | ACLibImportWizard | 76 |
| `AddInConfiguration.cls` | ACLibImportWizard | 2 |
| `AddInInstaller.cls` | ACLibImportWizard | 4 |
| `CodeModuleHeaderItem.cls` | ACLibImportWizard | 12 |
| `CodeModuleProcedure.cls` | ACLibImportWizard | 6 |
| `CodeModuleReader.cls` | ACLibImportWizard | 120 |
| `AppFileCodeModulTransfer.cls` | FilterFormWizard | 46 |
| `CodeModuleWriter.cls` | FilterFormWizard | 32 |
| `FormDesigner.cls` | FilterFormWizard | 56 |
| `VbeTools.cls` | FilterFormWizard | 8 |
| `modWizardCodeModulesData.bas` | FilterFormWizard | 10 |
| `ACLibGitHubImporter.cls` | **both, diverged** (ImportWizard 76, FilterFormWizard 110, between them 54) | merge required |

`InstallAddInForm.frm` is used by ACLibImportWizard, but stored there in another export format
(`source/forms/InstallAddInForm.cls` + `.form`): compare manually.

Not used by the two add-in repositories: `ACLibWebImporter.cls`, `CodeModuleStructReader.cls`, `Test/`.

## Plan

1. Remove `ImportWizard`, `FilterFormWizard`, `AccUnitLoader`, `StructViewer` and replace `_codelib/addins/_info.txt`
   with a `README.md` that links to the add-in repositories:
   - ImportWizard: https://github.com/AccessCodeLib/ACLibImportWizard
   - FilterFormWizard: https://github.com/AccessCodeLib/FilterFormWizard
   - AccUnitLoader, StructViewer: repositories still to be clarified (StructViewer -> ACLibCodeStructReader?)
2. Sync `shared/`: take the newer version of the 11 modules with one source after reviewing the diff
   (keep encoding and CRLF); merge `ACLibGitHubImporter.cls` from both add-ins.
3. Leave the unused `shared/` modules unchanged (decide later).
4. Commits: "Replace outdated add-in sources with links to add-in repositories",
   "Sync shared add-in modules from add-in repositories".

## Notes

- Compare with `diff --strip-trailing-cr` (line endings differ between repositories).
- After the sync the add-in repositories get the merged state with their next update via the Import Wizard.
