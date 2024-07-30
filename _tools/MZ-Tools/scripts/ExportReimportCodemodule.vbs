' Call: ExportReimportCodemodule.vbs  {accdb full path} {codemodule name}
' Mz-Tools parameter for external utilitites: $P[PROJECT_FILENAME_WITH_PATH] $P[PROJECT_ITEM_NAME]
'
appFile = WScript.Arguments(0)
codeModuleName = WScript.Arguments(1)

set app = getObject(appFile)

tempPath = WScript.CreateObject("Scripting.FileSystemObject").GetSpecialFolder(2)
vbcExportFileName = tempPath & "\" & codeModuleName & ".mztexportreimport.cls"

set vbp = app.VBE.ActiveVBProject
set vbc = vbp.VBComponents(codeModuleName)
vbc.Export vbcExportFileName
vbp.VBComponents.Remove vbc

vbp.VBComponents.Import vbcExportFileName
vbp.VBComponents(codeModuleName).Activate

Set fso = CreateObject("Scripting.FileSystemObject")
fso.DeleteFile vbcExportFileName
