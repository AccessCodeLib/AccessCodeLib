<#
.SYNOPSIS
    Builds the AccessCodeLib module catalog from the <codelib> header blocks.

.DESCRIPTION
    Scans all code files (*.cls, *.bas, *.frm, *.rpt, *.package) and writes:
      catalog.md   - one line per module, grouped by folder (overview for humans and AI agents)
      catalog.json - details: dependencies (direct/transitive), used-by, references, tests,
                     same object name, examples, executes, public members
      report.md    - modules without <description> tag and header issues

    Files in _test, _examples, _codelib, _dotnetlib and _tools are not listed in the catalog,
    but they are used to resolve dependencies.

.PARAMETER RepositoryRoot
    Root folder of the AccessCodeLib repository. Default: two levels above this script.

.PARAMETER OutputDirectory
    Target folder for the generated files. Default: <RepositoryRoot>\_codelib\catalog

.PARAMETER Check
    Do not write files. Exit code 1 if the generated files are missing or out of date.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File _tools\catalog\Build-Catalog.ps1

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File _tools\catalog\Build-Catalog.ps1 -Check
#>
[CmdletBinding()]
param(
    [string]$RepositoryRoot,
    [string]$OutputDirectory,
    [switch]$Check
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

if (-not $RepositoryRoot) {
    $RepositoryRoot = (Resolve-Path (Join-Path $PSScriptRoot '..\..')).Path
}
$RepositoryRoot = (Resolve-Path $RepositoryRoot).Path.TrimEnd('\')
if (-not $OutputDirectory) {
    $OutputDirectory = Join-Path $RepositoryRoot '_codelib\catalog'
}

$ExcludedTopFolders = @('_test', '_examples', '_codelib', '_dotnetlib', '_tools')
$CodeFileExtensions = @('.cls', '.bas', '.frm', '.rpt', '.package')
$NewLine = "`r`n"

#region File reading

function Read-SourceFile {
    param([string]$Path)

    $Bytes = [System.IO.File]::ReadAllBytes($Path)
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) {
        return [System.Text.Encoding]::UTF8.GetString($Bytes, 3, $Bytes.Length - 3)
    }
    if ($Bytes.Length -ge 2 -and $Bytes[0] -eq 0xFF -and $Bytes[1] -eq 0xFE) {
        return [System.Text.Encoding]::Unicode.GetString($Bytes, 2, $Bytes.Length - 2)
    }
    try {
        $StrictUtf8 = New-Object System.Text.UTF8Encoding($false, $true)
        return $StrictUtf8.GetString($Bytes)
    }
    catch {
        return [System.Text.Encoding]::GetEncoding(1252).GetString($Bytes)
    }
}

#endregion

#region Header parsing

function Get-CommentText {
    param([string]$Line)
    return ($Line -replace "^\s*'", '').Trim()
}

function Get-TagValues {
    param([string[]]$BlockLines, [string]$Tag)

    $Values = New-Object System.Collections.Generic.List[string]
    foreach ($Line in $BlockLines) {
        foreach ($Match in [regex]::Matches($Line, "<$Tag>(.*?)</$Tag>")) {
            $Value = $Match.Groups[1].Value.Trim()
            if ($Value) { $Values.Add($Value) }
        }
    }
    return , $Values.ToArray()
}

function Get-MalformedTagLines {
    param([string[]]$BlockLines)

    $Result = New-Object System.Collections.Generic.List[string]
    foreach ($Line in $BlockLines) {
        foreach ($Match in [regex]::Matches($Line, '<([a-z]+)>')) {
            $Tag = $Match.Groups[1].Value
            if ($Tag -eq 'codelib') { continue }
            if ($Line -notmatch "</$Tag>") {
                $Result.Add($Line.Trim())
                break
            }
        }
    }
    return , $Result.ToArray()
}

function Get-References {
    param([string[]]$BlockLines)

    $Result = New-Object System.Collections.Generic.List[object]
    foreach ($Value in (Get-TagValues $BlockLines 'ref')) {
        $Ref = [ordered]@{ name = ''; major = $null; minor = $null; guid = '' }
        if ($Value -match '<name>(.*?)</name>') { $Ref.name = $Matches[1].Trim() }
        if ($Value -match '<major>(\d+)</major>') { $Ref.major = [int]$Matches[1] }
        if ($Value -match '<minor>(\d+)</minor>') { $Ref.minor = [int]$Matches[1] }
        if ($Value -match '<guid>(.*?)</guid>') { $Ref.guid = $Matches[1].Trim() }
        $Result.Add($Ref)
    }
    return , $Result.ToArray()
}

function Get-HeaderSummary {
    param([string[]]$HeaderLines)

    # Doxygen style: <summary> ... </summary>
    $HeaderText = ($HeaderLines | ForEach-Object { Get-CommentText $_ }) -join ' '
    if ($HeaderText -match '<summary>\s*(.*?)\s*</summary>') {
        $Summary = ($Matches[1] -replace '\s+', ' ').Trim()
        if ($Summary) { return $Summary }
    }

    # NaturalDocs style: "' Class: name", separator, blank, description lines
    $TitlePattern = "^\s*'\s*(Class|Package|Module|Form|Report|Interface|Klasse|Modul|Formular)\s*:"
    for ($i = 0; $i -lt $HeaderLines.Count; $i++) {
        if ($HeaderLines[$i] -notmatch $TitlePattern) { continue }
        $Parts = New-Object System.Collections.Generic.List[string]
        for ($j = $i + 1; $j -lt $HeaderLines.Count; $j++) {
            $Text = Get-CommentText $HeaderLines[$j]
            $IsSeparator = $Text -match '^-{5,}$'
            $IsEmpty = $Text -eq ''
            if ($Parts.Count -eq 0) {
                if ($IsSeparator -or $IsEmpty) { continue }
            }
            if ($IsSeparator -or $IsEmpty -or $Text -match '^[A-Za-z][\w ]*:$' -or $Text.StartsWith('<')) { break }
            $Parts.Add($Text)
        }
        if ($Parts.Count -gt 0) { return (($Parts -join ' ') -replace '\s+', ' ').Trim() }
        break
    }
    return ''
}

#endregion

#region Code parsing

function Remove-TrailingComment {
    param([string]$Line)

    $InString = $false
    for ($i = 0; $i -lt $Line.Length; $i++) {
        $Char = $Line[$i]
        if ($Char -eq '"') { $InString = -not $InString }
        elseif ($Char -eq "'" -and -not $InString) { return $Line.Substring(0, $i) }
    }
    return $Line
}

function Get-PublicMembers {
    param([string[]]$CodeLines)

    $MemberPattern = '^\s*(?:Public\s+)?(?:Static\s+)?(Sub|Function|Property\s+Get|Property\s+Let|Property\s+Set|Event|Enum|Type)\s+(\w+)'
    $Result = New-Object System.Collections.Generic.List[string]
    $Seen = @{}

    for ($i = 0; $i -lt $CodeLines.Count; $i++) {
        $Line = $CodeLines[$i]
        if ($Line -notmatch $MemberPattern) { continue }
        $Kind = ($Matches[1] -replace '\s+', ' ')
        $Name = $Matches[2]

        # join line continuations
        $Statement = Remove-TrailingComment $Line
        while ($Statement -match '\s_\s*$' -and $i + 1 -lt $CodeLines.Count) {
            $i++
            $Statement = ($Statement -replace '\s_\s*$', ' ') + (Remove-TrailingComment $CodeLines[$i])
        }

        $Key = "$Kind $Name".ToLowerInvariant()
        if ($Seen.ContainsKey($Key)) { continue }   # e.g. #If EARLYBINDING variants
        $Seen[$Key] = $true

        if ($Kind -eq 'Enum' -or $Kind -eq 'Type') {
            $Signature = "$Kind $Name"
        }
        else {
            $Signature = ($Statement -replace '^\s*Public\s+', '' -replace '\s+', ' ').Trim()
        }
        $Result.Add($Signature)
    }
    return , $Result.ToArray()
}

#endregion

#region Module model

function Get-ModuleKind {
    param([string]$Extension)
    switch ($Extension.ToLowerInvariant()) {
        '.cls' { 'class' }
        '.bas' { 'module' }
        '.frm' { 'form' }
        '.rpt' { 'report' }
        '.package' { 'package' }
    }
}

function Read-ModuleInfo {
    param([System.IO.FileInfo]$File)

    $RelativePath = $File.FullName.Substring($RepositoryRoot.Length + 1).Replace('\', '/')
    $Lines = (Read-SourceFile $File.FullName) -split "\r?\n"
    $Kind = Get-ModuleKind $File.Extension

    # forms/reports: code starts after "CodeBehindForm"
    $CodeStart = 0
    if ($Kind -eq 'form' -or $Kind -eq 'report') {
        $CodeStart = $Lines.Count
        for ($i = 0; $i -lt $Lines.Count; $i++) {
            if ($Lines[$i] -match '^\s*CodeBehind(Form|Report)\s*$') { $CodeStart = $i + 1; break }
        }
    }

    $BlockStart = -1
    $BlockEnd = -1
    for ($i = $CodeStart; $i -lt $Lines.Count; $i++) {
        if ($BlockStart -lt 0 -and $Lines[$i] -match "^\s*'\s*<codelib>") { $BlockStart = $i; continue }
        if ($BlockStart -ge 0 -and $Lines[$i] -match "^\s*'\s*</codelib>") { $BlockEnd = $i; break }
    }

    $Name = [System.IO.Path]::GetFileNameWithoutExtension($File.Name)
    $Predeclared = $false
    foreach ($Line in $Lines) {
        if ($Line -match '^Attribute VB_Name = "(.+)"') { $Name = $Matches[1] }
        elseif ($Line -match '^Attribute VB_PredeclaredId = True') { $Predeclared = $true }
    }

    $Info = [ordered]@{
        path              = $RelativePath
        name              = $Name
        kind              = $Kind
        hasCodeLibBlock   = ($BlockStart -ge 0 -and $BlockEnd -gt $BlockStart)
        fileTag           = ''
        packageName       = ''
        description       = ''
        descriptionSource = 'none'
        headerSummary     = ''
        predeclaredId     = $Predeclared
        license           = ''
        uses              = @()
        replaces          = @()
        references        = @()
        tests             = @()
        examples          = @()
        executes          = @()
        publicMembers     = @()
        malformedTags     = @()
    }

    if ($Info.hasCodeLibBlock) {
        $BlockLines = $Lines[($BlockStart + 1)..($BlockEnd - 1)]
        $HeaderLines = if ($BlockStart -gt $CodeStart) { $Lines[$CodeStart..($BlockStart - 1)] } else { @() }
        $CodeLines = if ($BlockEnd + 1 -lt $Lines.Count) { $Lines[($BlockEnd + 1)..($Lines.Count - 1)] } else { @() }

        $FileTags = Get-TagValues $BlockLines 'file'
        if ($FileTags.Count -gt 0) { $Info.fileTag = $FileTags[0] }
        $PackageTags = Get-TagValues $BlockLines 'package'
        if ($PackageTags.Count -gt 0) { $Info.packageName = $PackageTags[0] }
        $Descriptions = Get-TagValues $BlockLines 'description'
        $LicenseTags = Get-TagValues $BlockLines 'license'
        if ($LicenseTags.Count -gt 0) { $Info.license = $LicenseTags[0] }

        $Info.uses = Get-TagValues $BlockLines 'use'
        $Info.replaces = Get-TagValues $BlockLines 'replace'
        $Info.references = Get-References $BlockLines
        $Info.tests = Get-TagValues $BlockLines 'test'
        $Info.examples = Get-TagValues $BlockLines 'example'
        $Info.executes = Get-TagValues $BlockLines 'execute'
        $Info.malformedTags = Get-MalformedTagLines $BlockLines
        $Info.headerSummary = Get-HeaderSummary $HeaderLines
        if ($Kind -ne 'package') { $Info.publicMembers = Get-PublicMembers $CodeLines }

        if ($Descriptions.Count -gt 0 -and $Descriptions[0]) {
            $Info.description = $Descriptions[0]
            $Info.descriptionSource = 'tag'
        }
        elseif ($Info.headerSummary) {
            $Info.description = $Info.headerSummary
            $Info.descriptionSource = 'header'
        }
    }
    return $Info
}

function Test-InCatalogScope {
    param([string]$RelativePath)
    $TopFolder = $RelativePath.Split('/')[0]
    return -not ($ExcludedTopFolders -contains $TopFolder)
}

function ConvertTo-PathKey {
    param([string]$Path)
    return $Path.Trim().Replace('\', '/').TrimStart('/').ToLowerInvariant()
}

#endregion

#region JSON writer (stable formatting, no PowerShell 5.1 escaping quirks)

function ConvertTo-JsonString {
    param([string]$Text)

    $Builder = New-Object System.Text.StringBuilder
    [void]$Builder.Append('"')
    foreach ($Char in $Text.ToCharArray()) {
        switch ($Char) {
            '"' { [void]$Builder.Append('\"'); continue }
            '\' { [void]$Builder.Append('\\'); continue }
            "`n" { [void]$Builder.Append('\n'); continue }
            "`r" { [void]$Builder.Append('\r'); continue }
            "`t" { [void]$Builder.Append('\t'); continue }
            default {
                if ([int]$Char -lt 0x20) { [void]$Builder.AppendFormat('\u{0:x4}', [int]$Char) }
                else { [void]$Builder.Append($Char) }
            }
        }
    }
    [void]$Builder.Append('"')
    return $Builder.ToString()
}

function ConvertTo-JsonText {
    param($Value, [int]$Level = 0)

    $Indent = '  ' * ($Level + 1)
    $ClosingIndent = '  ' * $Level

    if ($null -eq $Value) { return 'null' }
    if ($Value -is [bool]) { if ($Value) { return 'true' } else { return 'false' } }
    if ($Value -is [int] -or $Value -is [long]) { return $Value.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
    if ($Value -is [string]) { return ConvertTo-JsonString $Value }
    if ($Value -is [System.Collections.IDictionary]) {
        if ($Value.Count -eq 0) { return '{}' }
        $Items = foreach ($Key in $Value.Keys) {
            $Indent + (ConvertTo-JsonString ([string]$Key)) + ': ' + (ConvertTo-JsonText $Value[$Key] ($Level + 1))
        }
        return '{' + $NewLine + ($Items -join (',' + $NewLine)) + $NewLine + $ClosingIndent + '}'
    }
    if ($Value -is [System.Collections.IEnumerable]) {
        $Array = @($Value)
        if ($Array.Count -eq 0) { return '[]' }
        $Items = foreach ($Item in $Array) { ConvertTo-JsonText $Item ($Level + 1) }
        return '[' + $NewLine + (($Items | ForEach-Object { $Indent + $_ }) -join (',' + $NewLine)) + $NewLine + $ClosingIndent + ']'
    }
    return ConvertTo-JsonString ([string]$Value)
}

#endregion

#region Collect

$AllFiles = Get-ChildItem -Path $RepositoryRoot -Recurse -File |
    Where-Object {
        ($CodeFileExtensions -contains $_.Extension.ToLowerInvariant()) -and
        ($_.FullName -notmatch '\\\.git\\')
    } |
    Sort-Object FullName

$AllModules = foreach ($File in $AllFiles) { Read-ModuleInfo $File }

# lookup of all existing files (incl. excluded folders) for dependency resolution
$ExistingPaths = @{}
foreach ($Module in $AllModules) { $ExistingPaths[(ConvertTo-PathKey $Module.path)] = $Module }

$CatalogModules = @($AllModules | Where-Object { $_.hasCodeLibBlock -and (Test-InCatalogScope $_.path) })

# key = repository path as used in <use> (file tag, falls back to physical path)
$ByKey = @{}
foreach ($Module in $CatalogModules) {
    $ByKey[(ConvertTo-PathKey $Module.path)] = $Module
    if ($Module.fileTag) { $ByKey[(ConvertTo-PathKey $Module.fileTag)] = $Module }
}

function Get-TransitiveUses {
    param($Module)

    $Visited = [ordered]@{}
    $Stack = New-Object System.Collections.Stack
    foreach ($Use in $Module.uses) { $Stack.Push($Use) }
    while ($Stack.Count -gt 0) {
        $Use = $Stack.Pop()
        $Key = ConvertTo-PathKey $Use
        if ($Visited.Contains($Key)) { continue }
        $Target = $ByKey[$Key]
        if ($null -eq $Target -and $ExistingPaths.ContainsKey($Key)) { $Target = $ExistingPaths[$Key] }
        $Visited[$Key] = if ($Target) { $Target.path } else { $Use }
        if ($Target) { foreach ($Next in $Target.uses) { $Stack.Push($Next) } }
    }
    $Self = ConvertTo-PathKey $Module.path
    return , @($Visited.Keys | Where-Object { $_ -ne $Self } | ForEach-Object { $Visited[$_] } | Sort-Object)
}

$UsedBy = @{}
foreach ($Module in $CatalogModules) {
    foreach ($Use in $Module.uses) {
        $Key = ConvertTo-PathKey $Use
        if (-not $UsedBy.ContainsKey($Key)) { $UsedBy[$Key] = New-Object System.Collections.Generic.List[string] }
        if (-not $UsedBy[$Key].Contains($Module.path)) { $UsedBy[$Key].Add($Module.path) }
    }
}

$Issues = New-Object System.Collections.Generic.List[string]

function Test-Reference {
    param($Module, [string]$TagName, [string[]]$Paths)
    foreach ($Path in $Paths) {
        if (-not $Path -or $Path.StartsWith('%')) { continue }
        $Key = ConvertTo-PathKey $Path
        if (-not $ExistingPaths.ContainsKey($Key)) {
            $Issues.Add("``$($Module.path)``: <$TagName> target not found: ``$Path``")
        }
        elseif ($Path -cne $ExistingPaths[$Key].path) {
            $Issues.Add("``$($Module.path)``: <$TagName> path differs from file system spelling: ``$Path`` (expected ``$($ExistingPaths[$Key].path)``)")
        }
    }
}

# example modules are not in the catalog, but their headers are checked as well
$ExampleModules = @($AllModules | Where-Object { $_.hasCodeLibBlock -and $_.path.StartsWith('_examples/') })

# headers of all importable modules are checked (library, tests, examples, DotNetLib integration);
# add-in sources, templates and samples in _codelib and the tools in _tools are not checked
$CheckedTopFolders = @('_test', '_examples', '_dotnetlib')
$CheckedModules = @($AllModules | Where-Object {
    $_.hasCodeLibBlock -and ((Test-InCatalogScope $_.path) -or ($CheckedTopFolders -contains $_.path.Split('/')[0]))
})

foreach ($Module in $CheckedModules) {
    if ($Module.kind -ne 'package') {
        if (-not $Module.fileTag) {
            $Issues.Add("``$($Module.path)``: <file> tag missing")
        }
        elseif (-not $Module.fileTag.StartsWith('%') -and
                (ConvertTo-PathKey $Module.fileTag) -ne (ConvertTo-PathKey $Module.path)) {
            $Issues.Add("``$($Module.path)``: <file> tag differs from location: ``$($Module.fileTag)``")
        }
        elseif (-not $Module.fileTag.StartsWith('%') -and $Module.fileTag -cne $Module.path) {
            $Issues.Add("``$($Module.path)``: <file> tag differs from file system spelling: ``$($Module.fileTag)``")
        }
    }
    foreach ($Line in $Module.malformedTags) {
        $Issues.Add("``$($Module.path)``: malformed tag line: ``$Line``")
    }
    Test-Reference $Module 'use' $Module.uses
    Test-Reference $Module 'license' @($Module.license)
    Test-Reference $Module 'test' $Module.tests
    Test-Reference $Module 'example' $Module.examples
}

# variants: several files with the same object name (cannot be imported together, the description names the use case)
$SameObjectName = @{}
$Variants = New-Object System.Collections.Generic.List[string]
$CodeModulesByName = $CatalogModules | Where-Object { $_.kind -ne 'package' } | Group-Object { $_.name.ToLowerInvariant() } | Where-Object { $_.Count -gt 1 }
foreach ($Group in $CodeModulesByName) {
    foreach ($Module in $Group.Group) {
        $SameObjectName[(ConvertTo-PathKey $Module.path)] = @($Group.Group | Where-Object { $_ -ne $Module } | ForEach-Object { $_.path } | Sort-Object)
    }
    $Paths = ($Group.Group | ForEach-Object { "``$($_.path)``" }) -join ', '
    $Variants.Add("``$($Group.Group[0].name)``: $Paths")
}

$LinkedExamples = @{}
foreach ($Module in $CatalogModules) {
    foreach ($Example in $Module.examples) { $LinkedExamples[(ConvertTo-PathKey $Example)] = $true }
}
foreach ($Module in $ExampleModules) {
    if (-not $LinkedExamples.ContainsKey((ConvertTo-PathKey $Module.path))) {
        $Issues.Add("``$($Module.path)``: example not linked by an <example> tag of a library module")
    }
}
foreach ($Module in $CatalogModules) {
    if ($Module.name -match '_(Example|Examples|Beispiel|Beispiele)(_|$)' -or $Module.name -match 'Example$') {
        $Issues.Add("``$($Module.path)``: example module outside _examples")
    }
}

# folder names of library, tests and examples are lower case (GitHub paths are case-sensitive)
$UpperCaseFolders = @{}
foreach ($Module in $AllModules) {
    $Folder = $Module.path.Substring(0, [Math]::Max(0, $Module.path.LastIndexOf('/')))
    if (-not $Folder) { continue }
    $TopFolder = $Folder.Split('/')[0]
    if ((Test-InCatalogScope $Module.path) -or $TopFolder -eq '_test' -or $TopFolder -eq '_examples') {
        if ($Folder -cmatch '[A-Z]') { $UpperCaseFolders[$Folder] = $true }
    }
}
foreach ($Folder in ($UpperCaseFolders.Keys | Sort-Object)) {
    $Issues.Add("folder name not lower case: ``$Folder``")
}

# recipes: code file paths in backticks and links to other recipes must exist
$RecipeFolder = Join-Path $RepositoryRoot '_codelib\recipes'
if (Test-Path $RecipeFolder) {
    foreach ($Recipe in (Get-ChildItem -Path $RecipeFolder -Filter '*.md' | Sort-Object Name)) {
        $RecipeText = Read-SourceFile $Recipe.FullName
        $RecipePath = "_codelib/recipes/$($Recipe.Name)"
        foreach ($Match in [regex]::Matches($RecipeText, '`([A-Za-z0-9_./\\-]+\.(?:cls|bas|frm|rpt|package))`')) {
            $Path = $Match.Groups[1].Value
            $Key = ConvertTo-PathKey $Path
            if (-not $ExistingPaths.ContainsKey($Key)) {
                $Issues.Add("``$RecipePath``: file not found: ``$Path``")
            }
            elseif ($Path -cne $ExistingPaths[$Key].path) {
                $Issues.Add("``$RecipePath``: path differs from file system spelling: ``$Path`` (expected ``$($ExistingPaths[$Key].path)``)")
            }
        }
        foreach ($Match in [regex]::Matches($RecipeText, '\]\(([A-Za-z0-9_./-]+\.md)\)')) {
            if (-not (Test-Path (Join-Path $Recipe.DirectoryName $Match.Groups[1].Value))) {
                $Issues.Add("``$RecipePath``: link target not found: ``$($Match.Groups[1].Value)``")
            }
        }
    }
}

#endregion

#region Build output

function Get-FolderTitle {
    param([string]$RelativePath)
    $Parts = $RelativePath.Split('/')
    if ($Parts.Count -le 1) { return '(root)' }
    return ($Parts[0..($Parts.Count - 2)] -join '/')
}

$CodeModules = @($CatalogModules | Where-Object { $_.kind -ne 'package' } | Sort-Object { $_.path })
$Packages = @($CatalogModules | Where-Object { $_.kind -eq 'package' } | Sort-Object { $_.path })

# catalog.json
$JsonModules = foreach ($Module in $CodeModules) {
    $Key = ConvertTo-PathKey $Module.path
    [ordered]@{
        path          = $Module.path
        name          = $Module.name
        kind          = $Module.kind
        description   = $Module.description
        descriptionSource = $Module.descriptionSource
        predeclaredId = $Module.predeclaredId
        uses          = $Module.uses
        usesTransitive = (Get-TransitiveUses $Module)
        usedBy        = @(if ($UsedBy.ContainsKey($Key)) { $UsedBy[$Key] | Sort-Object })
        replaces      = $Module.replaces
        sameObjectName = @(if ($SameObjectName.ContainsKey($Key)) { $SameObjectName[$Key] })
        references    = $Module.references
        tests         = $Module.tests
        examples      = $Module.examples
        executes      = $Module.executes
        publicMembers = $Module.publicMembers
    }
}
$JsonPackages = foreach ($Package in $Packages) {
    [ordered]@{
        path           = $Package.path
        name           = $Package.packageName
        description    = $Package.description
        uses           = $Package.uses
        usesTransitive = (Get-TransitiveUses $Package)
        references     = $Package.references
        executes       = $Package.executes
    }
}
$CatalogJson = [ordered]@{
    generator = '_tools/catalog/Build-Catalog.ps1'
    note      = 'Generated file - do not edit manually.'
    modules   = @($JsonModules)
    packages  = @($JsonPackages)
}
$CatalogJsonText = (ConvertTo-JsonText $CatalogJson) + $NewLine

# catalog.md
$Md = New-Object System.Collections.Generic.List[string]
$Md.Add('# AccessCodeLib catalog')
$Md.Add('')
$Md.Add('> Generated by `_tools/catalog/Build-Catalog.ps1` from the `<codelib>` header blocks - do not edit manually.')
$Md.Add('> Details (dependencies, used-by, references, tests, public members): `catalog.json`.')
$Md.Add('> Task-oriented guides: `../recipes/`. Import with the ACLib Import Wizard (path = repository path below).')
$Md.Add('')
$Md.Add("Modules: $($CodeModules.Count) | Packages: $($Packages.Count)")
$Md.Add('')
$Md.Add('Legend: `uses` = direct dependencies (object names), `[P]` = predeclared class (usable without `New`),')
$Md.Add('`same object name as` = variant, cannot be imported together (see description for the use case).')
foreach ($Group in ($CodeModules | Group-Object { Get-FolderTitle $_.path } | Sort-Object Name)) {
    $Md.Add('')
    $Md.Add("## $($Group.Name)")
    $Md.Add('')
    foreach ($Module in $Group.Group) {
        $Line = "- ``$($Module.path)`` ($($Module.kind))"
        if ($Module.predeclaredId) { $Line += ' [P]' }
        if ($Module.description) { $Line += " - $($Module.description)" }
        if ($Module.uses.Count -gt 0) {
            $UseNames = foreach ($Use in $Module.uses) {
                $Target = $ByKey[(ConvertTo-PathKey $Use)]
                if ($Target) { $Target.name } else { [System.IO.Path]::GetFileNameWithoutExtension($Use) }
            }
            $Line += " | uses: $($UseNames -join ', ')"
        }
        $SameNameKey = ConvertTo-PathKey $Module.path
        if ($SameObjectName.ContainsKey($SameNameKey)) {
            $SameNamePaths = foreach ($SameNamePath in $SameObjectName[$SameNameKey]) { '`' + $SameNamePath + '`' }
            $Line += " | same object name as: $($SameNamePaths -join ', ')"
        }
        $Md.Add($Line)
    }
}
if ($Packages.Count -gt 0) {
    $Md.Add('')
    $Md.Add('## Packages')
    $Md.Add('')
    foreach ($Package in $Packages) {
        $Line = "- ``$($Package.path)`` - $($Package.packageName)"
        if ($Package.description) { $Line += ": $($Package.description)" }
        $Md.Add($Line)
    }
}
$CatalogMdText = ($Md -join $NewLine) + $NewLine

# report.md
$WithTag = @($CodeModules | Where-Object { $_.descriptionSource -eq 'tag' }).Count
$HeaderOnly = @($CodeModules | Where-Object { $_.descriptionSource -eq 'header' })
$NoDescription = @($CodeModules | Where-Object { $_.descriptionSource -eq 'none' })

$Report = New-Object System.Collections.Generic.List[string]
$Report.Add('# Catalog report')
$Report.Add('')
$Report.Add('> Generated by `_tools/catalog/Build-Catalog.ps1` - do not edit manually.')
$Report.Add('')
$Report.Add('| | Count |')
$Report.Add('|---|---|')
$Report.Add("| Modules in catalog | $($CodeModules.Count) |")
$Report.Add("| with ``<description>`` tag | $WithTag |")
$Report.Add("| header summary only | $($HeaderOnly.Count) |")
$Report.Add("| no description | $($NoDescription.Count) |")
$Report.Add("| Header issues | $($Issues.Count) |")
$Report.Add('')
$Report.Add('## Header issues')
$Report.Add('')
if ($Issues.Count -eq 0) { $Report.Add('None.') }
foreach ($Issue in ($Issues | Sort-Object)) { $Report.Add("- $Issue") }
$Report.Add('')
$Report.Add('## Variants (same object name)')
$Report.Add('')
$Report.Add('Cannot be imported together; the description names the use case of each file.')
$Report.Add('')
if ($Variants.Count -eq 0) { $Report.Add('None.') }
foreach ($Variant in ($Variants | Sort-Object)) { $Report.Add("- $Variant") }
$Report.Add('')
$Report.Add('## Modules without `<description>` tag')
foreach ($Group in (@($HeaderOnly + $NoDescription) | Sort-Object { $_.path } | Group-Object { Get-FolderTitle $_.path } | Sort-Object Name)) {
    $Report.Add('')
    $Report.Add("### $($Group.Name)")
    $Report.Add('')
    foreach ($Module in $Group.Group) {
        if ($Module.headerSummary) { $Report.Add("- ``$($Module.path)`` - header: $($Module.headerSummary)") }
        else { $Report.Add("- ``$($Module.path)`` - (no header summary)") }
    }
}
$ReportMdText = ($Report -join $NewLine) + $NewLine

#endregion

#region Write / check

$Outputs = [ordered]@{
    'catalog.md'   = $CatalogMdText
    'catalog.json' = $CatalogJsonText
    'report.md'    = $ReportMdText
}
$Utf8NoBom = New-Object System.Text.UTF8Encoding($false)

if ($Check) {
    $OutOfDate = New-Object System.Collections.Generic.List[string]
    foreach ($FileName in $Outputs.Keys) {
        $Path = Join-Path $OutputDirectory $FileName
        if (-not (Test-Path $Path)) { $OutOfDate.Add($FileName); continue }
        $Current = [System.IO.File]::ReadAllText($Path, $Utf8NoBom)
        if ($Current -ne $Outputs[$FileName]) { $OutOfDate.Add($FileName) }
    }
    Write-Host "Modules: $($CodeModules.Count), with <description>: $WithTag, header only: $($HeaderOnly.Count), none: $($NoDescription.Count), issues: $($Issues.Count)"
    if ($OutOfDate.Count -gt 0) {
        Write-Host "Out of date: $($OutOfDate -join ', ') - run _tools/catalog/Build-Catalog.ps1"
        exit 1
    }
    Write-Host 'Catalog is up to date.'
    exit 0
}

if (-not (Test-Path $OutputDirectory)) { New-Item -ItemType Directory -Path $OutputDirectory | Out-Null }
foreach ($FileName in $Outputs.Keys) {
    [System.IO.File]::WriteAllText((Join-Path $OutputDirectory $FileName), $Outputs[$FileName], $Utf8NoBom)
}
Write-Host "Catalog written to $OutputDirectory"
Write-Host "Modules: $($CodeModules.Count), packages: $($Packages.Count), with <description>: $WithTag, header only: $($HeaderOnly.Count), none: $($NoDescription.Count), issues: $($Issues.Count)"

#endregion
