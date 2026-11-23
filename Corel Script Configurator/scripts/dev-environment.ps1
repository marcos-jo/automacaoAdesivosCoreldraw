# Diagnostics and a fixed MIG-004B2B3 Node-only check set. No application or build.
[CmdletBinding()]
param(
    [ValidateSet('Diagnostic', 'MigrationChecks')]
    [string] $Mode = 'Diagnostic',
    [string] $SourcePath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-EnvironmentSnapshot {
    param([System.EnvironmentVariableTarget] $Target)
    $snapshot = @{}
    foreach ($entry in [Environment]::GetEnvironmentVariables($Target).GetEnumerator()) {
        $snapshot[$entry.Key] = $entry.Value
    }
    return ,$snapshot
}

function Assert-EnvironmentUnchanged {
    param([hashtable] $Before, [System.EnvironmentVariableTarget] $Target)
    $after = Get-EnvironmentSnapshot $Target
    if ($Before.Count -ne $after.Count) { throw "Environment changed: $Target" }
    foreach ($key in $Before.Keys) {
        if (-not $after.ContainsKey($key) -or $Before[$key] -cne $after[$key]) {
            throw "Environment changed: $Target (values omitted)."
        }
    }
}

function Assert-PlainPath {
    param([string] $Path)
    # Reject existing reparse points in the entire ancestry, including cache parents.
    $cursor = [IO.Path]::GetFullPath($Path)
    while ($cursor) {
        if (Test-Path -LiteralPath $cursor) {
            $item = Get-Item -LiteralPath $cursor -Force
            if ($item.Attributes -band [IO.FileAttributes]::ReparsePoint) {
                throw "Reparse point requires separate review: $cursor"
            }
        }
        $cursor = Split-Path -Path $cursor -Parent
    }
}

function Ensure-Directory {
    param([string] $Path)
    Assert-PlainPath $Path
    if (Test-Path -LiteralPath $Path) {
        if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
            throw "Expected directory: $Path"
        }
    } else {
        [void][IO.Directory]::CreateDirectory($Path)
    }
}

function Invoke-DiagnosticNode {
    param([string] $Arguments, [string] $InputText = '')
    $start = New-Object System.Diagnostics.ProcessStartInfo
    $start.FileName = $node
    $start.Arguments = $Arguments
    $start.WorkingDirectory = $appRoot
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    $start.RedirectStandardInput = $true
    foreach ($key in $childEnvironment.Keys) {
        $start.EnvironmentVariables[$key] = $childEnvironment[$key]
    }
    [void]$start.EnvironmentVariables.Remove('ELECTRON_RUN_AS_NODE')
    [void]$start.EnvironmentVariables.Remove('NODE_OPTIONS')
    [void]$start.EnvironmentVariables.Remove('NODE_COMPILE_CACHE')
    [void]$start.EnvironmentVariables.Remove('NODE_V8_COVERAGE')
    $process = New-Object System.Diagnostics.Process
    $process.StartInfo = $start
    try {
        [void]$process.Start()
        if ($InputText) { $process.StandardInput.Write($InputText) }
        $process.StandardInput.Close()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        if (-not $process.WaitForExit(15000)) {
            $process.Kill()
            $process.WaitForExit()
            throw 'Diagnostic Node exceeded 15 seconds; only its own process was stopped.'
        }
        if ($process.ExitCode -ne 0 -or $stderr.Result.Length -gt 0) {
            throw 'Diagnostic Node failed (child output omitted to protect environment data).'
        }
        return $stdout.Result.Trim()
    } finally {
        $process.Dispose()
    }
}

function Get-DevelopmentFreeSpace {
    $systemDrive = New-Object IO.DriveInfo ([IO.Path]::GetPathRoot($env:SystemRoot))
    $projectDrive = New-Object IO.DriveInfo $volumeRoot
    return @{ SystemBytes = $systemDrive.AvailableFreeSpace; ProjectBytes = $projectDrive.AvailableFreeSpace }
}

function Invoke-MigrationChecks {
    # Fixed reviewed list only: in-memory fixtures, no Electron, network or filesystem writes.
    $tests = @('tests/pdf-formats.test.mjs', 'tests/pdf-format-catalog.test.mjs', 'tests/pdf-sheet-orientation.test.mjs')
    $testPaths = @($tests | ForEach-Object { Join-Path $isolatedSource $_ })
    foreach ($file in $testPaths) {
        Assert-PlainPath $file
        if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { throw 'Reviewed migration test missing.' }
    }
    $reviewedHashes = @($testPaths | ForEach-Object { Get-FileHash -LiteralPath $_ -Algorithm SHA256 } |
        Select-Object Path, Hash)
    $before = Get-DevelopmentFreeSpace
    if ($before.SystemBytes -lt 10GB) { throw 'Space guard: system drive has less than 10 GiB; checks not started.' }
    $results = Join-Path $runRoot 'results'
    Ensure-Directory $results
    $start = New-Object Diagnostics.ProcessStartInfo
    $start.FileName = $node
    $arguments = @('--test', '--test-concurrency=1', '--test-reporter=tap') + $testPaths
    $start.Arguments = ($arguments | ForEach-Object { '"' + $_ + '"' }) -join ' '
    $start.WorkingDirectory = $isolatedSource
    $start.UseShellExecute = $false
    $start.CreateNoWindow = $true
    $start.RedirectStandardOutput = $true
    $start.RedirectStandardError = $true
    foreach ($key in $childEnvironment.Keys) { $start.EnvironmentVariables[$key] = $childEnvironment[$key] }
    foreach ($key in @('ELECTRON_RUN_AS_NODE', 'NODE_OPTIONS', 'NODE_COMPILE_CACHE', 'NODE_V8_COVERAGE')) {
        [void]$start.EnvironmentVariables.Remove($key)
    }
    $process = New-Object Diagnostics.Process
    $process.StartInfo = $start
    $began = [DateTime]::UtcNow.ToString('o')
    try {
        [void]$process.Start()
        $stdout = $process.StandardOutput.ReadToEndAsync()
        $stderr = $process.StandardError.ReadToEndAsync()
        $timedOut = -not $process.WaitForExit(60000)
        if ($timedOut) { $process.Kill(); $process.WaitForExit() }
        $exitCode = $process.ExitCode
        [IO.File]::WriteAllText((Join-Path $results 'stdout.tap'), $stdout.Result)
        [IO.File]::WriteAllText((Join-Path $results 'stderr.txt'), $stderr.Result)
    } finally { $process.Dispose() }
    $after = Get-DevelopmentFreeSpace
    $spaceStopped = $after.SystemBytes -lt 10GB -or ($before.SystemBytes - $after.SystemBytes) -ge 256MB
    $record = [ordered]@{
        stage = 'MIG-004B2B3'; node = $node; nodeVersion = 'v26.13.1'; arguments = $arguments; cwd = $isolatedSource
        testHashes = $reviewedHashes
        startedUtc = $began; finishedUtc = [DateTime]::UtcNow.ToString('o')
        exitCode = $exitCode; timedOut = $timedOut; before = $before; after = $after
        spaceGuardStopped = $spaceStopped; significantSystemDropBytes = 256MB
        temp = $runRoot; appBuilderTemp = $childEnvironment.APP_BUILDER_TMP_DIR
        npmCache = $childEnvironment.npm_config_cache; builderCache = $childEnvironment.ELECTRON_BUILDER_CACHE
        electronCache = $childEnvironment.electron_config_cache; dependencyRoot = $modules
    }
    [IO.File]::WriteAllText((Join-Path $results 'execution.json'), ($record | ConvertTo-Json -Depth 5))
    Write-Output "Migration check results: $results"
    Write-Output ("Space before: system={0}; project={1} bytes" -f $before.SystemBytes, $before.ProjectBytes)
    Write-Output ("Space after: system={0}; project={1} bytes" -f $after.SystemBytes, $after.ProjectBytes)
    Write-Output ("Fixed Node checks exit: {0}; timeout: {1}; space guard stopped: {2}" -f $exitCode, $timedOut, $spaceStopped)
    if ($spaceStopped) { throw 'Space guard: no further executions permitted in this run; see execution.json.' }
    if ($timedOut -or $exitCode -ne 0) { throw 'Fixed migration checks failed; see recorded output. No automatic retry.' }
}

$parentBefore = Get-EnvironmentSnapshot Process
$userBefore = Get-EnvironmentSnapshot User
$machineBefore = Get-EnvironmentSnapshot Machine
try {
    $appRoot = Split-Path -Path $PSScriptRoot -Parent
    $repoRoot = Split-Path -Path $appRoot -Parent
    $devRoot = Split-Path -Path $repoRoot -Parent
    $volumeRoot = [IO.Path]::GetPathRoot($repoRoot)
    if ($volumeRoot -notmatch '^[A-Za-z]:\\$') { throw 'A local drive is required.' }
    if ($Mode -eq 'MigrationChecks' -and (Get-DevelopmentFreeSpace).SystemBytes -lt 10GB) {
        throw 'Space guard: system drive has less than 10 GiB; diagnostics and checks not started.'
    }
    Assert-PlainPath $appRoot
    if (-not (Test-Path -LiteralPath (Join-Path $repoRoot '.git'))) {
        throw "Git repository marker missing: $repoRoot"
    }
    $toolsRoot = Join-Path $devRoot '4PrintShops-Tools'
    $nodeRoot = Join-Path $toolsRoot 'node-v26.13.1-win-x64'
    $node = Join-Path $nodeRoot 'node.exe'
    $buildsRoot = Join-Path $devRoot '4PrintShops-Builds'
    if ($Mode -eq 'MigrationChecks' -and -not $PSBoundParameters.ContainsKey('SourcePath')) {
        throw 'MigrationChecks requires an explicit independent SourcePath; no historical default.'
    }
    if ($PSBoundParameters.ContainsKey('SourcePath')) {
        if ([string]::IsNullOrWhiteSpace($SourcePath) -or $SourcePath -notmatch '^(?:[A-Za-z]:[\\/]|\\\\)') {
            throw 'SourcePath must be a full filesystem directory path.'
        }
        $selectedSource = $SourcePath
        $selection = 'explicit'
    } else {
        $candidates = @(Get-ChildItem -LiteralPath $buildsRoot -Directory -Filter 'experimental.42-*' |
            Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'source/node_modules/electron/package.json') })
        if ($candidates.Count -ne 1) { throw 'Expected exactly one prepared experimental.42 source; review candidates manually.' }
        $selectedSource = Join-Path $candidates[0].FullName 'source'
        $selection = 'default'
    }
    # Explicit selection never falls back. Resolve literal filesystem paths and reject links.
    if (-not (Test-Path -LiteralPath $selectedSource -PathType Container)) {
        throw 'Selected source directory does not exist.'
    }
    Assert-PlainPath $selectedSource
    $resolvedSource = Resolve-Path -LiteralPath $selectedSource -ErrorAction Stop
    if ($resolvedSource.Provider.Name -ne 'FileSystem') { throw 'Selected source must use the filesystem provider.' }
    $isolatedSource = [IO.Path]::GetFullPath($resolvedSource.ProviderPath)
    Assert-PlainPath $isolatedSource
    foreach ($manifest in @('package.json', 'package-lock.json')) {
        $manifestPath = Join-Path $isolatedSource $manifest
        Assert-PlainPath $manifestPath
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) { throw "Selected source manifest missing: $manifest" }
    }
    $sourcePackage = Get-Content -Raw -LiteralPath (Join-Path $isolatedSource 'package.json') | ConvertFrom-Json
    # Windows PowerShell 5.1 ConvertFrom-Json rejects the lockfile's empty package key.
    Add-Type -AssemblyName System.Web.Extensions
    $lockReader = New-Object System.Web.Script.Serialization.JavaScriptSerializer
    $sourceLock = $lockReader.DeserializeObject([IO.File]::ReadAllText((Join-Path $isolatedSource 'package-lock.json')))
    if ($sourcePackage.name -ne 'corel-script-configurator' -or -not ($sourceLock['packages'] -is [Collections.IDictionary])) {
        throw 'Selected source manifests do not describe the expected project.'
    }
    $modules = Join-Path $isolatedSource 'node_modules'
    foreach ($dependency in @('electron', 'electron-builder', 'app-builder-lib', 'electron-updater', '@electron/asar', 'builder-util-runtime', 'pdfjs-dist', 'pdf-lib')) {
        $dependencyManifest = Join-Path $modules ($dependency + '/package.json')
        Assert-PlainPath $dependencyManifest
        if (-not (Test-Path -LiteralPath $dependencyManifest -PathType Leaf)) { throw "Selected source dependency missing: $dependency" }
        $package = Get-Content -Raw -LiteralPath $dependencyManifest | ConvertFrom-Json
        $locked = $sourceLock['packages']['node_modules/' + $dependency]
        if ($null -eq $locked -or $package.name -ne $dependency -or $package.version -ne $locked['version']) {
            throw "Selected source dependency differs from lockfile: $dependency"
        }
        Write-Output ("Dependency present: {0} {1}" -f $dependency, $package.version)
    }
    $tempRoot = Join-Path $volumeRoot 'Temp/4PrintShops'
    $cacheRoot = Join-Path $volumeRoot 'Cache/4PrintShops'
    $runRoot = Join-Path $tempRoot ('dev-' + [Guid]::NewGuid().ToString('N'))
    $resources = [ordered]@{
        Node = $node
        NpmCli = Join-Path $nodeRoot 'node_modules/npm/bin/npm-cli.js'
        Electron = Join-Path $modules 'electron/dist/electron.exe'
        BuilderCli = Join-Path $modules 'electron-builder/cli.js'
        Office = Join-Path $isolatedSource 'build/libreoffice/runtime/program/soffice.com'
        OfficeManifest = Join-Path $isolatedSource 'build/libreoffice/runtime/runtime-manifest.json'
        OfficeTrust = Join-Path $isolatedSource 'build/libreoffice/trust.json'
        Pdfium = Join-Path $isolatedSource 'build/pdfium/bin/pdfium.dll'
        NativeHelper = Join-Path $isolatedSource 'build/native/4PrintShops.PrintConfig.exe'
        HoldHelper = Join-Path $isolatedSource 'build/hold-native/4PrintShops.PrintConfig.exe'
        HistoricalBuilderCache = Join-Path $toolsRoot 'desktop-prep-2026-11-16/builder-cache'
        WindowsCompiler = Join-Path $env:SystemRoot 'Microsoft.NET/Framework64/v4.0.30319/csc.exe'
    }
    Write-Output "Mode: $Mode"
    Write-Output "Repository: $repoRoot"
    Write-Output "Application: $appRoot"
    Write-Output "Tools: $toolsRoot"
    Write-Output "Source selection: $selection"
    Write-Output "Prepared source: $isolatedSource"
    foreach ($entry in $resources.GetEnumerator()) {
        Assert-PlainPath $entry.Value
        if (-not (Test-Path -LiteralPath $entry.Value)) { throw "Required resource missing: $($entry.Value)" }
        Write-Output ("Resource present [{0}]: {1}" -f $entry.Key, $entry.Value)
    }
    $childEnvironment = @{
        TEMP = $runRoot
        TMP = $runRoot
        APP_BUILDER_TMP_DIR = Join-Path $runRoot 'builder'
        npm_config_cache = Join-Path $cacheRoot 'npm'
        ELECTRON_BUILDER_CACHE = Join-Path $cacheRoot 'electron-builder'
        electron_config_cache = Join-Path $cacheRoot 'electron'
        PATH = $nodeRoot + [IO.Path]::PathSeparator + $env:PATH
        DESKTOP_TEST_MODULES = $modules
        ELECTRON_PATH = $resources.Electron
        NODE_DISABLE_COMPILE_CACHE = '1'
    }
    # Only create empty development directories; never seed, replace or delete caches.
    foreach ($key in @('TEMP', 'APP_BUILDER_TMP_DIR', 'npm_config_cache', 'ELECTRON_BUILDER_CACHE', 'electron_config_cache')) {
        Ensure-Directory $childEnvironment[$key]
        Write-Output ("Child {0}: {1}" -f $key, $childEnvironment[$key])
    }
    Write-Output "Child PATH prefix: $nodeRoot (remaining entries omitted)"
    Write-Output 'Child ELECTRON_RUN_AS_NODE and NODE_OPTIONS: removed; inherited values omitted.'
    $version = Invoke-DiagnosticNode '--version'
    if ($version -cne 'v26.13.1') { throw 'Unexpected portable Node version.' }
    Write-Output "Portable Node: $version"

    # Input goes over stdin, not through a shell. No arbitrary command mode is exposed.
    $expected = @{}
    foreach ($key in $childEnvironment.Keys) { $expected[$key] = $childEnvironment[$key] }
    $expected['APPDATA'] = $parentBefore['APPDATA']
    $expected['LOCALAPPDATA'] = $parentBefore['LOCALAPPDATA']
    $payload = @{ expected = $expected; node = $node; cwd = $appRoot } | ConvertTo-Json -Depth 4 -Compress
    $probe = @'
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const expected = PAYLOAD;
for (const [key, value] of Object.entries(expected.expected)) {
  assert.equal(process.env[key] ?? null, value);
}
assert.equal(process.env.ELECTRON_RUN_AS_NODE, undefined);
assert.equal(process.env.NODE_OPTIONS, undefined);
assert.equal(process.version, 'v26.13.1');
assert.equal(path.resolve(process.execPath).toLowerCase(), path.resolve(expected.node).toLowerCase());
assert.equal(process.cwd().toLowerCase(), expected.cwd.toLowerCase());
assert.equal(process.env.PATH.split(path.delimiter)[0].toLowerCase(), path.dirname(expected.node).toLowerCase());
for (const key of ['TEMP', 'TMP', 'APP_BUILDER_TMP_DIR', 'npm_config_cache', 'ELECTRON_BUILDER_CACHE', 'electron_config_cache']) {
  assert.ok(fs.statSync(process.env[key]).isDirectory());
}
console.log('PASS: child paths, Node, cwd, PATH prefix, filtered variables and unchanged APPDATA/LOCALAPPDATA.');
'@
    Write-Output (Invoke-DiagnosticNode '-' ($probe.Replace('PAYLOAD', $payload)))
    if ($Mode -eq 'MigrationChecks') { Invoke-MigrationChecks }

    # Report file/line only, never matching text or arbitrary environment values.
    $scanFiles = @()
    foreach ($folder in @('scripts', 'tests', 'desktop', 'src')) {
        $scanFiles += @(Get-ChildItem -LiteralPath (Join-Path $appRoot $folder) -Recurse -File |
            Where-Object { $_.Extension -in @('.ps1', '.mjs', '.cjs', '.js', '.nsh', '.nsi', '.cmd') })
    }
    $scanFiles += @(Get-ChildItem -LiteralPath $appRoot -File -Filter '*.cjs')
    $scanFiles += @(Get-ChildItem -LiteralPath (Split-Path -Path $isolatedSource -Parent) -File |
        Where-Object { $_.Extension -in @('.mjs', '.ps1', '.cjs', '.cmd') })
    $scanFiles += @(Get-ChildItem -LiteralPath $toolsRoot -File |
        Where-Object { $_.Extension -in @('.mjs', '.ps1', '.cjs', '.cmd') })
    $references = @($scanFiles | Select-String -Pattern '(?i)c:[\\/]+Dev[\\/]')
    foreach ($reference in $references) {
        Write-Output ("Previous-location reference: {0}:{1}" -f $reference.Path, $reference.LineNumber)
    }
    Write-Output "Previous-location references: $($references.Count). Documentation, dependencies and archived nested evidence are outside this scan."
    Write-Output 'Historical caches were not copied or overwritten. Build toolchains remain unqualified.'
    if ($Mode -eq 'Diagnostic') { Write-Output 'No tests were started in Diagnostic mode.' }
    Write-Output 'No application, Electron, npm, build, download or browser test was started.'
} finally {
    Assert-EnvironmentUnchanged $parentBefore Process
    Assert-EnvironmentUnchanged $userBefore User
    Assert-EnvironmentUnchanged $machineBefore Machine
    Write-Output 'PASS: parent, User and Machine environment snapshots unchanged.'
}
