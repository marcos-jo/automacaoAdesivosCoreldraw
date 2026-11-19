# MIG-004B1: diagnostics only. No application, package manager or build is started.
[CmdletBinding()]
param(
    [ValidateSet('Diagnostic')]
    [string] $Mode = 'Diagnostic'
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

$parentBefore = Get-EnvironmentSnapshot Process
$userBefore = Get-EnvironmentSnapshot User
$machineBefore = Get-EnvironmentSnapshot Machine
try {
    $appRoot = Split-Path -Path $PSScriptRoot -Parent
    $repoRoot = Split-Path -Path $appRoot -Parent
    $devRoot = Split-Path -Path $repoRoot -Parent
    $volumeRoot = [IO.Path]::GetPathRoot($repoRoot)
    if ($volumeRoot -notmatch '^[A-Za-z]:\\$') { throw 'A local drive is required.' }
    Assert-PlainPath $appRoot
    if (-not (Test-Path -LiteralPath (Join-Path $repoRoot '.git'))) {
        throw "Git repository marker missing: $repoRoot"
    }
    $toolsRoot = Join-Path $devRoot '4PrintShops-Tools'
    $nodeRoot = Join-Path $toolsRoot 'node-v26.13.1-win-x64'
    $node = Join-Path $nodeRoot 'node.exe'
    $buildsRoot = Join-Path $devRoot '4PrintShops-Builds'
    $candidates = @(Get-ChildItem -LiteralPath $buildsRoot -Directory -Filter 'experimental.42-*' |
        Where-Object { Test-Path -LiteralPath (Join-Path $_.FullName 'source/node_modules/electron/package.json') })
    if ($candidates.Count -ne 1) { throw 'Expected exactly one prepared experimental.42 source; review candidates manually.' }
    $isolatedSource = Join-Path $candidates[0].FullName 'source'
    $modules = Join-Path $isolatedSource 'node_modules'
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

    # Report file/line only, never matching text or arbitrary environment values.
    $scanFiles = @()
    foreach ($folder in @('scripts', 'tests', 'desktop', 'src')) {
        $scanFiles += @(Get-ChildItem -LiteralPath (Join-Path $appRoot $folder) -Recurse -File |
            Where-Object { $_.Extension -in @('.ps1', '.mjs', '.cjs', '.js', '.nsh', '.nsi', '.cmd') })
    }
    $scanFiles += @(Get-ChildItem -LiteralPath $appRoot -File -Filter '*.cjs')
    $scanFiles += @(Get-ChildItem -LiteralPath $candidates[0].FullName -File |
        Where-Object { $_.Extension -in @('.mjs', '.ps1', '.cjs', '.cmd') })
    $scanFiles += @(Get-ChildItem -LiteralPath $toolsRoot -File |
        Where-Object { $_.Extension -in @('.mjs', '.ps1', '.cjs', '.cmd') })
    $references = @($scanFiles | Select-String -Pattern '(?i)c:[\\/]+Dev[\\/]')
    foreach ($reference in $references) {
        Write-Output ("Previous-location reference: {0}:{1}" -f $reference.Path, $reference.LineNumber)
    }
    Write-Output "Previous-location references: $($references.Count). Documentation, dependencies and archived nested evidence are outside this scan."
    Write-Output 'Historical caches were not copied or overwritten. Build toolchains remain unqualified.'
    Write-Output 'No application, Electron, npm, build, download or functional test was started.'
} finally {
    Assert-EnvironmentUnchanged $parentBefore Process
    Assert-EnvironmentUnchanged $userBefore User
    Assert-EnvironmentUnchanged $machineBefore Machine
    Write-Output 'PASS: parent, User and Machine environment snapshots unchanged.'
}
