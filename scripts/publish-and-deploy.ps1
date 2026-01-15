[CmdletBinding()]
param(
    [ValidateSet('Debug', 'Release')]
    [string]$Configuration = 'Release',

    [string]$RuntimeIdentifier = 'win-x86',

    [string]$ProjectPath = 'GlobalHotkeys/GlobalHotkeys.csproj',

    [string]$InstallDir = 'C:\Program Files\GlobalHotkeys',

    [string]$ExeName = 'GlobalHotkeys.exe',

    [switch]$NoRestart
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsAdmin {
    $id = [Security.Principal.WindowsIdentity]::GetCurrent()
    $p = New-Object Security.Principal.WindowsPrincipal($id)
    return $p.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

if (-not (Test-IsAdmin)) {
    $self = $MyInvocation.MyCommand.Path
    $argList = @(
        '-NoProfile',
        '-ExecutionPolicy', 'Bypass',
        '-File', "`"$self`"",
        '-Configuration', $Configuration,
        '-RuntimeIdentifier', $RuntimeIdentifier,
        '-ProjectPath', "`"$ProjectPath`"",
        '-InstallDir', "`"$InstallDir`"",
        '-ExeName', "`"$ExeName`""
    )
    if ($NoRestart) { $argList += '-NoRestart' }

    Start-Process -FilePath 'pwsh' -Verb RunAs -ArgumentList $argList | Out-Null
    return
}

$repoRoot = Split-Path -Parent $PSScriptRoot
$fullProjectPath = Join-Path $repoRoot $ProjectPath
if (-not (Test-Path $fullProjectPath)) {
    throw "Project not found: $fullProjectPath"
}

Write-Host "Publishing single-file exe..." -ForegroundColor Cyan
Write-Host "  Project: $fullProjectPath"
Write-Host "  Config : $Configuration"
Write-Host "  RID    : $RuntimeIdentifier"

$publishRoot = Join-Path $repoRoot 'publish'
$tempPublishDir = Join-Path $publishRoot '.publish-tmp'
if (Test-Path $tempPublishDir) {
    Remove-Item -Recurse -Force $tempPublishDir
}
New-Item -ItemType Directory -Path $tempPublishDir | Out-Null

& dotnet publish $fullProjectPath `
    -c $Configuration `
    -r $RuntimeIdentifier `
    -p:SelfContained=true `
    -p:PublishSingleFile=true `
    -p:IncludeNativeLibrariesForSelfExtract=true `
    -p:EnableCompressionInSingleFile=true `
    -p:DebugType=None `
    -p:DebugSymbols=false `
    -o $tempPublishDir

$publishedExe = Join-Path $tempPublishDir $ExeName
if (-not (Test-Path $publishedExe)) {
    # Fallback: if AssemblyName differs, take the first exe found.
    $exeCandidates = Get-ChildItem -Path $tempPublishDir -Filter '*.exe' -File | Select-Object -First 1
    if (-not $exeCandidates) {
        throw "Publish completed but no .exe found in: $tempPublishDir"
    }
    $publishedExe = $exeCandidates.FullName
}

Write-Host "Deploying to: $InstallDir" -ForegroundColor Cyan
New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null

$processName = [IO.Path]::GetFileNameWithoutExtension($ExeName)
Get-Process -Name $processName -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Milliseconds 500

$destExe = Join-Path $InstallDir $ExeName
Copy-Item -Force $publishedExe $destExe

Write-Host "OK: $destExe" -ForegroundColor Green

if (-not $NoRestart) {
    Start-Process -FilePath $destExe -WorkingDirectory $InstallDir | Out-Null
    Write-Host "Restarted: $ExeName" -ForegroundColor Green
}

Remove-Item -Recurse -Force $tempPublishDir
