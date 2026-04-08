[CmdletBinding()]
param(
    [string]$Entry = "nycu.v2.ahk",
    [switch]$Compile,
    [string]$AhkExe = "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe",
    [string]$AhkCompiler = "C:\Program Files\AutoHotkey\Compiler\Ahk2Exe.exe",
    [string]$OutputDir
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

if ([System.IO.Path]::IsPathRooted($Entry)) {
    $entryPath = $Entry
} else {
    $entryPath = Join-Path $repoRoot $Entry
}

$entryPath = [System.IO.Path]::GetFullPath($entryPath)

if (-not (Test-Path -LiteralPath $entryPath)) {
    Write-Error "Entry script not found: $entryPath"
}

if (-not (Test-Path -LiteralPath $AhkExe)) {
    Write-Error "AutoHotkey executable not found: $AhkExe"
}

Write-Host "Validating $entryPath"
$validateProcess = Start-Process -FilePath $AhkExe -ArgumentList "/Validate", $entryPath -Wait -PassThru
$validateExitCode = $validateProcess.ExitCode

if ($validateExitCode -ne 0) {
    Write-Host "Validation failed with exit code $validateExitCode"
    exit $validateExitCode
}

Write-Host "Validation passed."

if (-not $Compile) {
    exit 0
}

if (-not (Test-Path -LiteralPath $AhkCompiler)) {
    Write-Error "AHK compiler not found: $AhkCompiler"
}

if ([string]::IsNullOrWhiteSpace($OutputDir)) {
    $OutputDir = Join-Path $env:TEMP ("ahk-compile-check-" + [DateTime]::Now.ToString("yyyyMMdd-HHmmss"))
}

$OutputDir = [System.IO.Path]::GetFullPath($OutputDir)
New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null

$outputExe = Join-Path $OutputDir (([System.IO.Path]::GetFileNameWithoutExtension($entryPath)) + ".exe")

Write-Host "Compiling to $outputExe"
$compileProcess = Start-Process -FilePath $AhkCompiler -ArgumentList "/in", $entryPath, "/out", $outputExe -Wait -PassThru
$compileExitCode = $compileProcess.ExitCode

if ($compileExitCode -ne 0) {
    Write-Host "Compile failed with exit code $compileExitCode"
    exit $compileExitCode
}

Write-Host "Compile passed."
Write-Host "Output: $outputExe"
exit 0
