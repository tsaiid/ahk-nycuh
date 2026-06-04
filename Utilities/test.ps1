[CmdletBinding()]
param(
    [string]$TestPath,
    [string]$AhkExe = "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

function Resolve-AbsolutePath {
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )

    if ([System.IO.Path]::IsPathRooted($Path)) {
        return [System.IO.Path]::GetFullPath($Path)
    }

    return [System.IO.Path]::GetFullPath((Join-Path $repoRoot $Path))
}

if (-not (Test-Path -LiteralPath $AhkExe)) {
    Write-Error "AutoHotkey executable not found: $AhkExe"
}

if ([string]::IsNullOrWhiteSpace($TestPath)) {
    $testFiles = @(
        Get-ChildItem -LiteralPath (Join-Path $repoRoot "Tests") -Filter "*.test.v2.ahk" -File |
            Sort-Object FullName
    )
} else {
    $resolvedTestPath = Resolve-AbsolutePath $TestPath
    if (-not (Test-Path -LiteralPath $resolvedTestPath -PathType Leaf)) {
        Write-Error "Test file not found: $TestPath"
    }
    $testFiles = @(Get-Item -LiteralPath $resolvedTestPath)
}

if ($testFiles.Count -eq 0) {
    Write-Host "No AHK test files found."
    exit 0
}

$failedCount = 0

foreach ($testFile in $testFiles) {
    Write-Host "Running $($testFile.FullName)"

    $stdoutPath = Join-Path $env:TEMP ("ahk-test-stdout-" + [System.Guid]::NewGuid().ToString("N") + ".log")
    $stderrPath = Join-Path $env:TEMP ("ahk-test-stderr-" + [System.Guid]::NewGuid().ToString("N") + ".log")

    try {
        $testProcess = Start-Process `
            -FilePath $AhkExe `
            -ArgumentList $testFile.FullName `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath `
            -Wait `
            -PassThru `
            -NoNewWindow

        if (Test-Path -LiteralPath $stdoutPath) {
            Get-Content -LiteralPath $stdoutPath | Write-Host
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Get-Content -LiteralPath $stderrPath | Write-Host
        }

        if ($testProcess.ExitCode -ne 0) {
            $failedCount += 1
        }
    } finally {
        if (Test-Path -LiteralPath $stdoutPath) {
            Remove-Item -LiteralPath $stdoutPath -Force
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Remove-Item -LiteralPath $stderrPath -Force
        }
    }
}

if ($failedCount -gt 0) {
    Write-Host "$failedCount test file(s) failed."
    exit 1
}

Write-Host "All AHK test files passed."
exit 0
