[CmdletBinding()]
param(
    [string]$TestPath,
    [string]$AhkExe = "",
    [int]$TimeoutSeconds = 30
)

$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $repoRoot

function Resolve-DefaultAhkExe {
    $candidates = [System.Collections.Generic.List[string]]::new()
    if ($env:SCOOP) {
        $candidates.Add((Join-Path $env:SCOOP "apps\autohotkey\current\v2\AutoHotkey64.exe"))
        $candidates.Add((Join-Path $env:SCOOP "apps\autohotkey\current\AutoHotkey64.exe"))
    }
    if ($env:USERPROFILE) {
        $candidates.Add((Join-Path $env:USERPROFILE "scoop\apps\autohotkey\current\v2\AutoHotkey64.exe"))
        $candidates.Add((Join-Path $env:USERPROFILE "scoop\apps\autohotkey\current\AutoHotkey64.exe"))
    }
    if ($env:SCOOP_GLOBAL) {
        $candidates.Add((Join-Path $env:SCOOP_GLOBAL "apps\autohotkey\current\v2\AutoHotkey64.exe"))
        $candidates.Add((Join-Path $env:SCOOP_GLOBAL "apps\autohotkey\current\AutoHotkey64.exe"))
    }
    $candidates.Add("C:\ProgramData\scoop\apps\autohotkey\current\v2\AutoHotkey64.exe")
    $candidates.Add("C:\ProgramData\scoop\apps\autohotkey\current\AutoHotkey64.exe")
    $candidates.Add("C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe")

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return "C:\Program Files\AutoHotkey\v2\AutoHotkey64.exe"
}

if ([string]::IsNullOrWhiteSpace($AhkExe)) {
    $AhkExe = Resolve-DefaultAhkExe
}

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

Write-Host "Using AutoHotkey: $AhkExe"

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

    $wrapperPath = Join-Path $env:TEMP ("ahk-test-wrapper-" + [System.Guid]::NewGuid().ToString("N") + ".ahk")
    $stdoutPath = Join-Path $env:TEMP ("ahk-test-stdout-" + [System.Guid]::NewGuid().ToString("N") + ".log")
    $stderrPath = Join-Path $env:TEMP ("ahk-test-stderr-" + [System.Guid]::NewGuid().ToString("N") + ".log")

    try {
        $escapedTestPath = $testFile.FullName.Replace('"', '""')
        $wrapperContent = @(
            "#Requires AutoHotkey v2.0"
            "#Warn All, StdOut"
            "#ErrorStdOut"
            ('#Include "{0}"' -f $escapedTestPath)
        )
        Set-Content -LiteralPath $wrapperPath -Value $wrapperContent -Encoding ascii

        $testProcess = Start-Process `
            -FilePath $AhkExe `
            -ArgumentList $wrapperPath `
            -RedirectStandardOutput $stdoutPath `
            -RedirectStandardError $stderrPath `
            -PassThru `
            -NoNewWindow

        $null = $testProcess.Handle
        if (-not $testProcess.WaitForExit($TimeoutSeconds * 1000)) {
            $testProcess.Kill()
            $testProcess.WaitForExit()
            Write-Host "FAIL Test timed out after $TimeoutSeconds seconds."
            $failedCount += 1
            continue
        }
        $testProcess.WaitForExit()
        $testProcess.Refresh()

        if (Test-Path -LiteralPath $stdoutPath) {
            Get-Content -LiteralPath $stdoutPath | Write-Host
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Get-Content -LiteralPath $stderrPath | Write-Host
        }

        if ($testProcess.ExitCode -ne 0) {
            Write-Host "FAIL AutoHotkey exited with code $($testProcess.ExitCode)."
            $failedCount += 1
        }
    } finally {
        if (Test-Path -LiteralPath $stdoutPath) {
            Remove-Item -LiteralPath $stdoutPath -Force
        }

        if (Test-Path -LiteralPath $stderrPath) {
            Remove-Item -LiteralPath $stderrPath -Force
        }

        if (Test-Path -LiteralPath $wrapperPath) {
            Remove-Item -LiteralPath $wrapperPath -Force
        }
    }
}

if ($failedCount -gt 0) {
    Write-Host "$failedCount test file(s) failed."
    exit 1
}

Write-Host "All AHK test files passed."
exit 0
