# Run LibreOffice Calc with SAL debug log to console (Linux equivalent):
#   SAL_LOG=+INFO.all libreoffice --calc
#
# On Windows soffice.exe is a launcher: it returns to the shell immediately.
# This script runs soffice.bin instead — the process keeps the console until Calc exits.
#
#   .\run_calc_sal_log.ps1
#   .\run_calc_sal_log.ps1 C:\path\to\file.ods
#
# Environment variables:
#   $env:SAL_LOG      — log level (default +INFO.all)
#   $env:SAL_LOG_FILE  — if set, log is also written to a file (otherwise console only)
#   $env:SOFFICE       — path to soffice.exe or soffice.bin

$ErrorActionPreference = 'Stop'

if (-not $env:SAL_LOG) {
    $env:SAL_LOG = '+INFO.all'
}

function Resolve-SofficeBin {
    param([string]$Path)

    if ($Path -match '\.exe$') {
        $binPath = [System.IO.Path]::ChangeExtension($Path, '.bin')
        if (Test-Path -LiteralPath $binPath) {
            return (Resolve-Path -LiteralPath $binPath).Path
        }
    }

    if (Test-Path -LiteralPath $Path) {
        return (Resolve-Path -LiteralPath $Path).Path
    }

    return $null
}

function Find-SofficeBin {
    if ($env:SOFFICE) {
        $resolved = Resolve-SofficeBin $env:SOFFICE
        if ($resolved) {
            return $resolved
        }
        $cmd = Get-Command $env:SOFFICE -ErrorAction SilentlyContinue
        if ($cmd) {
            $resolved = Resolve-SofficeBin $cmd.Source
            if ($resolved) {
                return $resolved
            }
        }
    }

    foreach ($name in @('soffice.bin', 'soffice')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) {
            $resolved = Resolve-SofficeBin $cmd.Source
            if ($resolved) {
                return $resolved
            }
        }
    }

    $exeCandidates = @(
        (Join-Path $env:ProgramFiles 'LibreOffice\program\soffice.exe')
        (Join-Path ${env:ProgramFiles(x86)} 'LibreOffice\program\soffice.exe')
    )
    foreach ($exe in $exeCandidates) {
        if ($exe -and (Test-Path -LiteralPath $exe)) {
            $resolved = Resolve-SofficeBin $exe
            if ($resolved) {
                return $resolved
            }
        }
    }

    throw @"
soffice.bin not found.
Install LibreOffice or set the path:
  `$env:SOFFICE = 'C:\Program Files\LibreOffice\program\soffice.exe'
"@
}

$SofficeBin = Find-SofficeBin
$calcArgs = @('--calc') + $args

Write-Host "SAL_LOG=$($env:SAL_LOG)"
if ($env:SAL_LOG_FILE) {
    Write-Host "SAL_LOG_FILE=$($env:SAL_LOG_FILE)"
} else {
    Write-Host 'SAL_LOG_FILE=(not set, output to console)'
}
Write-Host "Starting: $SofficeBin $($calcArgs -join ' ')"
Write-Host 'Waiting for Calc to exit...'
Write-Host ''

& $SofficeBin @calcArgs

Write-Host ''
Write-Host 'Calc exited.'
