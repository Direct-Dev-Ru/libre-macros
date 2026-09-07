# Incremental pull of macro-lib into LibreOffice Scripts/python via scp + SHA-256.
#
# 1) Download checksums JSON from the dev machine
# 2) Compare local files; scp only mismatches / missing
# 3) Mirror Scripts/python into AlterOffice3 (same Roaming root, same tail)
#
#   .\pull_macros_scp.ps1
#
# Environment:
#   $env:REMOTE            — SSH host (default su@192.168.89.18)
#   $env:REMOTE_SRC        — remote macro-lib directory
#   $env:REMOTE_CHECKSUMS  — remote JSON with sha256 (default: <repo>/scripts/macro_lib_py_checksums.json)
#   $env:DEST              — LibreOffice Scripts/python
#   $env:DEST_AO           — AlterOffice3 Scripts/python (optional; auto from DEST)

$ErrorActionPreference = 'Stop'

$Remote = if ($env:REMOTE) { $env:REMOTE } else { 'su@192.168.89.18' }
$RemoteSrc = if ($env:REMOTE_SRC) { $env:REMOTE_SRC } else { '/home/su/projects/python/libre-macros/macro-lib' }
$RemoteChecksums = if ($env:REMOTE_CHECKSUMS) {
    $env:REMOTE_CHECKSUMS
} else {
    '/home/su/projects/python/libre-macros/scripts/macro_lib_py_checksums.json'
}

$Dest = if ($env:DEST) {
    $env:DEST
} else {
    'C:\Users\KuznectcovaEN_local\AppData\Roaming\LibreOffice\4\user\Scripts\python'
}

function Get-AlterOfficeDest([string]$LibreDest) {
    if ($env:DEST_AO -and $env:DEST_AO.Trim() -ne '') {
        return $env:DEST_AO.Trim()
    }
    # ...\AppData\Roaming\LibreOffice\4\user\Scripts\python
    # → ...\AppData\Roaming\AlterOffice3\user\Scripts\python
    $m = [regex]::Match($LibreDest, '(?i)^(.*?AppData[/\\]Roaming)[/\\].+$')
    if ($m.Success) {
        return (Join-Path $m.Groups[1].Value 'AlterOffice3\4\user\Scripts\python')
    }
    return ($LibreDest -replace '(?i)LibreOffice[/\\]\d+', 'AlterOffice3')
}

function Test-InstallableRelPath([string]$Rel) {
    $Rel = $Rel.Replace('\', '/')
    if ([string]::IsNullOrWhiteSpace($Rel)) { return $false }
    if ($Rel.StartsWith('py2/')) { return $false }
    if ($Rel -match '(^|/)__pycache__/') { return $false }
    if ($Rel.StartsWith('pythonpath/')) { return $true }
    if ($Rel.Contains('/')) { return $false }
    # root tooling — not for Scripts/python
    if ($Rel -match '^(aoffice_|sync_macro_version\.py$)') { return $false }
    if ($Rel -match '\.(py|txt)$') { return $true }
    return $false
}

function Get-Sha256Hex([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return $null
    }
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Ensure-ParentDir([string]$FilePath) {
    $dir = Split-Path -Parent $FilePath
    if ($dir -and -not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Force -Path $dir | Out-Null
    }
}

function Copy-IfDifferent([string]$Src, [string]$Dst) {
    if (-not (Test-Path -LiteralPath $Src -PathType Leaf)) {
        return $false
    }
    Ensure-ParentDir $Dst
    $h1 = Get-Sha256Hex $Src
    $h2 = Get-Sha256Hex $Dst
    if ($h1 -and $h2 -and ($h1 -eq $h2)) {
        return $false
    }
    Copy-Item -LiteralPath $Src -Destination $Dst -Force
    return $true
}

$DestAo = Get-AlterOfficeDest $Dest
$PythonpathDest = Join-Path $Dest 'pythonpath'
New-Item -ItemType Directory -Force -Path $Dest | Out-Null
New-Item -ItemType Directory -Force -Path $PythonpathDest | Out-Null
New-Item -ItemType Directory -Force -Path (Join-Path $DestAo 'pythonpath') | Out-Null

$TmpJson = Join-Path $env:TEMP ("macro_lib_py_checksums_{0}.json" -f [guid]::NewGuid().ToString('N'))

Write-Host "Source:         ${Remote}:${RemoteSrc}"
Write-Host "Checksums:      ${Remote}:${RemoteChecksums}"
Write-Host "Destination LO: $Dest"
Write-Host "Destination AO: $DestAo"
Write-Host ""

Write-Host "scp checksums JSON..."
scp "${Remote}:${RemoteChecksums}" $TmpJson
if (-not (Test-Path -LiteralPath $TmpJson)) {
    throw "Failed to download checksums: $TmpJson"
}

$manifest = Get-Content -LiteralPath $TmpJson -Raw -Encoding UTF8 | ConvertFrom-Json
if (-not $manifest.files) {
    throw "Checksums JSON has no 'files' array"
}

$ok = 0
$downloaded = 0
$failed = 0
$aoSynced = 0
$entries = @($manifest.files | Where-Object { Test-InstallableRelPath ([string]$_.path) })

Write-Host ("Installable entries: {0} (of {1} in manifest)" -f $entries.Count, $manifest.count)
Write-Host ""

foreach ($entry in $entries) {
    $rel = ([string]$entry.path).Replace('\', '/')
    $want = ([string]$entry.sha256).ToLowerInvariant()
    $local = Join-Path $Dest ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)
    $have = Get-Sha256Hex $local

    if ($have -and $want -and ($have -eq $want)) {
        Write-Host "ok      $rel"
        $ok++
    }
    else {
        $why = if (-not $have) { 'missing' } else { 'mismatch' }
        Write-Host "scp     $rel  ($why)"
        Ensure-ParentDir $local
        try {
            scp "${Remote}:${RemoteSrc}/$rel" $local
            $have2 = Get-Sha256Hex $local
            if ($want -and $have2 -and ($have2 -ne $want)) {
                Write-Host "WARN    $rel — sha256 still differs after scp"
            }
            $downloaded++
        }
        catch {
            Write-Host "FAIL    $rel — $_"
            $failed++
            continue
        }
    }

    $aoLocal = Join-Path $DestAo ($rel -replace '/', [IO.Path]::DirectorySeparatorChar)
    if (Copy-IfDifferent $local $aoLocal) {
        Write-Host "ao      $rel"
        $aoSynced++
    }
}

# Optional: terms_of_use if present remotely but not in py-only manifest
$termsRel = 'pythonpath/terms_of_use_ru.txt'
$termsLocal = Join-Path $Dest ($termsRel -replace '/', [IO.Path]::DirectorySeparatorChar)
if (-not (Test-Path -LiteralPath $termsLocal)) {
    Write-Host "scp     $termsRel  (optional)"
    try {
        Ensure-ParentDir $termsLocal
        scp "${Remote}:${RemoteSrc}/$termsRel" $termsLocal
        $downloaded++
        $aoTerms = Join-Path $DestAo ($termsRel -replace '/', [IO.Path]::DirectorySeparatorChar)
        if (Copy-IfDifferent $termsLocal $aoTerms) { $aoSynced++ }
    }
    catch {
        Write-Host "skip    $termsRel — not on remote or scp failed"
    }
}

Remove-Item -LiteralPath $TmpJson -Force -ErrorAction SilentlyContinue

Write-Host ""
Write-Host ("Done: ok={0} downloaded={1} ao_synced={2} failed={3}" -f $ok, $downloaded, $aoSynced, $failed)
if ($failed -gt 0) { exit 1 }
