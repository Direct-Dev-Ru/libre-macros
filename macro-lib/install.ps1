# Установка Python-макросов LibreOffice / AlterOffice Calc из каталога macro-lib.
#
# Использование:
#   .\install.ps1                          — скопировать все *.py из этой папки
#   .\install.ps1 collect_workbooks.py     — только указанные файлы
#   .\install.ps1 -Version 3.6.0           — установить с версией 3.6.0
#   .\install.ps1 -Version 3.6.0 collect_workbooks.py
#
# Если файл уже есть — сравнивается SHA-256; при совпадении копирование
# пропускается. Иначе запрашивается подтверждение перезаписи.
# Ответ «a» / «all» / «v» / «в» / «все» — перезаписать все оставшиеся.
#
# Каталоги назначения (весь пакет ставится в каждый):
#   LibreOffice:  $env:LO_MACROS_DIR или %APPDATA%\LibreOffice\4\user\Scripts\python
#   AlterOffice:  $env:AO_MACROS_DIR или %APPDATA%\AlterOffice3\4\user\Scripts\python
#
# Пример для другого каталога:
#   $env:LO_MACROS_DIR = 'D:\custom\Scripts\python'; .\install.ps1

[CmdletBinding()]
param(
    [Alias('v')]
    [string]$Version = '',

    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$Files = @(),

    [switch]$Help
)

$ErrorActionPreference = 'Stop'

$ScriptDir = $PSScriptRoot
$Roaming = [Environment]::GetFolderPath('ApplicationData')
$LoTargetDir = if ($env:LO_MACROS_DIR) {
    $env:LO_MACROS_DIR
} else {
    Join-Path $Roaming 'LibreOffice\4\user\Scripts\python'
}
$AoTargetDir = if ($env:AO_MACROS_DIR) {
    $env:AO_MACROS_DIR
} else {
    Join-Path $Roaming 'AlterOffice3\4\user\Scripts\python'
}

$VersionFile = Join-Path $ScriptDir 'version.txt'
$MacroFile = Join-Path $ScriptDir 'collect_workbooks.py'

$script:OverwriteAll = $false
$script:InstalledNew = 0
$script:Overwritten = 0
$script:Identical = 0
$script:Skipped = 0
$script:Errors = 0

function Show-Usage {
    @'
Установка макросов LibreOffice / AlterOffice из папки macro-lib.

  .\install.ps1 [опции] [файл ...]

Опции:
  -Version, -v ВЕРСИЯ   Задать версию макроса (сохраняет в version.txt
                        и обновляет MACRO_VERSION в collect_workbooks.py)
  -Help, -h             Показать эту справку

Аргументы:
  файл ...              Имена макросов для установки (без .py или с .py)
                        Без аргументов — копируются все *.py

Примеры:
  .\install.ps1                          — все файлы
  .\install.ps1 collect_workbooks        — только collect_workbooks.py
  .\install.ps1 -Version 3.6.0           — все файлы, версия 3.6.0
  .\install.ps1 -v 3.6.0 functions_pp.py

Если макрос уже установлен — сравнивается SHA-256 с копией; при совпадении
файл не трогается. Иначе скрипт спросит, перезаписывать ли файл.
Ответ a / all / v / в / все — перезаписать все оставшиеся без повторных вопросов.
В отчёте отдельно: новые, перезаписанные, идентичные (хэш), пропущенные (отказ).

Переменные окружения:
  LO_MACROS_DIR   каталог Scripts/python LibreOffice
  AO_MACROS_DIR   каталог Scripts/python AlterOffice
                  (по умолчанию %APPDATA%\AlterOffice3\4\user\Scripts\python)

После установки перезапустите Calc / AlterOffice или обновите список макросов.
'@
}

function Normalize-Name([string]$Name) {
    if ($Name -like '*.py') { return $Name }
    return "$Name.py"
}

function Update-VersionInMacro([string]$Ver, [string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        Write-Host "  Предупреждение: файл макроса не найден: $Path" -ForegroundColor Yellow
        return $false
    }
    $content = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
    $pattern = '(?m)^MACRO_VERSION = "[^"]*"'
    $replacement = "MACRO_VERSION = `"$Ver`""
    $newContent = [regex]::Replace($content, $pattern, $replacement, 1)
    if ($newContent -eq $content) {
        Write-Host "  Предупреждение: не удалось обновить версию в $Path" -ForegroundColor Yellow
        return $false
    }
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($Path, $newContent, $utf8NoBom)
    Write-Host "  Версия обновлена: $Ver (в $Path)"
    return $true
}

function Save-Version([string]$Ver) {
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($VersionFile, $Ver, $utf8NoBom)
    Write-Host "  Версия сохранена: $Ver (в version.txt)"
}

function Get-Sha256Hex([string]$Path) {
    return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash.ToLowerInvariant()
}

function Ask-Overwrite([string]$FileLabel) {
    if ($script:OverwriteAll) { return $true }

    while ($true) {
        $reply = Read-Host "  Уже установлен: $FileLabel — перезаписать? [д/а/N]"
        switch -Regex ($reply.Trim()) {
            '^(?i)(a|all|v|в|все)$' {
                $script:OverwriteAll = $true
                return $true
            }
            '^(?i)(y|yes|d|da|д|да)$' {
                return $true
            }
            '^$|^(?i)(n|no|н|нет)$' {
                return $false
            }
            default {
                Write-Host '  Введите д (да), а (все), н (нет); также y / n / all.'
            }
        }
    }
}

# Результат: new | overwritten | identical | skipped | error
function Install-OneFile([string]$Src, [string]$Dest, [string]$Label) {
    if (-not (Test-Path -LiteralPath $Src -PathType Leaf)) {
        Write-Host "  пропуск — не найден в macro-lib: $Label" -ForegroundColor Yellow
        return 'error'
    }

    if (-not (Test-Path -LiteralPath $Dest -PathType Leaf)) {
        $parent = Split-Path -Parent $Dest
        if ($parent -and -not (Test-Path -LiteralPath $parent)) {
            New-Item -ItemType Directory -Force -Path $parent | Out-Null
        }
        Copy-Item -LiteralPath $Src -Destination $Dest -Force
        Write-Host "  установлен (новый): $Label"
        return 'new'
    }

    $srcHash = Get-Sha256Hex $Src
    $destHash = Get-Sha256Hex $Dest
    if ($srcHash -eq $destHash) {
        $short = $srcHash.Substring(0, [Math]::Min(12, $srcHash.Length))
        Write-Host "  пропуск — идентичен (sha256 ${short}…): $Label"
        return 'identical'
    }

    if (Ask-Overwrite $Label) {
        Copy-Item -LiteralPath $Src -Destination $Dest -Force
        Write-Host "  перезаписан: $Label"
        return 'overwritten'
    }

    Write-Host "  оставлен без изменений: $Label"
    return 'skipped'
}

function Accrue-Result([string]$Result) {
    switch ($Result) {
        'new' { $script:InstalledNew++ }
        'overwritten' { $script:Overwritten++ }
        'identical' { $script:Identical++ }
        'skipped' { $script:Skipped++ }
        'error' { $script:Errors++ }
    }
}

function Find-Python {
    foreach ($name in @('python', 'python3', 'py')) {
        $cmd = Get-Command $name -ErrorAction SilentlyContinue
        if ($cmd) {
            if ($name -eq 'py') {
                return @{ Exe = $cmd.Source; Prefix = @('-3') }
            }
            return @{ Exe = $cmd.Source; Prefix = @() }
        }
    }
    return $null
}

function Invoke-Python([string[]]$PyArgs) {
    $py = Find-Python
    if (-not $py) {
        throw 'Ошибка: нужен python, python3 или py (launcher) в PATH'
    }
    $allArgs = @($py.Prefix) + $PyArgs
    & $py.Exe @allArgs
    if ($LASTEXITCODE -ne 0) {
        throw "Команда Python завершилась с кодом $LASTEXITCODE"
    }
}

function Install-PackageToDir([string]$TargetDir, [string]$TargetLabel, [string[]]$FilesToInstall) {
    if (-not (Test-Path -LiteralPath $TargetDir)) {
        New-Item -ItemType Directory -Force -Path $TargetDir | Out-Null
    }

    Write-Host '────────────────────────────────────────'
    Write-Host "Назначение ($TargetLabel): $TargetDir"
    Write-Host ''

    foreach ($file in $FilesToInstall) {
        $src = Join-Path $ScriptDir $file
        $dest = Join-Path $TargetDir $file
        Accrue-Result (Install-OneFile $src $dest $file)
    }

    $ppSrc = Join-Path $ScriptDir 'pythonpath'
    $ppDest = Join-Path $TargetDir 'pythonpath'
    if (Test-Path -LiteralPath $ppSrc -PathType Container) {
        if (-not (Test-Path -LiteralPath $ppDest)) {
            New-Item -ItemType Directory -Force -Path $ppDest | Out-Null
        }
        Get-ChildItem -LiteralPath $ppSrc -Filter '*.py' -File | ForEach-Object {
            $label = "pythonpath/$($_.Name)"
            $dest = Join-Path $ppDest $_.Name
            Accrue-Result (Install-OneFile $_.FullName $dest $label)
        }
    }
    Write-Host ''
}

# --- main ---

if ($Help -or ($Files -contains '-h') -or ($Files -contains '--help') -or ($Files -contains '-Help')) {
    Show-Usage
    exit 0
}

# Отфильтровать служебные флаги, если переданы как позиционные (совместимость с bash-стилем)
$filesToInstall = [System.Collections.Generic.List[string]]::new()
$i = 0
$raw = @($Files)
while ($i -lt $raw.Count) {
    $arg = $raw[$i]
    switch -Regex ($arg) {
        '^(-h|--help)$' {
            Show-Usage
            exit 0
        }
        '^(-v|--version)$' {
            if (($i + 1) -ge $raw.Count -or $raw[$i + 1] -like '-*') {
                Write-Error "Ошибка: после $arg требуется указать версию"
                exit 1
            }
            $Version = $raw[$i + 1]
            $i += 2
            continue
        }
        '^-' {
            Write-Error "Ошибка: неизвестная опция: $arg`nИспользуйте -Help для справки"
            exit 1
        }
        default {
            [void]$filesToInstall.Add((Normalize-Name $arg))
            $i++
        }
    }
}

if ($Version) {
    Write-Host "Установка версии: $Version"
    Save-Version $Version
    if (Test-Path -LiteralPath $MacroFile -PathType Leaf) {
        Update-VersionInMacro $Version $MacroFile | Out-Null
    }
    Write-Host ''
}

if (-not (Test-Path -LiteralPath $ScriptDir -PathType Container)) {
    Write-Error "Ошибка: каталог макросов не найден: $ScriptDir"
    exit 1
}

$flattenScript = Join-Path $ScriptDir 'aoffice_flatten_defs.py'
$compatScript = Join-Path $ScriptDir 'aoffice_compat_check.py'
$py = Find-Python

if ($py -and (Test-Path -LiteralPath $flattenScript -PathType Leaf)) {
    Write-Host 'Совместимость AlterOffice: свёртка многострочных def…'
    Invoke-Python @($flattenScript, '--in-place', '--tree', $ScriptDir)
    Write-Host ''
}

if ($py -and (Test-Path -LiteralPath $compatScript -PathType Leaf)) {
    Write-Host 'Совместимость AlterOffice: проверка фильтра pythonscript…'
    $ppPath = Join-Path $ScriptDir 'pythonpath'
    $prevPp = $env:PYTHONPATH
    if ($prevPp) {
        $env:PYTHONPATH = "$ppPath;$prevPp"
    } else {
        $env:PYTHONPATH = $ppPath
    }
    try {
        Invoke-Python @($compatScript, '--tree', $ScriptDir)
    } finally {
        $env:PYTHONPATH = $prevPp
    }
    Write-Host ''
}

$collectPy = Join-Path $ScriptDir 'collect_workbooks.py'
$genNds = Join-Path $ScriptDir 'generate_no_docstrings.py'
if ((Test-Path -LiteralPath $collectPy -PathType Leaf) -and (Test-Path -LiteralPath $genNds -PathType Leaf)) {
    $needNds = ($filesToInstall.Count -eq 0) -or ($filesToInstall -contains 'collect_workbooks_no_docstrings.py')
    if ($needNds) {
        Write-Host 'Сборка collect_workbooks_no_docstrings.py из collect_workbooks.py…'
        Invoke-Python @($genNds)
        Write-Host ''
    }
}

if ($filesToInstall.Count -eq 0) {
    $skipNames = @(
        'generate_no_docstrings.py',
        'generate_py2_version.py',
        'collect_workbooks_no_docstrings.py'
    )
    Get-ChildItem -LiteralPath $ScriptDir -Filter '*.py' -File | ForEach-Object {
        if ($skipNames -notcontains $_.Name) {
            [void]$filesToInstall.Add($_.Name)
        }
    }
    if ($filesToInstall.Count -eq 0) {
        Write-Error "В $ScriptDir нет файлов *.py для установки."
        exit 1
    }
}

$alwaysWithCollect = @('functions_pp.py', 'functions_final.py')
$needsHelpers = $false
foreach ($f in $filesToInstall) {
    if ($f -in @('collect_workbooks.py', 'collect_workbooks_no_docstrings.py', 'param_wizard.py')) {
        $needsHelpers = $true
        break
    }
}
if ($needsHelpers) {
    foreach ($helper in $alwaysWithCollect) {
        if (($filesToInstall -notcontains $helper) -and (Test-Path -LiteralPath (Join-Path $ScriptDir $helper) -PathType Leaf)) {
            [void]$filesToInstall.Add($helper)
        }
    }
}

$filesArr = @($filesToInstall)

Write-Host "Источник: $ScriptDir"
Write-Host ''

Install-PackageToDir $LoTargetDir 'LibreOffice' $filesArr
# Сброс «перезаписать все» между назначениями не делаем — как в install.sh (один OVERWRITE_ALL на весь прогон)

Install-PackageToDir $AoTargetDir 'AlterOffice' $filesArr

Write-Host 'Отчёт (оба назначения):'
Write-Host "  новых:          $($script:InstalledNew)"
Write-Host "  перезаписано:   $($script:Overwritten)"
Write-Host "  идентичны:      $($script:Identical)"
Write-Host "  без изменений:  $($script:Skipped)"

if ($script:Errors -gt 0) {
    Write-Host "  не найдено:     $($script:Errors)"
    Write-Host ''
    Write-Host 'Готово с ошибками.'
    exit 1
}

Write-Host ''
Write-Host 'Готово.'
