<#
    AutoDelete — автоочистка папок day / week / year.

    Без ключей          — интерактивно: мастер установки или меню управления.
    -Run                — тихая чистка (так его запускает планировщик задач).
    -Run -DryRun        — показать, что было бы удалено, ничего не трогая.
    -Install            — поставить автозапуск без вопросов (настройки по умолчанию).
    -Uninstall          — снять автозапуск, сохранив настройки и накопленные сроки.

    Срок жизни объекта:  deadline = max(firstSeen, lastChange) + days
      firstSeen  — когда объект впервые увидели в папке-сроке (state.json);
      lastChange — LastWriteTime файла, либо максимум по всему дереву для папки.
    Любая правка или переименование обнуляют отсчёт.
#>

[CmdletBinding()]
param(
    [switch]$Run,
    [switch]$DryRun,
    [switch]$Install,
    [switch]$Uninstall
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

# --------------------------------------------------------------------------- #
#  Пути и константы
# --------------------------------------------------------------------------- #

$Root = $PSScriptRoot
if (-not $Root) { $Root = Split-Path -Parent $MyInvocation.MyCommand.Path }

$ServiceDir = Join-Path $Root '.autodelete'
$ConfigPath = Join-Path $ServiceDir 'config.json'
$StatePath  = Join-Path $ServiceDir 'state.json'
$LogPath    = Join-Path $ServiceDir 'autodelete.log'
$VbsPath    = Join-Path $ServiceDir 'run-hidden.vbs'

# Windows PowerShell 5.1 — фиксированный путь, есть на любой Windows.
$PsExe = Join-Path $env:WINDIR 'System32\WindowsPowerShell\v1.0\powershell.exe'
if (-not (Test-Path -LiteralPath $PsExe)) { $PsExe = 'powershell.exe' }

# Эти имена никогда не удаляются, даже если окажутся внутри day/week/year.
$ReservedNames = @(
    '.autodelete', 'autodelete.ps1', 'autodelete.cmd', 'autodelete.log',
    'run-hidden.vbs', 'config.json', 'state.json', 'desktop.ini', 'thumbs.db'
)

# --------------------------------------------------------------------------- #
#  Мелкие утилиты
# --------------------------------------------------------------------------- #

function Get-TaskName {
    param([string]$Path)
    $md5   = [System.Security.Cryptography.MD5]::Create()
    $bytes = [System.Text.Encoding]::Unicode.GetBytes($Path.TrimEnd('\').ToLowerInvariant())
    $hash  = $md5.ComputeHash($bytes)
    $md5.Dispose()
    'AutoDelete-' + (($hash[0..3] | ForEach-Object { $_.ToString('x2') }) -join '')
}

$TaskName = Get-TaskName $Root

function Format-Size {
    param([double]$Bytes)
    if ($Bytes -lt 1KB)  { return "$([int]$Bytes) Б" }
    if ($Bytes -lt 1MB)  { return "{0:N1} КБ" -f ($Bytes / 1KB) }
    if ($Bytes -lt 1GB)  { return "{0:N1} МБ" -f ($Bytes / 1MB) }
    return "{0:N2} ГБ" -f ($Bytes / 1GB)
}

function ConvertTo-Utc {
    param([string]$Text)
    try {
        $dt = [datetime]::Parse($Text, [System.Globalization.CultureInfo]::InvariantCulture,
                                [System.Globalization.DateTimeStyles]::RoundtripKind)
        return $dt.ToUniversalTime()
    } catch { return $null }
}

function Initialize-ServiceDir {
    if (-not (Test-Path -LiteralPath $ServiceDir)) {
        New-Item -ItemType Directory -Path $ServiceDir -Force | Out-Null
    }
    try {
        $di = Get-Item -LiteralPath $ServiceDir -Force
        $di.Attributes = $di.Attributes -bor [System.IO.FileAttributes]::Hidden
    } catch { }
}

function Move-BrokenFile {
    param([string]$Path)
    try {
        $stamp = (Get-Date).ToString('yyyyMMdd-HHmmss')
        Move-Item -LiteralPath $Path -Destination "$Path.bad-$stamp" -Force
    } catch { try { Remove-Item -LiteralPath $Path -Force } catch { } }
}

# --------------------------------------------------------------------------- #
#  Конфиг
# --------------------------------------------------------------------------- #

function New-DefaultConfig {
    [ordered]@{
        version        = 1
        buckets        = @(
            [ordered]@{ name = 'day';  days = 1 }
            [ordered]@{ name = 'week'; days = 7 }
            [ordered]@{ name = 'year'; days = 365 }
        )
        log            = $true
        toast          = $true
        recycleBin     = $true
        skipExtensions = @('.crdownload', '.part', '.partial', '.tmp', '.!ut', '.download')
        logMaxBytes    = 1048576
        taskName       = $TaskName
    }
}

function Read-Config {
    if (-not (Test-Path -LiteralPath $ConfigPath)) { return $null }
    try {
        $cfg = Get-Content -LiteralPath $ConfigPath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Move-BrokenFile $ConfigPath
        return $null
    }
    # Минимальная валидация: без внятных buckets работать нельзя.
    if (-not $cfg -or -not $cfg.buckets -or @($cfg.buckets).Count -eq 0) {
        Move-BrokenFile $ConfigPath
        return $null
    }
    foreach ($b in $cfg.buckets) {
        if (-not $b.name -or $null -eq $b.days) { Move-BrokenFile $ConfigPath; return $null }
    }
    if (-not $cfg.skipExtensions) {
        $cfg | Add-Member -NotePropertyName skipExtensions -NotePropertyValue @() -Force
    }
    if (-not $cfg.logMaxBytes) {
        $cfg | Add-Member -NotePropertyName logMaxBytes -NotePropertyValue 1048576 -Force
    }
    return $cfg
}

function Write-Config {
    param($Config)
    Initialize-ServiceDir
    $json = $Config | ConvertTo-Json -Depth 6
    [System.IO.File]::WriteAllText($ConfigPath, $json, (New-Object System.Text.UTF8Encoding $false))
}

# --------------------------------------------------------------------------- #
#  Состояние (firstSeen по каждому объекту)
# --------------------------------------------------------------------------- #

function Read-State {
    $result = @{}
    if (-not (Test-Path -LiteralPath $StatePath)) { return $result }
    try {
        $raw = Get-Content -LiteralPath $StatePath -Raw -Encoding UTF8 | ConvertFrom-Json
    } catch {
        Move-BrokenFile $StatePath
        return $result
    }
    if (-not $raw) { return $result }
    foreach ($p in $raw.PSObject.Properties) {
        $seen = ConvertTo-Utc $p.Value.firstSeen
        if ($null -eq $seen) { continue }
        $entry = @{ firstSeen = $seen }
        if ($null -ne $p.Value.size) { $entry.size = [int64]$p.Value.size }
        $result[$p.Name] = $entry
    }
    return $result
}

function Write-State {
    param([hashtable]$State)
    Initialize-ServiceDir
    $out = [ordered]@{}
    foreach ($key in ($State.Keys | Sort-Object)) {
        $e = [ordered]@{ firstSeen = $State[$key].firstSeen.ToString('o') }
        if ($null -ne $State[$key].size) { $e.size = $State[$key].size }
        $out[$key] = $e
    }
    $json = $out | ConvertTo-Json -Depth 5
    $tmp  = "$StatePath.tmp"
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding $false))
    Move-Item -LiteralPath $tmp -Destination $StatePath -Force
}

# --------------------------------------------------------------------------- #
#  Лог
# --------------------------------------------------------------------------- #

function Write-Log {
    param([string[]]$Lines, $Config)
    if (-not $Config.log -or -not $Lines -or $Lines.Count -eq 0) { return }
    try {
        Initialize-ServiceDir
        $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
        $text  = ($Lines | ForEach-Object { "$stamp  $_" }) -join "`r`n"
        Add-Content -LiteralPath $LogPath -Value $text -Encoding UTF8
        $max = [int64]$Config.logMaxBytes
        if ($max -gt 0 -and (Get-Item -LiteralPath $LogPath).Length -gt $max) {
            $all  = Get-Content -LiteralPath $LogPath -Encoding UTF8
            $keep = $all[[int]($all.Count / 2)..($all.Count - 1)]
            [System.IO.File]::WriteAllLines($LogPath, $keep, (New-Object System.Text.UTF8Encoding $false))
        }
    } catch { }
}

# --------------------------------------------------------------------------- #
#  Уведомление Windows
# --------------------------------------------------------------------------- #

function Show-Toast {
    param([string]$Title, [string]$Text)
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]
        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        $xml = '<toast><visual><binding template="ToastGeneric"><text>' +
               [System.Security.SecurityElement]::Escape($Title) + '</text><text>' +
               [System.Security.SecurityElement]::Escape($Text) + '</text></binding></visual></toast>'
        $doc = New-Object Windows.Data.Xml.Dom.XmlDocument
        $doc.LoadXml($xml)
        $toast = New-Object Windows.UI.Notifications.ToastNotification $doc
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    } catch { }
}

# --------------------------------------------------------------------------- #
#  Удаление в Корзину (SHFileOperation)
# --------------------------------------------------------------------------- #

function Initialize-Native {
    if ('AutoDelete.Native' -as [type]) { return }
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

namespace AutoDelete {
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public struct SHFILEOPSTRUCT {
        public IntPtr hwnd;
        public uint   wFunc;
        public string pFrom;
        public string pTo;
        public ushort fFlags;
        public int    fAnyOperationsAborted;
        public IntPtr hNameMappings;
        public string lpszProgressTitle;
    }

    public static class Native {
        [DllImport("shell32.dll", CharSet = CharSet.Unicode, EntryPoint = "SHFileOperationW")]
        private static extern int SHFileOperation(ref SHFILEOPSTRUCT lpFileOp);

        const uint   FO_DELETE           = 0x0003;
        const ushort FOF_SILENT          = 0x0004;
        const ushort FOF_NOCONFIRMATION  = 0x0010;
        const ushort FOF_ALLOWUNDO       = 0x0040;
        const ushort FOF_NOERRORUI       = 0x0400;

        // Возвращает 0 при успехе, иначе код ошибки SHFileOperation.
        public static int Recycle(string path) {
            SHFILEOPSTRUCT op = new SHFILEOPSTRUCT();
            op.wFunc  = FO_DELETE;
            op.pFrom  = path + "\0\0";
            op.fFlags = (ushort)(FOF_SILENT | FOF_NOCONFIRMATION | FOF_ALLOWUNDO | FOF_NOERRORUI);
            int rc = SHFileOperation(ref op);
            if (rc == 0 && op.fAnyOperationsAborted != 0) { return -1; }
            return rc;
        }
    }
}
'@
}

function Remove-Target {
    param([string]$Path, [bool]$ToRecycleBin)
    if ($ToRecycleBin) {
        Initialize-Native
        $rc = [AutoDelete.Native]::Recycle($Path)
        if ($rc -ne 0) { throw "SHFileOperation вернул код $rc" }
    } else {
        Remove-Item -LiteralPath $Path -Recurse -Force -ErrorAction Stop
    }
}

# --------------------------------------------------------------------------- #
#  Определение времени последнего изменения
# --------------------------------------------------------------------------- #

function Get-ItemLastChangeUtc {
    <#
        Для файла — LastWriteTimeUtc.
        Для папки — максимум LastWriteTimeUtc по всему дереву (файлы и подпапки),
        так что правка вложенного файла обнуляет срок всей родительской папки.
        Обход прерывается, как только найдено что-то новее $StopAtUtc: решение
        от этого не меняется, а «живые» папки проверяются почти мгновенно.
    #>
    param(
        [System.IO.FileSystemInfo]$Item,
        [datetime]$StopAtUtc
    )

    if ($Item -is [System.IO.FileInfo]) { return $Item.LastWriteTimeUtc }

    $max = $Item.LastWriteTimeUtc
    if ($max -gt $StopAtUtc) { return $max }

    $stack = New-Object System.Collections.Stack
    $stack.Push($Item)
    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        try { $entries = $cur.GetFileSystemInfos() } catch { continue }
        foreach ($e in $entries) {
            if ($e.LastWriteTimeUtc -gt $max) {
                $max = $e.LastWriteTimeUtc
                if ($max -gt $StopAtUtc) { return $max }
            }
            if (($e -is [System.IO.DirectoryInfo]) -and
                -not ($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                $stack.Push($e)
            }
        }
    }
    return $max
}

function Test-Excluded {
    param([System.IO.FileSystemInfo]$Item, $Config)
    if ($ReservedNames -contains $Item.Name.ToLowerInvariant()) { return $true }
    if ($Item -is [System.IO.FileInfo]) {
        $ext = $Item.Extension.ToLowerInvariant()
        foreach ($skip in $Config.skipExtensions) {
            if ($ext -eq $skip.ToLowerInvariant()) { return $true }
        }
    }
    return $false
}

# --------------------------------------------------------------------------- #
#  Движок очистки
# --------------------------------------------------------------------------- #

function Invoke-Cleanup {
    param(
        $Config,
        [switch]$Preview,      # ничего не удалять, только показать
        [switch]$Echo          # печатать ход работы в консоль
    )

    $now      = [datetime]::UtcNow
    $state    = Read-State
    $newState = @{}
    $logLines = @()
    $deleted  = @()
    $failed   = @()
    $freed    = [double]0

    foreach ($bucket in $Config.buckets) {
        $bucketPath = Join-Path $Root $bucket.name
        if (-not (Test-Path -LiteralPath $bucketPath)) {
            try { New-Item -ItemType Directory -Path $bucketPath -Force | Out-Null } catch { continue }
        }

        $days      = [double]$bucket.days
        $threshold = $now.AddDays(-$days)   # всё новее этой отметки ещё живёт

        try {
            $items = @(Get-ChildItem -LiteralPath $bucketPath -Force -ErrorAction Stop)
        } catch {
            $logLines += "[$($bucket.name)] ОШИБКА чтения папки: $($_.Exception.Message)"
            continue
        }

        foreach ($item in $items) {
            if (Test-Excluded -Item $item -Config $Config) { continue }

            $key    = "$($bucket.name)\$($item.Name)"
            $isFile = $item -is [System.IO.FileInfo]
            $prev   = $state[$key]

            $firstSeen = $now
            if ($prev) {
                $firstSeen = $prev.firstSeen
                # Файл подменили другим того же имени — отсчёт заново.
                if ($isFile -and $null -ne $prev.size -and $prev.size -ne $item.Length) {
                    $firstSeen = $now
                }
            }

            $lastChange = Get-ItemLastChangeUtc -Item $item -StopAtUtc $threshold
            $effective  = if ($firstSeen -gt $lastChange) { $firstSeen } else { $lastChange }
            $deadline   = $effective.AddDays($days)

            $entry = @{ firstSeen = $firstSeen }
            if ($isFile) { $entry.size = $item.Length }

            if ($now -lt $deadline) {
                $newState[$key] = $entry
                if ($Echo -and $Preview) {
                    $left = $deadline - $now
                    Write-Host ("    остаётся {0,-40} ещё {1:N1} ч" -f $key, $left.TotalHours) -ForegroundColor DarkGray
                }
                continue
            }

            # Срок вышел.
            $size = if ($isFile) { [double]$item.Length } else { [double]0 }

            if ($Preview) {
                $newState[$key] = $entry
                $deleted += $key
                if ($Echo) { Write-Host "    удалил бы  $key" -ForegroundColor Yellow }
                continue
            }

            try {
                Remove-Target -Path $item.FullName -ToRecycleBin ([bool]$Config.recycleBin)
                $deleted += $key
                $freed   += $size
                $sizeText = if ($isFile) { " ($(Format-Size $size))" } else { ' (папка)' }
                $logLines += "[$($bucket.name)] удалено: $($item.Name)$sizeText"
                if ($Echo) { Write-Host "    удалено: $key" -ForegroundColor Green }
            } catch {
                # Скорее всего файл занят — оставляем запись, повторим на следующем запуске.
                $newState[$key] = $entry
                $failed += $key
                $logLines += "[$($bucket.name)] ОШИБКА удаления: $($item.Name) — $($_.Exception.Message)"
                if ($Echo) { Write-Host "    не удалось: $key — $($_.Exception.Message)" -ForegroundColor Red }
            }
        }
    }

    try { Write-State $newState } catch {
        $logLines += "ОШИБКА записи state.json: $($_.Exception.Message)"
    }

    if (-not $Preview) {
        Write-Log -Lines $logLines -Config $Config
        if ($Config.toast -and $deleted.Count -gt 0) {
            $names = ($deleted | Select-Object -First 5) -join ', '
            if ($deleted.Count -gt 5) { $names += " и ещё $($deleted.Count - 5)" }
            $head = "Автоочистка: удалено объектов — $($deleted.Count)"
            if ($freed -gt 0) { $head += ", освобождено $(Format-Size $freed)" }
            Show-Toast -Title $head -Text $names
        }
    }

    [pscustomobject]@{
        Deleted = $deleted
        Failed  = $failed
        Freed   = $freed
        Tracked = $newState.Count
    }
}

# --------------------------------------------------------------------------- #
#  Планировщик задач
# --------------------------------------------------------------------------- #

function Test-TaskInstalled {
    try {
        if (Get-Command Get-ScheduledTask -ErrorAction SilentlyContinue) {
            return $null -ne (Get-ScheduledTask -TaskName $TaskName -ErrorAction SilentlyContinue)
        }
    } catch { }
    & schtasks.exe /Query /TN $TaskName *>$null
    return ($LASTEXITCODE -eq 0)
}

function Get-TaskNextRun {
    # Get-ScheduledTaskInfo для триггеров с повтором часто отдаёт пустое поле,
    # поэтому сначала спрашиваем службу планировщика напрямую.
    try {
        $svc = New-Object -ComObject Schedule.Service
        $svc.Connect()
        $next = $svc.GetFolder('').GetTask($TaskName).NextRunTime
        if ($next -and $next.Year -gt 1900) { return $next }
    } catch { }
    try {
        $info = Get-ScheduledTaskInfo -TaskName $TaskName -ErrorAction SilentlyContinue
        if ($info -and $info.NextRunTime) { return $info.NextRunTime }
    } catch { }
    return $null
}

function Write-LauncherVbs {
    # Тихий запуск без вспышки консоли + самоочистка задачи, если скрипт пропал.
    Initialize-ServiceDir
    $vbs = @"
Option Explicit
Dim fso, sh, root, ps1
Set fso = CreateObject("Scripting.FileSystemObject")
Set sh  = CreateObject("WScript.Shell")
root = fso.GetParentFolderName(fso.GetParentFolderName(WScript.ScriptFullName))
ps1  = fso.BuildPath(root, "AutoDelete.ps1")
If Not fso.FileExists(ps1) Then
    sh.Run "schtasks /Delete /TN ""$TaskName"" /F", 0, True
    WScript.Quit 0
End If
sh.Run """$PsExe"" -NoProfile -ExecutionPolicy Bypass -File """ & ps1 & """ -Run", 0, False
"@
    [System.IO.File]::WriteAllText($VbsPath, $vbs, (New-Object System.Text.UTF8Encoding $false))
}

function New-TaskXml {
    $user = "$env:USERDOMAIN\$env:USERNAME"
    $esc  = { param($s) [System.Security.SecurityElement]::Escape($s) }
    $taskArgs = '//B //Nologo "' + $VbsPath + '"'
    $start = (Get-Date).ToString('yyyy-MM-ddTHH:mm:ss')

    @"
<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Author>$(& $esc $user)</Author>
    <Description>AutoDelete — автоочистка day/week/year в папке $(& $esc $Root)</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
      <UserId>$(& $esc $user)</UserId>
      <Delay>PT1M</Delay>
    </LogonTrigger>
    <TimeTrigger>
      <Repetition>
        <Interval>PT1H</Interval>
        <StopAtDurationEnd>false</StopAtDurationEnd>
      </Repetition>
      <StartBoundary>$start</StartBoundary>
      <Enabled>true</Enabled>
    </TimeTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <UserId>$(& $esc $user)</UserId>
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>LeastPrivilege</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>false</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT1H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>$(& $esc (Join-Path $env:WINDIR 'System32\wscript.exe'))</Command>
      <Arguments>$(& $esc $taskArgs)</Arguments>
      <WorkingDirectory>$(& $esc $Root)</WorkingDirectory>
    </Exec>
  </Actions>
</Task>
"@
}

function Install-Task {
    Write-LauncherVbs
    $xml = New-TaskXml
    try {
        Register-ScheduledTask -TaskName $TaskName -Xml $xml -Force -ErrorAction Stop | Out-Null
        return $true
    } catch {
        # Фолбэк: schtasks ждёт XML в UTF-16.
        try {
            $tmp = Join-Path $env:TEMP "$TaskName.xml"
            [System.IO.File]::WriteAllText($tmp, $xml, [System.Text.Encoding]::Unicode)
            & schtasks.exe /Create /TN $TaskName /XML $tmp /F *>$null
            $ok = ($LASTEXITCODE -eq 0)
            Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
            return $ok
        } catch { return $false }
    }
}

function Uninstall-Task {
    try {
        if (Get-Command Unregister-ScheduledTask -ErrorAction SilentlyContinue) {
            Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false -ErrorAction Stop
            return $true
        }
    } catch { }
    & schtasks.exe /Delete /TN $TaskName /F *>$null
    return ($LASTEXITCODE -eq 0)
}

# --------------------------------------------------------------------------- #
#  Интерактивный ввод
# --------------------------------------------------------------------------- #

function Read-YesNo {
    param([string]$Prompt, [bool]$Default)
    $hint = if ($Default) { 'Д/н' } else { 'д/Н' }
    while ($true) {
        $answer = (Read-Host "  $Prompt ($hint)").Trim().ToLowerInvariant()
        if ($answer -eq '') { return $Default }
        if ($answer -in @('д', 'да', 'y', 'yes', '1', '+')) { return $true }
        if ($answer -in @('н', 'нет', 'n', 'no', '0', '-')) { return $false }
        Write-Host '  Ответьте «д» или «н».' -ForegroundColor DarkYellow
    }
}

function Read-Days {
    param([string]$Prompt, [double]$Default)
    while ($true) {
        $answer = (Read-Host "  $Prompt [$Default]").Trim()
        if ($answer -eq '') { return $Default }
        $value = 0.0
        if ([double]::TryParse($answer.Replace(',', '.'),
                [System.Globalization.NumberStyles]::Float,
                [System.Globalization.CultureInfo]::InvariantCulture, [ref]$value) -and $value -ge 0) {
            return $value
        }
        Write-Host '  Нужно неотрицательное число дней.' -ForegroundColor DarkYellow
    }
}

function Set-ConfigInteractive {
    param($Config)
    Write-Host ''
    $Config.log        = Read-YesNo 'Вести лог-файл удалений?'                    ([bool]$Config.log)
    $Config.toast      = Read-YesNo 'Показывать системные уведомления?'           ([bool]$Config.toast)
    $Config.recycleBin = Read-YesNo 'Удалять в Корзину (иначе безвозвратно)?'     ([bool]$Config.recycleBin)
    return $Config
}

function Set-BucketsInteractive {
    param($Config)
    Write-Host ''
    Write-Host '  Сроки хранения (0 — удалять при первой же проверке):' -ForegroundColor DarkGray
    foreach ($b in $Config.buckets) {
        $b.days = Read-Days "Папка «$($b.name)», дней" ([double]$b.days)
    }
    return $Config
}

# --------------------------------------------------------------------------- #
#  Экраны
# --------------------------------------------------------------------------- #

function Write-Header {
    Write-Host ''
    Write-Host '  ┌────────────────────────────────────────────┐' -ForegroundColor Cyan
    Write-Host '  │            А В Т О О Ч И С Т К А           │' -ForegroundColor Cyan
    Write-Host '  └────────────────────────────────────────────┘' -ForegroundColor Cyan
}

function Invoke-Wizard {
    param($ExistingConfig)

    Write-Header
    $config = if ($ExistingConfig) { $ExistingConfig } else { New-DefaultConfig }
    $fresh  = -not $ExistingConfig

    Write-Host "  Папки будут созданы здесь:  $Root" -ForegroundColor White
    foreach ($b in $config.buckets) {
        Write-Host ("     {0,-5} — удаление через {1} дн." -f $b.name, $b.days) -ForegroundColor Gray
    }
    Write-Host '  Срок обнуляется при любом изменении файла или переименовании.' -ForegroundColor DarkGray
    Write-Host '  Изменение файла внутри вложенной папки обнуляет срок всей папки.' -ForegroundColor DarkGray

    if (-not $fresh) {
        Write-Host ''
        Write-Host '  Настройки уже есть, автозапуск отключён — включаю обратно.' -ForegroundColor Yellow
        Write-Host '  Накопленные сроки сохранены.' -ForegroundColor DarkGray
    } else {
        $config = Set-ConfigInteractive $config
    }

    foreach ($b in $config.buckets) {
        $path = Join-Path $Root $b.name
        if (-not (Test-Path -LiteralPath $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    }
    $config.taskName = $TaskName
    Write-Config $config

    Write-Host ''
    if (Install-Task) {
        Write-Host '  Автозапуск установлен: проверка при входе в систему и раз в час.' -ForegroundColor Green
    } else {
        Write-Host '  Не удалось создать задачу в планировщике.' -ForegroundColor Red
        Write-Host '  Проверьте политику на этом компьютере и запустите скрипт ещё раз.' -ForegroundColor Red
    }

    $result = Invoke-Cleanup -Config $config
    Write-Host "  Взято под наблюдение объектов: $($result.Tracked)" -ForegroundColor Gray
    if ($config.log) { Write-Host "  Лог: $LogPath" -ForegroundColor DarkGray }
    Write-Host ''
    Write-Host '  Готово. Запустите AutoDelete.cmd ещё раз, чтобы открыть меню.' -ForegroundColor White
    Write-Host ''
}

function Show-Status {
    param($Config)

    $installed = Test-TaskInstalled
    $nextRun   = if ($installed) { Get-TaskNextRun } else { $null }

    Write-Header
    Write-Host "  Папка:         $Root"
    if ($installed) {
        $when = if ($nextRun) { $nextRun.ToString('dd.MM.yyyy HH:mm') } else { 'по расписанию' }
        Write-Host "  Статус:        работает, следующая проверка $when" -ForegroundColor Green
    } else {
        Write-Host '  Статус:        остановлена (автозапуска нет)' -ForegroundColor Yellow
    }

    $parts = foreach ($b in $Config.buckets) {
        $path = Join-Path $Root $b.name
        $n = 0
        if (Test-Path -LiteralPath $path) {
            $n = @(Get-ChildItem -LiteralPath $path -Force -ErrorAction SilentlyContinue |
                   Where-Object { -not (Test-Excluded -Item $_ -Config $Config) }).Count
        }
        "$($b.name) $n (срок $($b.days) дн.)"
    }
    Write-Host "  Отслеживается: $($parts -join '  ·  ')"

    $logText   = if ($Config.log)        { 'вкл' } else { 'выкл' }
    $toastText = if ($Config.toast)      { 'вкл' } else { 'выкл' }
    $binText   = if ($Config.recycleBin) { 'в Корзину' } else { 'безвозвратно' }
    Write-Host "  Лог: $logText    Уведомления: $toastText    Удаление: $binText"
    Write-Host ''
    return $installed
}

function Show-Menu {
    param($Config)

    while ($true) {
        $installed = Show-Status $Config

        Write-Host '  1) Изменить настройки'
        Write-Host '  2) Изменить сроки хранения'
        if ($installed) {
            Write-Host '  3) Остановить навсегда (снять автозапуск)'
        } else {
            Write-Host '  3) Включить автозапуск' -ForegroundColor Green
        }
        Write-Host '  4) Запустить очистку сейчас'
        Write-Host '  5) Показать, что будет удалено (без удаления)'
        Write-Host '  6) Показать лог'
        Write-Host '  0) Выход'
        Write-Host ''

        switch ((Read-Host '  Выбор').Trim()) {
            '1' {
                $Config = Set-ConfigInteractive $Config
                Write-Config $Config
                Write-Host '  Сохранено.' -ForegroundColor Green
            }
            '2' {
                $Config = Set-BucketsInteractive $Config
                Write-Config $Config
                Write-Host '  Сохранено.' -ForegroundColor Green
            }
            '3' {
                if ($installed) {
                    if (Uninstall-Task) {
                        Write-Host '  Автозапуск снят. Настройки и накопленные сроки сохранены.' -ForegroundColor Yellow
                        Write-Host '  Чтобы включить обратно — запустите скрипт ещё раз.' -ForegroundColor DarkGray
                    } else {
                        Write-Host '  Не удалось снять задачу.' -ForegroundColor Red
                    }
                } else {
                    if (Install-Task) { Write-Host '  Автозапуск включён.' -ForegroundColor Green }
                    else              { Write-Host '  Не удалось создать задачу.' -ForegroundColor Red }
                }
            }
            '4' {
                Write-Host ''
                $r = Invoke-Cleanup -Config $Config -Echo
                Write-Host ''
                Write-Host "  Удалено: $($r.Deleted.Count)   не удалось: $($r.Failed.Count)   под наблюдением: $($r.Tracked)"
                if ($r.Freed -gt 0) { Write-Host "  Освобождено: $(Format-Size $r.Freed)" }
            }
            '5' {
                Write-Host ''
                $r = Invoke-Cleanup -Config $Config -Preview -Echo
                Write-Host ''
                Write-Host "  Под удаление попадёт: $($r.Deleted.Count)"
            }
            '6' {
                Write-Host ''
                if (Test-Path -LiteralPath $LogPath) {
                    Get-Content -LiteralPath $LogPath -Tail 30 -Encoding UTF8 |
                        ForEach-Object { Write-Host "  $_" -ForegroundColor DarkGray }
                } else {
                    Write-Host '  Лога пока нет.' -ForegroundColor DarkGray
                }
            }
            '0' { return }
            default { Write-Host '  Нет такого пункта.' -ForegroundColor DarkYellow }
        }

        Write-Host ''
        Read-Host '  Enter — вернуться в меню' | Out-Null
        Clear-Host
    }
}

# --------------------------------------------------------------------------- #
#  Точка входа
# --------------------------------------------------------------------------- #

if ($Uninstall) {
    # Неинтерактивно снять автозапуск. Конфиг и накопленные сроки сохраняются.
    if (Uninstall-Task) { Write-Host 'Автозапуск снят.'; exit 0 }
    Write-Host 'Не удалось снять задачу.' -ForegroundColor Red
    exit 1
}

if ($Install) {
    # Неинтерактивно поставить автозапуск с настройками по умолчанию.
    $config = Read-Config
    if (-not $config) { $config = New-DefaultConfig }
    $config.taskName = $TaskName
    foreach ($b in $config.buckets) {
        $path = Join-Path $Root $b.name
        if (-not (Test-Path -LiteralPath $path)) { New-Item -ItemType Directory -Path $path -Force | Out-Null }
    }
    Write-Config $config
    if (Install-Task) {
        Invoke-Cleanup -Config $config | Out-Null
        Write-Host "Автозапуск установлен: $TaskName"
        exit 0
    }
    Write-Host 'Не удалось создать задачу.' -ForegroundColor Red
    exit 1
}

if ($Run) {
    # Фоновый режим. Нет конфига — ничего не удаляем, молча выходим.
    $config = Read-Config
    if (-not $config) { exit 0 }
    try {
        Invoke-Cleanup -Config $config -Preview:$DryRun -Echo:$DryRun | Out-Null
    } catch {
        Write-Log -Lines @("ФАТАЛЬНО: $($_.Exception.Message)") -Config $config
        exit 1
    }
    exit 0
}

$config = Read-Config
if (-not $config) {
    Invoke-Wizard -ExistingConfig $null
} elseif (-not (Test-TaskInstalled)) {
    # Конфиг есть, автозапуска нет — предлагаем включить, сроки не теряем.
    Write-Header
    Write-Host "  Папка:  $Root"
    Write-Host '  Статус: остановлена (автозапуска нет)' -ForegroundColor Yellow
    Write-Host ''
    if (Read-YesNo 'Включить автозапуск обратно?' $true) {
        if (Install-Task) { Write-Host '  Автозапуск включён.' -ForegroundColor Green }
        else              { Write-Host '  Не удалось создать задачу.' -ForegroundColor Red }
    }
    Write-Host ''
    Show-Menu $config
} else {
    Show-Menu $config
}
