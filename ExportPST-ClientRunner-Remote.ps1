<#
.SYNOPSIS
Client-side runner for exporting Outlook mailboxes to PST (one-time per user).
#>

param(
  [string]$OutputPath = "\\wilkensnas.wilkensusa.com\shared\Spiceworks\Exchbkup",
  [string]$CsvPath    = "C:\Scripts\migrate.csv",
  [string]$ExporterPath = "C:\Scripts\Local-ExportOutlookPST.ps1",
  [switch]$ResilientCopy,
  [int]$PerItemTimeoutSeconds = 30,
  [string]$StatusCsv
)

$ErrorActionPreference = 'Stop'
$ps51 = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"

$markerKey = 'HKCU:\Software\Wilkens\Migration'
$markerName = 'PSTExportDone'
$markerTs   = 'PSTExportTimestamp'
$markerUPN  = 'PSTExportUPN'
$markerFile = 'PSTExportFile'
$markerLog  = 'PSTExportLog'
$markerErr  = 'PSTExportError'

function Ensure-Key { param([string]$Path) if (-not (Test-Path $Path)) { New-Item -Path $Path -Force | Out-Null } }
function Set-Marker { param([int]$Done,[string]$UPN,[string]$PstFile,[string]$LogFile,[string]$Err=''])
  Ensure-Key $markerKey
  New-ItemProperty -Path $markerKey -Name $markerName -Value $Done -PropertyType DWord -Force | Out-Null
  New-ItemProperty -Path $markerKey -Name $markerTs   -Value (Get-Date).ToString('s') -PropertyType String -Force | Out-Null
  if ($UPN)     { New-ItemProperty -Path $markerKey -Name $markerUPN  -Value $UPN     -PropertyType String -Force | Out-Null }
  if ($PstFile) { New-ItemProperty -Path $markerKey -Name $markerFile -Value $PstFile -PropertyType String -Force | Out-Null }
  if ($LogFile) { New-ItemProperty -Path $markerKey -Name $markerLog  -Value $LogFile -PropertyType String -Force | Out-Null }
  if ($Err)     { New-ItemProperty -Path $markerKey -Name $markerErr  -Value $Err     -PropertyType String -Force | Out-Null }
}

function Append-StatusCsv {
  param([string]$Csv,[string]$Computer,[string]$User,[string]$Result,[string]$Pst,[decimal]$SizeMB,[int]$Items,[int]$PstVerify,[string]$Log,[string]$Error)
  if (-not $Csv) { return }
  try {
    $hdr = 'Timestamp,Computer,User,Result,PST,SizeMB,Items,PSTVerify,Log,Error'
    if (-not (Test-Path $Csv)) { New-Item -ItemType File -Path $Csv -Force | Out-Null; Add-Content -Path $Csv -Value $hdr }
    $line = ('"{0}","{1}","{2}",{3},"{4}",{5},{6},{7},"{8}","{9}"' -f (Get-Date).ToString('s'), $Computer, $User, $Result, $Pst, $SizeMB, $Items, $PstVerify, $Log, ($Error -replace '"',''''))
    Add-Content -Path $Csv -Value $line
  } catch { }
}

function Get-UPNFromCsvOrEnv {
  try {
    if (Test-Path $CsvPath) { $row = Import-Csv -Path $CsvPath | Select-Object -First 1; if ($row -and $row.UPN) { return [string]$row.UPN } }
  } catch { }
  if ($env:USERDNSDOMAIN) { return ("{0}@{1}" -f $env:USERNAME, $env:USERDNSDOMAIN.ToLower()) }
  return $env:USERNAME
}
function Ensure-OutlookRunning {
  if (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
    try { Start-Process OUTLOOK.EXE | Out-Null } catch { }
    $deadline = (Get-Date).AddSeconds(60)
    while (-not (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue)) {
      if ((Get-Date) -gt $deadline) { break }
      Start-Sleep -Milliseconds 500
    }
  }
}

if (Test-Path $markerKey) {
  try {
    $v = (Get-ItemProperty -Path $markerKey -Name $markerName -ErrorAction SilentlyContinue).$markerName
    if ($null -ne $v -and [int]$v -eq 1) { Write-Host 'PST export already marked complete for this user. Exiting.' -ForegroundColor Yellow; return }
  } catch { }
}

if (-not (Test-Path $ExporterPath)) { throw "Exporter not found: $ExporterPath" }
if (-not (Test-Path $CsvPath))      { throw "CSV not found: $CsvPath" }

$logRoot = Join-Path $env:TEMP 'PST_Export'
if (-not (Test-Path $logRoot)) { New-Item -ItemType Directory -Path $logRoot -Force | Out-Null }
$logPath = Join-Path $logRoot 'PST_Export_Log.txt'

$upn = Get-UPNFromCsvOrEnv
Ensure-OutlookRunning

$argList = @('-NoProfile','-ExecutionPolicy','Bypass','-File', $ExporterPath,
  '-CSVFilePath', $CsvPath,
  '-UseExistingOutlook',
  '-PerItemTimeoutSeconds', 30,
  '-LogPath', $logPath,
  '-OutputPath', $OutputPath
)
if ($ResilientCopy) { $argList += '-ResilientCopy' }

$psi = New-Object System.Diagnostics.ProcessStartInfo
$psi.FileName = $ps51
$psi.Arguments = ($argList -join ' ')
$psi.RedirectStandardOutput = $true
$psi.RedirectStandardError  = $true
$psi.UseShellExecute = $false
$psi.CreateNoWindow = $true
$proc = [System.Diagnostics.Process]::Start($psi)
$proc.WaitForExit()

$success = $false; $pstFile = $null; $sizeMB = 0; $items = 0; $pstVerify = 0; $errMsg = ''
try {
  if (Test-Path $logPath) {
    $log = Get-Content -Path $logPath -Raw
    if ($log -match '✓ Export successful!') { $success = $true }
    if ($log -match 'File:\s*(.+)')          { $pstFile = $Matches[1].Trim() }
    if ($log -match 'Size:\s*([0-9\.]+) MB'){ $sizeMB  = [decimal]$Matches[1] }
    if ($log -match 'Items \(counted\):\s*(\d+)')    { $items = [int]$Matches[1] }
    if ($log -match 'Items \(PST verify\):\s*(\d+)') { $pstVerify = [int]$Matches[1] }
    if (-not $success) {
      if ($log -match 'ERROR:\s*(.+)') { $errMsg = $Matches[1].Trim() }
      elseif ($log -match '✗ Export failed!\s*\r?\n\s*Error:\s*(.+)') { $errMsg = $Matches[1].Trim() }
      else { $errMsg = 'Unknown failure. See log.' }
    }
  } else { $errMsg = 'Exporter did not write the log file' }
} catch { $errMsg = $_.Exception.Message }

function Set-MarkerWrap([int]$ok,[string]$err='') {
  Ensure-Key $markerKey
  New-ItemProperty -Path $markerKey -Name $markerName -Value $ok -PropertyType DWord -Force | Out-Null
  New-ItemProperty -Path $markerKey -Name $markerTs   -Value (Get-Date).ToString('s') -PropertyType String -Force | Out-Null
  New-ItemProperty -Path $markerKey -Name $markerUPN  -Value $upn -PropertyType String -Force | Out-Null
  if ($pstFile) { New-ItemProperty -Path $markerKey -Name $markerFile -Value $pstFile -PropertyType String -Force | Out-Null }
  New-ItemProperty -Path $markerKey -Name $markerLog  -Value $logPath -PropertyType String -Force | Out-Null
  if ($err)     { New-ItemProperty -Path $markerKey -Name $markerErr  -Value $err -PropertyType String -Force | Out-Null }
}

if ($success) {
  Set-MarkerWrap 1
  if ($StatusCsv) { Append-StatusCsv -Csv $StatusCsv -Computer $env:COMPUTERNAME -User $upn -Result 'Success' -Pst $pstFile -SizeMB $sizeMB -Items $items -PstVerify $pstVerify -Log $logPath -Error '' }
  Write-Host "PST export complete for $upn" -ForegroundColor Green
} else {
  Set-MarkerWrap 0 $errMsg
  if ($StatusCsv) { Append-StatusCsv -Csv $StatusCsv -Computer $env:COMPUTERNAME -User $upn -Result 'Failed' -Pst $pstFile -SizeMB $sizeMB -Items $items -PstVerify $pstVerify -Log $logPath -Error $errMsg }
  Write-Host "PST export FAILED for $upn: $errMsg" -ForegroundColor Red
  exit 1
}
