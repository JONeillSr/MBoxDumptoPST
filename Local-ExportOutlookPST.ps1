<#
.SYNOPSIS
Local Outlook PST export with:
- Faster I/O to %TEMP%\PST_Export
- Progress updates every 40 items
- Optional per-row credential launch (-LaunchWithCreds)
- Optional use of an already-signed-in Outlook (-UseExistingOutlook)
- Auto-close Outlook **if the script started it**, to release PST lock
- Accurate item counting in the final summary
- Optional network copy **after** Outlook is closed so the PST isn't locked
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CSVFilePath,
    [string]$OutputPath = "\\wilkensnas.wilkensusa.com\shared\Spiceworks\Exchbkup",
    [string]$LogPath = (Join-Path ([System.IO.Path]::GetTempPath()) 'PST_Export_Log.txt'),
    [switch]$UseExistingOutlook,
    [switch]$LaunchWithCreds,
    [switch]$SkipNetworkCopy
)

# ---------- Helpers ----------

function Split-CredParts {
    param([string]$User)
    if ($User -match '^(?<Domain>[^\\]+)\\(?<Name>.+)$') { return @($Matches.Domain, $Matches.Name) }
    elseif ($User -match '^(?<Name>[^@]+)@(?<Domain>.+)$') { return @('', $User) } # UPN as user, blank domain
    else { return @('', $User) }
}

function Add-GenericCred {
    param([Parameter(Mandatory)][string]$Target,[Parameter(Mandatory)][string]$User,[Parameter(Mandatory)][string]$Password)
    Start-Process -FilePath cmdkey.exe -ArgumentList @("/generic:$Target","/user:$User","/pass:$Password") -WindowStyle Hidden -Wait
}
function Remove-GenericCred {
    param([Parameter(Mandatory)][string]$Target)
    Start-Process -FilePath cmdkey.exe -ArgumentList @("/delete:$Target") -WindowStyle Hidden -Wait
}

function Start-OutlookNetOnly {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] [string]$User,        # domain\user or user@domain
        [Parameter(Mandatory)] [string]$Password,
        [string]$ProfileName = $null
    )

    if (-not ("Win32.Logon" -as [type])) {
Add-Type -TypeDefinition @"
using System;
using System.Runtime.InteropServices;
namespace Win32 {
  public static class Logon {
    [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
    public struct STARTUPINFO {
      public int cb; public string lpReserved, lpDesktop, lpTitle;
      public int dwX, dwY, dwXSize, dwYSize, dwXCountChars, dwYCountChars, dwFillAttribute, dwFlags;
      public short wShowWindow, cbReserved2; public IntPtr lpReserved2, hStdInput, hStdOutput, hStdError;
    }
    [StructLayout(LayoutKind.Sequential)]
    public struct PROCESS_INFORMATION { public IntPtr hProcess, hThread; public int dwProcessId, dwThreadId; }
    [DllImport("advapi32.dll", SetLastError=true, CharSet=CharSet.Unicode)]
    public static extern bool CreateProcessWithLogonW(
      string userName, string domain, string password, int logonFlags,
      string applicationName, string commandLine, int creationFlags,
      IntPtr environment, string currentDirectory,
      ref STARTUPINFO startupInfo, out PROCESS_INFORMATION processInformation);
  }
}
"@ -Language CSharp
    }

    # Find Outlook path robustly
    $outlookPath = $null
    try { $outlookPath = (Get-Command OUTLOOK.EXE -ErrorAction SilentlyContinue).Source } catch {}
    if (-not $outlookPath) {
        $regPaths = @(
            'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE',
            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE'
        )
        foreach ($rp in $regPaths) {
            try {
                if (Test-Path $rp) {
                    $k = Get-Item $rp
                    $p = $k.GetValue('')  # default
                    if (-not $p) { $p = $k.GetValue('Path') }
                    if ($p) {
                        if (Test-Path $p) { $outlookPath = $p }
                        elseif (Test-Path (Join-Path $p 'OUTLOOK.EXE')) { $outlookPath = (Join-Path $p 'OUTLOOK.EXE') }
                    }
                }
            } catch {}
        }
    }
    if (-not $outlookPath) {
        $cands = @(
            (Join-Path $env:ProgramFiles        'Microsoft Office\root\Office16\OUTLOOK.EXE'),
            (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\OUTLOOK.EXE'),
            (Join-Path $env:ProgramFiles        'Microsoft Office\Office16\OUTLOOK.EXE'),
            (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Office16\OUTLOOK.EXE')
        )
        foreach ($c in $cands) { if (Test-Path $c) { $outlookPath = $c; break } }
    }
    if (-not $outlookPath -or -not (Test-Path $outlookPath)) { throw "Could not find OUTLOOK.EXE" }

    $parts  = Split-CredParts $User
    $domain = $parts[0]; $userArg = $parts[1]

    $si = New-Object Win32.Logon+STARTUPINFO
    $si.cb = [Runtime.InteropServices.Marshal]::SizeOf($si)
    $pi = New-Object Win32.Logon+PROCESS_INFORMATION

    $workDir = Split-Path -Path $outlookPath -Parent
    $app = $outlookPath
    $cmd = $null
    if ($ProfileName) { $cmd = "/profile `"$ProfileName`"" }

    $ok = [Win32.Logon]::CreateProcessWithLogonW($userArg, $domain, $Password, 2, $app, $cmd, 0, [IntPtr]::Zero, $workDir, [ref]$si, [ref]$pi)
    if (-not $ok) {
        $err = [ComponentModel.Win32Exception]::new([Runtime.InteropServices.Marshal]::GetLastWin32Error()).Message
        throw "CreateProcessWithLogonW failed: $err"
    }
    return $pi.dwProcessId
}

$global:OverallId = 1
$global:FolderId  = 2

function Get-TotalItemCount {
    param($RootFolder)
    $total = 0
    try { $total += $RootFolder.Items.Count } catch {}
    foreach ($sf in @($RootFolder.Folders)) {
        $total += Get-TotalItemCount -RootFolder $sf
    }
    return $total
}

function Export-FolderRecursively {
    param(
        $SourceFolder,
        $DestinationParent,
        [string]$FolderPath,
        [System.Collections.ArrayList]$DebugInfo,
        [ref]$OverallProcessed,       # ref int
        [int]$OverallTotal            # total items for whole mailbox (from pre-scan)
    )

    $result = @{ ItemsCopied = 0; FoldersProcessed = 0 }
    if (-not $SourceFolder) { return $result }
    if (-not $DestinationParent) { return $result }

    # Skip known system/internal folders by name
    $name = $SourceFolder.Name
    $skipNames = @('Sync Issues','Conversation Action Settings','Quick Step Settings','GAL Contacts')
    if ($skipNames -contains $name) {
        [void]$DebugInfo.Add("  Skipping system folder: $FolderPath")
        return $result
    }

    # Create PST destination folder (or reuse)
    try {
        $destFolder = $DestinationParent.Folders.Add($name)
        [void]$DebugInfo.Add("  Created PST folder: $name")
    } catch {
        try { $destFolder = $DestinationParent.Folders.Item($name) } catch { $destFolder = $null }
    }

    $result.FoldersProcessed = 1

    # Items in this folder
    $itemCount = 0
    try { $itemCount = $SourceFolder.Items.Count } catch {}

    if ($itemCount -gt 0) {
        [void]$DebugInfo.Add("  Copying $itemCount items from $FolderPath...")
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $lastReported = 0
        for ($i = 1; $i -le $itemCount; $i++) {
            try {
                $item = $SourceFolder.Items.Item($i)
                if ($item -ne $null) {
                    $copy = $item.Copy()
                    if ($copy -ne $null) { $null = $copy.Move($destFolder) }
                    $result.ItemsCopied++
                }
            } catch { }

            # Progress update every 40 items
            if (($i % 40) -eq 0 -or $i -eq $itemCount) {
                $pctFolder = if ($itemCount) { [int](100 * $i / $itemCount) } else { 100 }
                $elapsed = [math]::Max($sw.Elapsed.TotalSeconds,0.001)
                $rate = '{0:n0}/s' -f ($i / $elapsed)
                $eta  = if ($i -gt 0) { [TimeSpan]::FromSeconds([math]::Max((($itemCount-$i) / ($i/$elapsed)),0)).ToString('hh\:mm\:ss') } else { '' }
                $status = "$i of $itemCount  |  $rate  |  ETA $eta"
                Write-Progress -Id $global:FolderId -ParentId $global:OverallId -Activity "Folder: $FolderPath" -Status $status -PercentComplete $pctFolder

                # overall progress
                if ($OverallTotal -gt 0) {
                    $delta = $i - $lastReported
                    if ($delta -lt 0) { $delta = 0 }
                    $lastReported = $i
                    $OverallProcessed.Value += $delta
                    if ($OverallProcessed.Value -gt $OverallTotal) { $OverallProcessed.Value = $OverallTotal }
                    $pctOverall = [int](100 * $OverallProcessed.Value / $OverallTotal)
                    Write-Progress -Id $global:OverallId -Activity "Exporting mailbox" -Status "$($OverallProcessed.Value) of $OverallTotal items" -PercentComplete $pctOverall
                }
            }
        }
    } else {
        [void]$DebugInfo.Add("  Folder $FolderPath is empty, skipping items")
    }

    # Recurse subfolders
    $subCount = 0
    try { $subCount = $SourceFolder.Folders.Count } catch {}
    if ($subCount -gt 0) {
        foreach ($sf in @($SourceFolder.Folders)) {
            $subPath = "$FolderPath\$($sf.Name)"
            $subRes = Export-FolderRecursively -SourceFolder $sf -DestinationParent $destFolder -FolderPath $subPath -DebugInfo $DebugInfo -OverallProcessed $OverallProcessed -OverallTotal $OverallTotal
            $result.ItemsCopied += [int]$subRes.ItemsCopied
            $result.FoldersProcessed += [int]$subRes.FoldersProcessed
        }
    }

    return $result
}

function Test-FileUnlocked {
    param([Parameter(Mandatory)][string]$Path)
    try {
        $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None')
        $fs.Close()
        $true
    } catch { $false }
}

function Export-OutlookPSTLocal {
    param(
        [string]$OutputFilename,
        [string]$NetworkOutputPath,
        [switch]$UseExisting,
        [switch]$RequireExisting,
        [switch]$SkipNetworkCopy
    )

    $results = @{
        Success = $false; Message = "";
        PSTPath = ""; FileSizeMB = 0; ItemCount = 0;
        DebugInfo = New-Object System.Collections.ArrayList
    }

    $outlook = $null; $namespace = $null; $pstStore = $null
    $createdNewOutlook = $false

    try {
        [void]$results.DebugInfo.Add("Starting local PST export...")

        # Create local export directory under %TEMP%
        $tempRoot = [System.IO.Path]::GetTempPath()
        $localExportDir = Join-Path $tempRoot 'PST_Export'
        if (!(Test-Path $localExportDir)) {
            New-Item -Path $localExportDir -ItemType Directory -Force | Out-Null
            [void]$results.DebugInfo.Add("Created export directory: $localExportDir")
        }

        if ([string]::IsNullOrEmpty($OutputFilename)) {
            $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
            $OutputFilename = "MailboxExport_$timestamp.pst"
        }
        if (-not $OutputFilename.EndsWith('.pst')) { $OutputFilename += '.pst' }

        $localPSTPath = Join-Path $localExportDir $OutputFilename
        [void]$results.DebugInfo.Add("Target PST file: $localPSTPath")

        # Attach/create Outlook session
        if ($UseExisting) {
            [void]$results.DebugInfo.Add("Attempting to connect to existing Outlook session...")
            $got = $false
            try {
                Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue
                $outlook = [Microsoft.VisualBasic.Interaction]::GetObject("", "Outlook.Application")
                $got = $null -ne $outlook
            } catch { $outlook = $null }

            if (-not $got -or -not $outlook) {
                if ($RequireExisting) {
                    [void]$results.DebugInfo.Add("No existing session found; aborting due to RequireExisting.")
                    throw "Outlook is not running (RequireExisting). Open Outlook and sign in, then re-run."
                } else {
                    [void]$results.DebugInfo.Add("No existing session found, creating new Outlook instance...")
                    $outlook = New-Object -ComObject Outlook.Application
                    $createdNewOutlook = $true
                }
            } else {
                [void]$results.DebugInfo.Add("Connected to existing Outlook session")
            }
        } else {
            [void]$results.DebugInfo.Add("Creating new Outlook instance...")
            $outlook = New-Object -ComObject Outlook.Application
            $createdNewOutlook = $true
        }

        try { (Get-Process OUTLOOK -ErrorAction SilentlyContinue).PriorityClass = 'High' | Out-Null } catch {}

        $namespace = $outlook.GetNamespace("MAPI")
        [void]$results.DebugInfo.Add("MAPI namespace obtained")

        if ($createdNewOutlook) {
            [void]$results.DebugInfo.Add("Attempting MAPI logon...")
            $namespace.Logon($null, $null, $false, $true)
        }

        Start-Sleep -Seconds 1

        # Choose primary mailbox store via Inbox
        $inbox = $namespace.GetDefaultFolder(6)  # olFolderInbox
        $primaryStore = $inbox.Parent.Store
        [void]$results.DebugInfo.Add("Selected primary store: " + $primaryStore.DisplayName)

        # Pre-scan total items for overall progress
        $root = $primaryStore.GetRootFolder()
        $totalItems = Get-TotalItemCount -RootFolder $root
        Write-Progress -Id 1 -Activity "Exporting mailbox" -Status "Starting..." -PercentComplete 0

        # Create Unicode PST
        [void]$results.DebugInfo.Add("Creating PST file (Unicode)...")
        $namespace.AddStoreEx($localPSTPath, 1)   # 1 = olStoreUnicode
        Start-Sleep -Milliseconds 500

        # Find the PST store we just added
        $pstStore = $null
        for ($i = 1; $i -le $namespace.Stores.Count; $i++) {
            $s = $namespace.Stores.Item($i)
            if ($s.FilePath -eq $localPSTPath) { $pstStore = $s; break }
        }
        if (-not $pstStore) { throw "Failed to create or locate PST store." }
        [void]$results.DebugInfo.Add("PST store created and located")

        # Export (sum items accurately)
        $overallProcessed = 0
        $refOverall = [ref]$overallProcessed
        $mailboxRoot = $primaryStore.GetRootFolder()
        $pstRoot = $pstStore.GetRootFolder()
        $topDest = $pstRoot

        $foldersProcessed = 0
        foreach ($f in @($mailboxRoot.Folders)) {
            $path = "$($primaryStore.DisplayName)\$($f.Name)"
            Write-Progress -Id 2 -ParentId 1 -Activity "Folder: $path" -Status "Starting..." -PercentComplete 0
            $r = Export-FolderRecursively -SourceFolder $f -DestinationParent $topDest -FolderPath $path -DebugInfo $results.DebugInfo -OverallProcessed $refOverall -OverallTotal $totalItems
            $results.ItemCount += [int]$r.ItemsCopied
            $foldersProcessed += [int]$r.FoldersProcessed
        }
        [void]$results.DebugInfo.Add(("Export completed: {0} folders processed, {1} total items copied" -f $foldersProcessed, $results.ItemCount))

        # Remove the PST store to free the handle
        try {
            $pstRootFolder = $pstStore.GetRootFolder()
            $namespace.RemoveStore($pstRootFolder)
        } catch {}

        # If we created Outlook, close it to release the PST lock
        if ($createdNewOutlook) {
            try { $outlook.Quit() } catch {}
            Start-Sleep -Seconds 2
        }

        # File info
        if (Test-Path $localPSTPath) {
            # Wait briefly for lock release if needed
            $waitLimit = [DateTime]::UtcNow.AddSeconds(60)
            while (-not (Test-FileUnlocked -Path $localPSTPath) -and [DateTime]::UtcNow -lt $waitLimit) {
                Start-Sleep -Milliseconds 600
            }

            $sizeMB = [math]::Round((Get-Item $localPSTPath).Length / 1MB, 1)
            $results.PSTPath = $localPSTPath
            $results.FileSizeMB = $sizeMB
            $results.Success = $true
            [void]$results.DebugInfo.Add("PST file created successfully: $sizeMB MB")
        } else {
            throw "PST file not found after export."
        }

        # Optional network copy (after Outlook closed if we created it)
        if (-not $SkipNetworkCopy -and $NetworkOutputPath) {
            try {
                $dest = Join-Path $NetworkOutputPath (Split-Path $localPSTPath -Leaf)
                Copy-Item -LiteralPath $localPSTPath -Destination $dest -Force
                [void]$results.DebugInfo.Add("Copied PST to network: $dest")
            } catch {
                [void]$results.DebugInfo.Add("Network copy failed: $($_.Exception.Message)")
            }
        }

    } catch {
        $results.Message = $_.Exception.Message
    } finally {
        # Complete progress bars
        Write-Progress -Id 2 -Activity "Folder" -Completed
        Write-Progress -Id 1 -Activity "Exporting mailbox" -Completed

        # Cleanup COM
        if ($namespace) { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null } catch {} }
        if ($outlook)   { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)   | Out-Null } catch {} }
        $namespace = $null; $outlook = $null; $pstStore = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    return $results
}

function Start-CredPromptWatcher {
    param([string]$LogPath, [int]$TimeoutSeconds = 240)

    $code = @'
param($logPath, $timeout)
Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
$deadline = (Get-Date).AddSeconds($timeout)
$seen = @{}
while ((Get-Date) -lt $deadline) {
    Start-Sleep -Milliseconds 400
    try {
        $root = [System.Windows.Automation.AutomationElement]::RootElement
        $condWin = New-Object System.Windows.Automation.PropertyCondition(
            [System.Windows.Automation.AutomationElement]::ControlTypeProperty,
            [System.Windows.Automation.ControlType]::Window
        )
        $wins = $root.FindAll([System.Windows.Automation.TreeScope]::Children, $condWin)
        for ($i = 0; $i -lt $wins.Count; $i++) {
            $w = $wins.Item($i)
            $buf = @()
            foreach ($ct in @([System.Windows.Automation.ControlType]::Text, [System.Windows.Automation.ControlType]::Edit)) {
                $cond = New-Object System.Windows.Automation.PropertyCondition(
                    [System.Windows.Automation.AutomationElement]::ControlTypeProperty, $ct)
                $elts = $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, $cond)
                for ($j = 0; $j -lt $elts.Count; $j++) { $buf += $elts.Item($j).Current.Name }
            }
            $content = ($buf -join ' ')
            if ([string]::IsNullOrWhiteSpace($content)) { continue }
            $matches = [regex]::Matches($content, '([a-z0-9][a-z0-9\-\.]+\.[a-z]{2,})', 'IgnoreCase')
            foreach ($m in $matches) {
                $host = $m.Groups[1].Value.ToLower()
                if ($host -and -not $seen.ContainsKey($host)) {
                    $seen[$host] = $true
                    $line = ("{0} DETECTED_CRED_PROMPT_HOST {1}" -f (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'), $host)
                    Add-Content -Path $logPath -Value $line
                    try { Write-Host "(watcher) $host" -ForegroundColor Yellow } catch {}
                }
            }
        }
    } catch {}
}
'@

    $ps = [powershell]::Create()
    $ps.Runspace.ApartmentState = 'STA'
    [void]$ps.AddScript($code).AddArgument($LogPath).AddArgument($TimeoutSeconds)
    $null = $ps.BeginInvoke()
    Write-Host "(watcher) Credential prompt logger running for $TimeoutSeconds seconds..." -ForegroundColor DarkGray
}

# ---------- Main ----------

try { $csvData = Import-Csv -Path $CSVFilePath -ErrorAction Stop } catch { Write-Error "Failed to read CSV: $($_.Exception.Message)"; exit 1 }
if ($csvData.Count -eq 0) { Write-Error "CSV contains no rows."; exit 1 }

Write-Host "=== Local Outlook PST Export Tool ===" -ForegroundColor Yellow
Write-Host "CSV File: $CSVFilePath" -ForegroundColor Gray
Write-Host "Output Path: $OutputPath" -ForegroundColor Gray
Write-Host "Log Path: $LogPath`n" -ForegroundColor Gray

$successCount = 0; $failureCount = 0

foreach ($entry in $csvData) {
    $outputFilename = $entry.OutputFilename
    if ([string]::IsNullOrEmpty($outputFilename)) {
        $outputFilename = "Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').pst"
    }

    Write-Host ("--- Processing Export: {0} ---" -f $outputFilename) -ForegroundColor Yellow

    # Per-row: you can still use -LaunchWithCreds; for speed, best is -UseExistingOutlook.
    if ($LaunchWithCreds) {
        # Seed creds if provided
        $targets = @()
        if ($entry.PSObject.Properties.Name -contains 'ExchangeHosts' -and $entry.ExchangeHosts) {
            $targets += ($entry.ExchangeHosts -split '[,\s;|]+' | Where-Object { $_ -and ($_ -match '\.') })
        }
        if (-not $targets -and $entry.PSObject.Properties.Name -contains 'ExchangeHost' -and $entry.ExchangeHost) {
            $targets += $entry.ExchangeHost
        }
        if ($entry.PSObject.Properties.Name -contains 'UPN' -and $entry.UPN) {
            try {
                $smtpDomain = ($entry.UPN -split '@')[1]
                if ($smtpDomain) { $targets += "autodiscover.$smtpDomain","mail.$smtpDomain" }
            } catch {}
        }
        $targets = $targets | Where-Object { $_ } | Select-Object -Unique
        foreach ($t in $targets) {
            if ($entry.ExchangeUsername -and $entry.ExchangePassword) {
                try { Add-GenericCred -Target $t -User $entry.ExchangeUsername -Password $entry.ExchangePassword } catch {}
            }
        }
        try { Start-CredPromptWatcher -LogPath $LogPath -TimeoutSeconds 240 } catch {}

        # Launch Outlook with those creds
        try {
            $null = Start-OutlookNetOnly -User $entry.ExchangeUsername -Password $entry.ExchangePassword -ProfileName $entry.ProfileName
            Start-Sleep -Seconds 5
        } catch {
            Write-Error "Failed to start Outlook with creds: $($_.Exception.Message)"
            continue
        }
    }

    $exportResult = Export-OutlookPSTLocal -OutputFilename $outputFilename -NetworkOutputPath $OutputPath -UseExisting:$UseExistingOutlook -RequireExisting:($UseExistingOutlook -and -not $LaunchWithCreds) -SkipNetworkCopy:$SkipNetworkCopy

    Write-Host "`n--- Debug Information ---" -ForegroundColor Cyan
    foreach ($d in $exportResult.DebugInfo) { Write-Host "  $d" -ForegroundColor Gray }
    Write-Host "--- End Debug ---`n" -ForegroundColor Cyan

    if ($exportResult.Success) {
        Write-Host "✓ Export successful!" -ForegroundColor Green
        Write-Host "  File: $($exportResult.PSTPath)" -ForegroundColor Gray
        Write-Host "  Size: $($exportResult.FileSizeMB) MB" -ForegroundColor Gray
        Write-Host "  Items: $($exportResult.ItemCount)" -ForegroundColor Gray
        $successCount++
    } else {
        Write-Host "✗ Export failed!" -ForegroundColor Red
        Write-Host "  Error: $($exportResult.Message)" -ForegroundColor Red
        $failureCount++
    }

    # Cleanup any seeded creds (if used)
    if ($LaunchWithCreds -and $targets) {
        foreach ($t in $targets) { try { Remove-GenericCred -Target $t } catch {} }
    }
}

Write-Host "`n=== Export Summary ===" -ForegroundColor Yellow
Write-Host "Successful exports: $successCount" -ForegroundColor Green
Write-Host "Failed exports: $failureCount" -ForegroundColor Red
Write-Host "Log file: $LogPath" -ForegroundColor Cyan

if ($successCount -gt 0 -and -not $SkipNetworkCopy) {
    Write-Host "`n=== Network Copy Hint ===" -ForegroundColor Yellow
    Write-Host "To copy PST files to network location later:" -ForegroundColor Cyan
    $defaultExportDir = Join-Path ([System.IO.Path]::GetTempPath()) 'PST_Export'
    Write-Host ("Copy-Item '{0}\*.pst' '{1}' -Force" -f $defaultExportDir, $OutputPath) -ForegroundColor Gray
}
