<#
.SYNOPSIS
Local Outlook PST export (Windows PowerShell 5.1). Includes:
- Bitness-aware sidecar per-item copy when -ResilientCopy is used (matches Outlook 32/64-bit).
- Inline copy fallback per item if sidecar fails.
- Logging + progress every 40 items.
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$CSVFilePath,
    [string]$OutputPath = "\\wilkensnas.wilkensusa.com\shared\Spiceworks\Exchbkup",
    [string]$LogPath = (Join-Path ([System.IO.Path]::GetTempPath()) 'PST_Export_Log.txt'),
    [switch]$UseExistingOutlook,
    [switch]$SkipNetworkCopy,
    [switch]$ResilientCopy,
    [int]$PerItemTimeoutSeconds = 45
)

$script:TranscriptStarted = $false
$script:TranscriptPath = $null

function Ensure-ParentDir ([string]$Path) {
    $dir = Split-Path $Path -Parent
    if ($dir -and -not (Test-Path $dir)) { New-Item -Path $dir -ItemType Directory -Force | Out-Null }
}
function Start-ExportTranscript { param([string]$Path)
    try {
        Ensure-ParentDir -Path $Path
        $script:TranscriptPath = Join-Path (Split-Path $Path -Parent) (
            (Split-Path $Path -LeafBase) + "_transcript" + (Split-Path $Path -Extension)
        )
        Start-Transcript -Path $script:TranscriptPath -Append -IncludeInvocationHeader -ErrorAction Stop | Out-Null
        $script:TranscriptStarted = $true
    } catch {}
}
function Stop-ExportTranscript { if ($script:TranscriptStarted) { try { Stop-Transcript | Out-Null } catch {} } }
function Write-Log { param([string]$msg)
    Ensure-ParentDir -Path $LogPath
    $line = ("[{0}] {1}" -f (Get-Date).ToString("yyyy-MM-dd HH:mm:ss"), $msg)
    try { Add-Content -Path $LogPath -Value $line } catch {}
    Write-Host $msg
}
function Copy-WithRetry { param([Parameter(Mandatory)][string]$From,[Parameter(Mandatory)][string]$To,[int]$Attempts=5)
    for($n=1;$n -le $Attempts;$n++){ try { Copy-Item -LiteralPath $From -Destination $To -Force -ErrorAction Stop; return $true } catch { Start-Sleep -Seconds (5*$n) } }
    return $false
}

$global:OverallId = 1
$global:FolderId  = 2

function Get-TotalItemCount { param($RootFolder)
    $total = 0
    try { $total += $RootFolder.Items.Count } catch {}
    foreach ($sf in @($RootFolder.Folders)) { $total += Get-TotalItemCount -RootFolder $sf }
    return $total
}

function Get-TotalItemCountRobust { 
    param($RootFolder) 
    try { 
        return Get-TotalItemCount -RootFolder $RootFolder 
    } catch { 
        return 0 
    } 
}

# --- Per-folder counts inside PST (for auditing) ---
function Get-PstPerFolderCounts {
    param([object]$RootFolder)
    $lines = New-Object System.Collections.Generic.List[string]
    function _walk([object]$Folder, [string]$Path) {
        $count = 0
        try { $null = $Folder.Items; $count = $Folder.Items.Count } catch {}
        $lines.Add(("{0,6}  {1}" -f $count, ($Path + $Folder.Name)))
        foreach ($sf in @($Folder.Folders)) {
            _walk -Folder $sf -Path ($Path + $Folder.Name + '\')
        }
    }
    _walk -Folder $RootFolder -Path ''
    return $lines
}

# --- Robust item counting helpers (retry & no-skip) ---
function Get-FolderItemCountRobust {
    param([object]$Folder)
    $tries = 0
    while ($tries -lt 3) {
        try {
            $null  = $Folder.Items | Out-Null   # nudge COM
            $count = $Folder.Items.Count
            return [int]$count
        } catch {
            Start-Sleep -Milliseconds (200 * (++$tries))
        }
    }
    return 0
}

# ---- Sidecar helpers (bitness-aware) ----
function Get-OutlookPath {
    try {
        $p = (Get-Process -Name OUTLOOK -ErrorAction SilentlyContinue | Select-Object -First 1).Path
        if ($p) { return $p }
    } catch {}
    $cands = @(
        (Join-Path $env:ProgramFiles        'Microsoft Office\root\Office16\OUTLOOK.EXE'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\root\Office16\OUTLOOK.EXE'),
        (Join-Path $env:ProgramFiles        'Microsoft Office\Office16\OUTLOOK.EXE'),
        (Join-Path ${env:ProgramFiles(x86)} 'Microsoft Office\Office16\OUTLOOK.EXE')
    )
    foreach ($c in $cands) { if (Test-Path $c) { return $c } }
    return $null
}
function Get-PS51ForOutlook {
    $ol = Get-OutlookPath
    $is32 = $false
    if ($ol) { $is32 = ($ol -like '*\Program Files (x86)\*') }
    $ps64 = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $ps32 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
    if ($is32 -and (Test-Path $ps32)) { return $ps32 }
    return $ps64
}
function Invoke-OutlookCopyExternal {
    param(
        [Parameter(Mandatory)][string]$SourceEntryID,
        [Parameter(Mandatory)][string]$DestFolderEntryID,
        [Parameter(Mandatory)][string]$DestStoreID,
        [int]$TimeoutSeconds = 15,
        [ref]$ErrOut
    )
    if ($PSBoundParameters.ContainsKey('ErrOut')) { $ErrOut.Value = '' }

$cmd = @"
try {
  try {
    \$ol = [Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')
  } catch {
    Add-Type -AssemblyName Microsoft.VisualBasic | Out-Null
    \$ol = [Microsoft.VisualBasic.Interaction]::GetObject('', 'Outlook.Application')
  }
  if (-not \$ol) { throw 'No Outlook.Application found in this bitness (check 32/64-bit mismatch).' }
  \$ns=\$ol.GetNamespace('MAPI')
  \$src=\$ns.GetItemFromID('$SourceEntryID')
  if (-not \$src) { throw 'GetItemFromID returned null (source).' }
  \$dest=\$ns.GetFolderFromID('$DestFolderEntryID', '$DestStoreID')
  if (-not \$dest) { throw 'GetFolderFromID returned null (dest).' }
  \$copy = \$src.Copy()
  if (\$copy -eq \$null) { throw 'Item.Copy() returned null.' }
  \$null = \$copy.Move(\$dest)
  [Console]::Out.WriteLine('OK')
  exit 0
} catch {
  [Console]::Error.WriteLine(\$_.Exception.GetType().FullName + ': ' + \$_.Exception.Message)
  exit 10
}
"@

    $psExe  = Get-PS51ForOutlook
    $outF   = [IO.Path]::GetTempFileName()
    $errF   = [IO.Path]::GetTempFileName()
    $args   = @('-NoProfile','-STA','-ExecutionPolicy','Bypass','-Command', $cmd)

    try {
        $p = Start-Process -FilePath $psExe -ArgumentList $args -PassThru -WindowStyle Hidden `
             -RedirectStandardOutput $outF -RedirectStandardError $errF
        if (-not $p.WaitForExit([math]::Max(1,$TimeoutSeconds) * 1000)) {
            try { $p.Kill() } catch {}
            if ($PSBoundParameters.ContainsKey('ErrOut')) { $ErrOut.Value = "Timeout after ${TimeoutSeconds}s" }
            return $false
        }
        if ($p.ExitCode -eq 0) { return $true }

        $err = ''
        try { $err = (Get-Content -Path $errF -Raw -ErrorAction SilentlyContinue) } catch {}
        if ([string]::IsNullOrWhiteSpace($err)) { try { $err = (Get-Content -Path $outF -Raw -ErrorAction SilentlyContinue) } catch {} }
        if ($PSBoundParameters.ContainsKey('ErrOut')) { $ErrOut.Value = $err.Trim() }
        return $false
    } finally {
        Remove-Item $outF,$errF -Force -ErrorAction SilentlyContinue
    }
}

# ---- Export engine ----
function Test-FileUnlocked { param([string]$Path) try { $fs = [System.IO.File]::Open($Path, 'Open', 'Read', 'None'); $fs.Close(); $true } catch { $false } }

function Test-FolderReadOnly {
    param($Folder, [System.Collections.ArrayList]$DebugInfo, [string]$FolderPath)
    
    try {
        # Try to access folder properties that indicate read-only status
        if ($Folder.Items.Count -gt 0) {
            # Test if we can access the first item for modification
            $testItem = $Folder.Items.Item(1)
            if ($testItem) {
                # Try to read a property that would fail on read-only items
                try {
                    $null = $testItem.Subject
                    # If we can read but folder seems like a system folder, check the name
                    $systemFolders = @('United States holidays', 'Holidays', 'Calendar subscriptions')
                    if ($systemFolders -contains $Folder.Name) {
                        [void]$DebugInfo.Add("  Detected system calendar folder: $FolderPath (likely read-only)")
                        return $true
                    }
                } catch {
                    [void]$DebugInfo.Add("  Detected read-only folder: $FolderPath - $($_.Exception.Message)")
                    return $true
                }
            }
        }
        return $false
    } catch {
        [void]$DebugInfo.Add("  Could not determine folder permissions for: $FolderPath - treating as read-only")
        return $true
    }
}

function Copy-CalendarItemDirectly {
    param($Item, $DestFolder, [System.Collections.ArrayList]$DebugInfo, [string]$FolderPath, [int]$ItemIndex)
    
    try {
        # Create a new appointment in the destination folder
        $newItem = $DestFolder.Items.Add(1) # olAppointmentItem = 1
        
        # Copy basic properties first
        if ($Item.Subject) { $newItem.Subject = $Item.Subject }
        if ($Item.Body) { $newItem.Body = $Item.Body }
        if ($Item.Start) { $newItem.Start = $Item.Start }
        if ($Item.End) { $newItem.End = $Item.End }
        if ($Item.Location) { $newItem.Location = $Item.Location }
        if ($Item.Categories) { $newItem.Categories = $Item.Categories }
        if ($Item.Sensitivity) { $newItem.Sensitivity = $Item.Sensitivity }
        if ($Item.BusyStatus) { $newItem.BusyStatus = $Item.BusyStatus }
        if ($Item.AllDayEvent) { $newItem.AllDayEvent = $Item.AllDayEvent }
        if ($Item.ReminderSet) { 
            $newItem.ReminderSet = $Item.ReminderSet 
            if ($Item.ReminderMinutesBeforeStart) {
                $newItem.ReminderMinutesBeforeStart = $Item.ReminderMinutesBeforeStart
            }
        }
        
        # Save the basic item first to ensure it exists
        try {
            $newItem.Save()
        } catch {
            [void]$DebugInfo.Add("    FAILED: Could not save basic calendar item $ItemIndex in ${FolderPath}: $($_.Exception.Message)")
            return $false
        }
        
        # Only try to handle recurrence after basic save succeeds
        if ($Item.IsRecurring) {
            try {
                $srcPattern = $Item.GetRecurrencePattern()
                $destPattern = $newItem.GetRecurrencePattern()
                
                # Copy only safe recurrence properties
                if ($srcPattern.RecurrenceType) { $destPattern.RecurrenceType = $srcPattern.RecurrenceType }
                if ($srcPattern.Interval -and $srcPattern.Interval -gt 0) { $destPattern.Interval = $srcPattern.Interval }
                
                # Validate and copy dates only if they seem reasonable
                $today = Get-Date
                $validStartDate = $srcPattern.PatternStartDate -and $srcPattern.PatternStartDate -gt $today.AddYears(-50) -and $srcPattern.PatternStartDate -lt $today.AddYears(50)
                $validEndDate = $srcPattern.PatternEndDate -and $srcPattern.PatternEndDate -gt $today.AddYears(-50) -and $srcPattern.PatternEndDate -lt $today.AddYears(50)
                
                if ($validStartDate) { $destPattern.PatternStartDate = $srcPattern.PatternStartDate }
                if ($validEndDate -and $srcPattern.PatternEndDate -gt $srcPattern.PatternStartDate) { 
                    $destPattern.PatternEndDate = $srcPattern.PatternEndDate 
                }
                
                if ($srcPattern.Occurrences -and $srcPattern.Occurrences -gt 0 -and $srcPattern.Occurrences -lt 1000) {
                    $destPattern.Occurrences = $srcPattern.Occurrences
                }
                
                # Try to save with recurrence - if this fails, we still have the basic appointment
                try {
                    $newItem.Save()
                } catch {
                    [void]$DebugInfo.Add("    WARNING: Could not save recurrence pattern for item $ItemIndex (basic appointment saved): $($_.Exception.Message)")
                }
                
            } catch {
                [void]$DebugInfo.Add("    WARNING: Could not copy recurrence pattern for item $ItemIndex (basic appointment saved): $($_.Exception.Message)")
            }
        }
        
        return $true
        
    } catch {
        [void]$DebugInfo.Add("    FAILED: Direct calendar copy for item $ItemIndex in ${FolderPath}: $($_.Exception.Message)")
        return $false
    }
}

function Export-FolderRecursively {
    param(
        $SourceFolder, $DestinationParent, [string]$FolderPath,
        [System.Collections.ArrayList]$DebugInfo, [ref]$OverallProcessed, [int]$OverallTotal,
        [switch]$ResilientCopy, [int]$PerItemTimeoutSeconds = 15
    )
    $result = @{ ItemsCopied = 0; FoldersProcessed = 0 }
    if (-not $SourceFolder -or -not $DestinationParent) { return $result }

    $name = $SourceFolder.Name
    $skipNames = @('Sync Issues','Conversation Action Settings','Quick Step Settings','GAL Contacts','United States holidays','Holidays','Calendar subscriptions')
    if ($skipNames -contains $name) { 
        [void]$DebugInfo.Add("  Skipping system folder: $FolderPath"); 
        return $result 
    }
    
    # Test if folder is read-only before attempting any operations
    if (Test-FolderReadOnly -Folder $SourceFolder -DebugInfo $DebugInfo -FolderPath $FolderPath) {
        [void]$DebugInfo.Add("  Skipping read-only folder: $FolderPath")
        return $result
    }

    try { $destFolder = $DestinationParent.Folders.Add($name) } catch { try { $destFolder = $DestinationParent.Folders.Item($name) } catch { $destFolder = $null } }
    $result.FoldersProcessed = 1

    $itemCount = 0; try { $itemCount = $SourceFolder.Items.Count } catch {}
    if ($itemCount -gt 0) {
        [void]$DebugInfo.Add("  Copying $itemCount items from $FolderPath...")
        $sw = [Diagnostics.Stopwatch]::StartNew()
        $lastReported = 0
        $folderSuccessCount = 0

        for ($i = 1; $i -le $itemCount; $i++) {
            try {
                $item = $SourceFolder.Items.Item($i)
                if ($item -ne $null) {
                    $itemCopySuccess = $false
                    if ($ResilientCopy) {
                        $err = ''
                        try {
                            $srcId  = $item.EntryID; $destId = $destFolder.EntryID; $storeId= $destFolder.StoreID
                            $ok = $false
                            if ($srcId -and $destId -and $storeId) {
                                $ok = Invoke-OutlookCopyExternal -SourceEntryID $srcId -DestFolderEntryID $destId -DestStoreID $storeId -TimeoutSeconds $PerItemTimeoutSeconds -ErrOut ([ref]$err)
                            } else { $err = "Missing EntryIDs for sidecar" }
                            if ($ok) { 
                                $result.ItemsCopied++ 
                                $folderSuccessCount++
                                $itemCopySuccess = $true
                            } else {
                                # Sidecar failed, try calendar-aware fallback
                                if ($item.MessageClass -like "*IPM.Appointment*") {
                                    if (Copy-CalendarItemDirectly -Item $item -DestFolder $destFolder -DebugInfo $DebugInfo -FolderPath $FolderPath -ItemIndex $i) {
                                        $result.ItemsCopied++
                                        $folderSuccessCount++
                                        $itemCopySuccess = $true
                                    }
                                } else {
                                    # Standard fallback for non-calendar items
                                    try {
                                        $copy = $item.Copy()
                                        if ($copy -ne $null) { 
                                            try {
                                                $null = $copy.Move($destFolder)
                                                $result.ItemsCopied++
                                                $folderSuccessCount++
                                                $itemCopySuccess = $true
                                            } catch {
                                                try { $copy.Delete() } catch {}
                                                [void]$DebugInfo.Add("    FAILED: Move failed, cleaned up copy for item $i in ${FolderPath}")
                                            }
                                        }
                                    } catch {
                                        [void]$DebugInfo.Add("    FAILED: Standard fallback failed for item $i in ${FolderPath}: $($_.Exception.Message)")
                                    }
                                }
                                
                                if (-not $itemCopySuccess) {
                                    [void]$DebugInfo.Add("    FAILED: Both sidecar and fallback failed for item $i ($($item.Subject)) in ${FolderPath}. Sidecar: $err")
                                }
                            }
                        } catch { 
                            [void]$DebugInfo.Add("    FAILED: Resilient path wrapper error item $i ($($item.Subject)) in ${FolderPath}: $($_.Exception.Message)") 
                        }
                        
                        # Additional verification for critical items
                        if (-not $itemCopySuccess -and ($item.MessageClass -like "*IPM.Appointment*" -or $FolderPath -like "*Calendar*")) {
                            [void]$DebugInfo.Add("    WARNING: Calendar item failed to copy - Subject: '$($item.Subject)', MessageClass: '$($item.MessageClass)'")
                        }
                    } else {
                        # Non-resilient copy path - use calendar-aware method
                        if ($item.MessageClass -like "*IPM.Appointment*") {
                            if (Copy-CalendarItemDirectly -Item $item -DestFolder $destFolder -DebugInfo $DebugInfo -FolderPath $FolderPath -ItemIndex $i) {
                                $result.ItemsCopied++
                                $folderSuccessCount++
                                $itemCopySuccess = $true
                            }
                        } else {
                            # Standard copy for non-calendar items with proper cleanup
                            try {
                                $copy = $item.Copy()
                                if ($copy -ne $null) {
                                    try {
                                        $null = $copy.Move($destFolder)
                                        $result.ItemsCopied++
                                        $folderSuccessCount++
                                        $itemCopySuccess = $true
                                    } catch {
                                        try { $copy.Delete() } catch {}
                                        [void]$DebugInfo.Add("    FAILED: Move failed, cleaned up copy for item $i in ${FolderPath}")
                                    }
                                } else {
                                    [void]$DebugInfo.Add("    FAILED: Copy returned null for item $i in ${FolderPath}")
                                }
                            } catch {
                                [void]$DebugInfo.Add("    FAILED: Copy exception for item $i in ${FolderPath}: $($_.Exception.Message)")
                            }
                        }
                    }
                }
            } catch { [void]$DebugInfo.Add("    Unexpected error reading item $i in ${FolderPath}: $($_.Exception.Message)") }

            if (($i % 40) -eq 0 -or $i -eq $itemCount) {
                $pctFolder = if ($itemCount) { [int](100 * $i / $itemCount) } else { 100 }
                $elapsed = [math]::Max($sw.Elapsed.TotalSeconds,0.001)
                $rate = '{0:n0}/s' -f ($i / $elapsed)
                $eta  = if ($i -gt 0) { [TimeSpan]::FromSeconds([math]::Max((($itemCount-$i) / ($i/$elapsed)),0)).ToString('hh\:mm\:ss') } else { '' }
                $status = "$i of $itemCount  |  $rate  |  ETA $eta"
                Write-Progress -Id $global:FolderId -ParentId $global:OverallId -Activity "Folder: $FolderPath" -Status $status -PercentComplete $pctFolder

                if ($OverallTotal -gt 0) {
                    $delta = $i - $lastReported; if ($delta -lt 0) { $delta = 0 }
                    $lastReported = $i
                    $OverallProcessed.Value += $delta
                    if ($OverallProcessed.Value -gt $OverallTotal) { $OverallProcessed.Value = $OverallTotal }
                    $pctOverall = [int](100 * $OverallProcessed.Value / $OverallTotal)
                    Write-Progress -Id $global:OverallId -Activity "Exporting mailbox" -Status "$($OverallProcessed.Value) of $OverallTotal items" -PercentComplete $pctOverall
                }
            }
        }
        
        # Summary for this folder
        $failureCount = $itemCount - $folderSuccessCount
        if ($failureCount -gt 0) {
            [void]$DebugInfo.Add("  FOLDER SUMMARY: $FolderPath - $folderSuccessCount of $itemCount items copied successfully ($failureCount failed)")
        } else {
            [void]$DebugInfo.Add("  FOLDER SUCCESS: $FolderPath - All $folderSuccessCount items copied successfully")
        }
    } else { [void]$DebugInfo.Add("  Folder $FolderPath is empty, skipping items") }

    $subCount = 0; try { $subCount = $SourceFolder.Folders.Count } catch {}
    if ($subCount -gt 0) {
        foreach ($sf in @($SourceFolder.Folders)) {
            $subPath = "$FolderPath\$($sf.Name)"
            $subRes = Export-FolderRecursively -SourceFolder $sf -DestinationParent $destFolder -FolderPath $subPath -DebugInfo $DebugInfo -OverallProcessed $OverallProcessed -OverallTotal $OverallTotal -ResilientCopy:$ResilientCopy -PerItemTimeoutSeconds $PerItemTimeoutSeconds
            $result.ItemsCopied += [int]$subRes.ItemsCopied
            $result.FoldersProcessed += [int]$subRes.FoldersProcessed
        }
    }
    return $result
}

function Export-OutlookPSTLocal {
    param(
        [string]$OutputFilename, [string]$NetworkOutputPath,
        [switch]$UseExisting, [switch]$RequireExisting,
        [switch]$SkipNetworkCopy, [switch]$ResilientCopy,
        [int]$PerItemTimeoutSeconds = 15
    )

    $results = @{ Success = $false; Message = ""; PSTPath = ""; FileSizeMB = 0; ItemCount = 0; PSTItemCount=0; DebugInfo = New-Object System.Collections.ArrayList }
    $outlook = $null; $namespace = $null; $pstStore = $null; $createdNewOutlook = $false

    try {
        Start-ExportTranscript -Path $LogPath
        Write-Log "Starting local PST export"

        $tempRoot = [System.IO.Path]::GetTempPath()
        $localExportDir = Join-Path $tempRoot 'PST_Export'
        if (!(Test-Path $localExportDir)) { New-Item -Path $localExportDir -ItemType Directory -Force | Out-Null }

        if ([string]::IsNullOrEmpty($OutputFilename)) { $OutputFilename = "MailboxExport_$(Get-Date -Format 'yyyyMMdd_HHmmss').pst" }
        if (-not $OutputFilename.EndsWith('.pst')) { $OutputFilename += '.pst' }

        $localPSTPath = Join-Path $localExportDir $OutputFilename
        [void]$results.DebugInfo.Add("Target PST file: $localPSTPath")

        if ($UseExisting) {
            [void]$results.DebugInfo.Add("Attempting to connect to existing Outlook session...")
            $got = $false
            try { Add-Type -AssemblyName Microsoft.VisualBasic -ErrorAction SilentlyContinue; $outlook = [Microsoft.VisualBasic.Interaction]::GetObject("", "Outlook.Application"); $got = $null -ne $outlook } catch {}
            if (-not $got -or -not $outlook) {
                if ($RequireExisting) { throw "Outlook is not running (RequireExisting). Open Outlook and sign in, then re-run." }
                [void]$results.DebugInfo.Add("No existing session found, creating new Outlook instance...")
                $outlook = New-Object -ComObject Outlook.Application; $createdNewOutlook = $true
            } else { [void]$results.DebugInfo.Add("Connected to existing Outlook session") }
        } else { [void]$results.DebugInfo.Add("Creating new Outlook instance..."); $outlook = New-Object -ComObject Outlook.Application; $createdNewOutlook = $true }

        $namespace = $outlook.GetNamespace("MAPI")
        if ($createdNewOutlook) { $namespace.Logon($null, $null, $false, $true) }

        $inbox = $namespace.GetDefaultFolder(6)
        $primaryStore = $inbox.Parent.Store
        [void]$results.DebugInfo.Add("Selected primary store: " + $primaryStore.DisplayName)

        $root = $primaryStore.GetRootFolder()
        $totalItems = Get-TotalItemCountRobust -RootFolder $root
        Write-Progress -Id $global:OverallId -Activity "Exporting mailbox" -Status "Starting..." -PercentComplete 0
        Write-Log ("Source mailbox total items detected: {0}" -f $totalItems)

        $namespace.AddStoreEx($localPSTPath, 1); Start-Sleep -Milliseconds 500
        for ($i = 1; $i -le $namespace.Stores.Count; $i++) { $s = $namespace.Stores.Item($i); if ($s.FilePath -eq $localPSTPath) { $pstStore = $s; break } }
        if (-not $pstStore) { throw "Failed to create or locate PST store." }

        $overallProcessed = 0; $refOverall = [ref]$overallProcessed
        $mailboxRoot = $primaryStore.GetRootFolder()
        $pstRoot = $pstStore.GetRootFolder()
        $topDest = $pstRoot

        $foldersProcessed = 0
        foreach ($f in @($mailboxRoot.Folders)) {
            $path = "$($primaryStore.DisplayName)\$($f.Name)"
            Write-Progress -Id $global:FolderId -ParentId $global:OverallId -Activity "Folder: $path" -Status "Starting..." -PercentComplete 0
            $r = Export-FolderRecursively -SourceFolder $f -DestinationParent $topDest -FolderPath $path -DebugInfo $results.DebugInfo -OverallProcessed $refOverall -OverallTotal $totalItems -ResilientCopy:$ResilientCopy -PerItemTimeoutSeconds $PerItemTimeoutSeconds
            $results.ItemCount += [int]$r.ItemsCopied
            $foldersProcessed += [int]$r.FoldersProcessed
        }
        [void]$results.DebugInfo.Add(("Export completed: {0} folders processed, {1} total items copied" -f $foldersProcessed, $results.ItemCount))
        Write-Log ("Copied items (counted): {0}" -f $results.ItemCount)

        # Immediate verification before detaching PST
        $pstCountImmediate = Get-TotalItemCountRobust -RootFolder $pstStore.GetRootFolder()
        $results.PSTItemCount = $pstCountImmediate
        Write-Log ("PST verification count prior to detach: {0}" -f $pstCountImmediate)
        
        # Alert on significant discrepancies
        if (($results.ItemCount - $pstCountImmediate) -gt 10) {
            $discrepancy = $results.ItemCount - $pstCountImmediate
            [void]$results.DebugInfo.Add("WARNING: Large discrepancy detected! Claimed copied: $($results.ItemCount), Actually in PST: $pstCountImmediate, Missing: $discrepancy items")
            Write-Log ("WARNING: $discrepancy items appear to have failed silent copy operations")
        }

        # Log per-folder counts in the PST (audit)
        try {
            $pstRootForCounts = $pstStore.GetRootFolder()
            $lines = Get-PstPerFolderCounts -RootFolder $pstRootForCounts
            [void]$results.DebugInfo.Add("PST per-folder counts:")
            foreach ($ln in $lines) { [void]$results.DebugInfo.Add("    " + $ln) }
            try {
                Add-Content -Path $LogPath -Value "PST per-folder counts:"
                Add-Content -Path $LogPath -Value ($lines -join [Environment]::NewLine)
            } catch {}
        } catch {
            Write-Log ("Failed to compute PST per-folder counts: {0}" -f $_.Exception.Message)
        }
        
        try { $namespace.RemoveStore($pstStore.GetRootFolder()) } catch {}
        if ($createdNewOutlook) { try { $outlook.Quit() } catch {}; Start-Sleep -Seconds 2 }

        if (Test-Path $localPSTPath) {
            $waitLimit = [DateTime]::UtcNow.AddSeconds(180)
            while (-not (Test-FileUnlocked -Path $localPSTPath) -and [DateTime]::UtcNow -lt $waitLimit) { Start-Sleep -Milliseconds 600 }
            $results.PSTPath = $localPSTPath
            $results.FileSizeMB = [math]::Round((Get-Item $localPSTPath).Length / 1MB, 1)
            $results.Success = $true
        } else { throw "PST file not found after export." }

        if (-not $SkipNetworkCopy -and $NetworkOutputPath) {
            try {
                $dest = Join-Path $NetworkOutputPath (Split-Path $localPSTPath -Leaf)
                if (Copy-WithRetry -From $localPSTPath -To $dest -Attempts 5) { [void]$results.DebugInfo.Add("Copied PST to network: $dest"); Write-Log "Copied PST to network: $dest" }
                else { throw "Gave up copying PST to network after retries." }
            } catch { [void]$results.DebugInfo.Add("Network copy failed: $($_.Exception.Message)"); Write-Log ("Network copy failed: {0}" -f $_.Exception.Message) }
        }

    } catch { $results.Message = $_.Exception.Message; Write-Log ("ERROR: " + $results.Message) }
    finally {
        Write-Progress -Id $global:FolderId -Activity "Folder" -Completed
        Write-Progress -Id $global:OverallId -Activity "Exporting mailbox" -Completed
        if ($namespace) { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($namespace) | Out-Null } catch {} }
        if ($outlook)   { try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)   | Out-Null } catch {} }
        $namespace = $null; $outlook = $null; $pstStore = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        Stop-ExportTranscript
    }
    return $results
}

# ---- Main ----
try { $csvData = Import-Csv -Path $CSVFilePath -ErrorAction Stop } catch { Write-Error "Failed to read CSV: $($_.Exception.Message)"; exit 1 }
if ($csvData.Count -eq 0) { Write-Error "CSV contains no rows."; exit 1 }

Write-Host "=== Local Outlook PST Export Tool (5.1) ===" -ForegroundColor Yellow
Write-Host "CSV File: $CSVFilePath" -ForegroundColor Gray
Write-Host "Output Path: $OutputPath" -ForegroundColor Gray
Write-Host "Log Path: $LogPath" -ForegroundColor Gray
Write-Host ""

$successCount = 0; $failureCount = 0
foreach ($entry in $csvData) {
    $outputFilename = $entry.OutputFilename
    if ([string]::IsNullOrEmpty($outputFilename)) { $outputFilename = "Export_$(Get-Date -Format 'yyyyMMdd_HHmmss').pst" }
    Write-Host ("--- Processing Export: {0} ---" -f $outputFilename) -ForegroundColor Yellow

    $exportResult = Export-OutlookPSTLocal -OutputFilename $outputFilename -NetworkOutputPath $OutputPath -UseExisting:$UseExistingOutlook -RequireExisting:$UseExistingOutlook -SkipNetworkCopy:$SkipNetworkCopy -ResilientCopy:$ResilientCopy -PerItemTimeoutSeconds $PerItemTimeoutSeconds

    Write-Host "`n--- Debug Information ---" -ForegroundColor Cyan
    foreach ($d in $exportResult.DebugInfo) { Write-Host "  $d" -ForegroundColor Gray }
    Write-Host "--- End Debug ---`n" -ForegroundColor Cyan

    try {
        Add-Content -Path $LogPath -Value ("--- DEBUG ({0}) ---" -f (Get-Date).ToString("s"))
        Add-Content -Path $LogPath -Value ($exportResult.DebugInfo -join [Environment]::NewLine)
        Add-Content -Path $LogPath -Value ("--- END DEBUG ---`r`n")
    } catch {}

    if ($exportResult.Success) {
        Write-Host "[SUCCESS] Export successful!" -ForegroundColor Green
        Write-Host "  File: $($exportResult.PSTPath)" -ForegroundColor Gray
        Write-Host "  Size: $($exportResult.FileSizeMB) MB" -ForegroundColor Gray
        Write-Host "  Items (counted): $($exportResult.ItemCount)" -ForegroundColor Gray
        Write-Host "  Items (PST verify): $($exportResult.PSTItemCount)" -ForegroundColor Gray
        $successCount++
    } else {
        Write-Host "[FAILED] Export failed!" -ForegroundColor Red
        Write-Host "  Error: $($exportResult.Message)" -ForegroundColor Red
        $failureCount++
    }
}

Write-Host "`n=== Export Summary ===" -ForegroundColor Yellow
Write-Host "Successful exports: $successCount" -ForegroundColor Green
Write-Host "Failed exports: $failureCount" -ForegroundColor Red
Write-Host "Log file: $LogPath" -ForegroundColor Cyan