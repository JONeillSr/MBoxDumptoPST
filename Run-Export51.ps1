<#
.SYNOPSIS
Shim to run any script in Windows PowerShell 5.1 with pass-through arguments.
#>
param(
  [Parameter(Mandatory=$true)][string]$File,
  [Parameter(ValueFromRemainingArguments=$true)][string[]]$Args
)
$ps51 = "$env:WINDIR\System32\WindowsPowerShell\v1.0\powershell.exe"
& $ps51 -NoProfile -ExecutionPolicy Bypass -File $File @Args
exit $LASTEXITCODE
