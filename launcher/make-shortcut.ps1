#requires -version 5.1
<#
.SYNOPSIS
    Create the "DeepSeek Harness" shortcut with the official logo and name.
.DESCRIPTION
    Points at deepseek-harness.cmd next to this script, applies the official
    DeepSeek Harness icon, and sets its display name and tooltip.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string] $Destination,
    [string] $LauncherDir,
    [string] $ShortcutName = 'DeepSeek Harness'
)

$ErrorActionPreference = 'Stop'

if (-not $LauncherDir) { $LauncherDir = $PSScriptRoot }
if (-not $LauncherDir) { throw 'Cannot determine the launcher directory; pass -LauncherDir.' }

$launcher = Join-Path $LauncherDir 'deepseek-harness.cmd'
$icon     = Join-Path $LauncherDir 'deepseek-harness.ico'
$workdir  = (Resolve-Path (Join-Path $LauncherDir '..')).Path

foreach ($required in @($launcher, $icon)) {
    if (-not (Test-Path -LiteralPath $required)) {
        throw "Missing required file: $required"
    }
}

if (-not (Test-Path -LiteralPath $Destination)) {
    New-Item -ItemType Directory -Path $Destination -Force | Out-Null
}

$linkPath = Join-Path $Destination "$ShortcutName.lnk"
$shell    = New-Object -ComObject WScript.Shell
$link     = $shell.CreateShortcut($linkPath)

$link.TargetPath        = $launcher
$link.Arguments         = ''
$link.WorkingDirectory  = $workdir
$link.IconLocation      = "$icon,0"
$link.Description       = 'DeepSeek Harness - launch the local web UI (npx @deepseek-ai/dsh web)'
$link.WindowStyle       = 1
$link.Save()

[void][Runtime.InteropServices.Marshal]::ReleaseComObject($shell)
Write-Output "created: $linkPath"
