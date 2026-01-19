<#
Adds users (by UPN) to an on-prem AD security group.

Requirements:
- Run on a machine with RSAT AD PowerShell module (ActiveDirectory)
- Account running the script must have rights to modify the group
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$GroupIdentity,          # e.g. "SG-MyGroup" or distinguishedName

    [Parameter(Mandatory=$true)]
    [string]$CsvPath,                # e.g. "C:\Temp\users.csv"

    [string]$UpnColumnName = "UPN",

    [switch]$WhatIfOnly,

    [string]$LogPath = ".\AddToGroup.log"
)

function Write-Log {
    param([string]$Message)
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$ts] $Message"
    Add-Content -Path $LogPath -Value $line
    Write-Host $line
}

try {
    Import-Module ActiveDirectory -ErrorAction Stop
} catch {
    throw "ActiveDirectory module not found. Install RSAT: Active Directory module for Windows PowerShell."
}

# Validate inputs
if (-not (Test-Path $CsvPath)) { throw "CSV not found: $CsvPath" }

# Resolve group
try {
    $group = Get-ADGroup -Identity $GroupIdentity -ErrorAction Stop
} catch {
    throw "Could not find group: $GroupIdentity"
}

Write-Log "Target group: $($group.Name) ($($group.DistinguishedName))"
Write-Log "CSV: $CsvPath | Column: $UpnColumnName"
Write-Log "WhatIfOnly: $WhatIfOnly"

# Load users
$rows = Import-Csv -Path $CsvPath
if (-not $rows -or $rows.Count -eq 0) { throw "CSV has no rows." }

# Get existing members to avoid duplicate attempts (fast path)
Write-Log "Reading current group members..."
$existingDns = New-Object 'System.Collections.Generic.HashSet[string]'
Get-ADGroupMember -Identity $group.DistinguishedName -Recursive:$false |
    ForEach-Object { [void]$existingDns.Add($_.DistinguishedName) }

$added = 0
$skippedAlreadyMember = 0
$notFound = 0
$errors = 0

foreach ($r in $rows) {
    $upn = ($r.$UpnColumnName).Trim()
    if ([string]::IsNullOrWhiteSpace($upn)) { continue }

    try {
        $user = Get-ADUser -Filter "UserPrincipalName -eq '$upn'" -Properties UserPrincipalName -ErrorAction Stop

        if ($existingDns.Contains($user.DistinguishedName)) {
            Write-Log "SKIP (already member): $upn"
            $skippedAlreadyMember++
            continue
        }

        if ($WhatIfOnly) {
            Write-Log "WHATIF add: $upn -> $($group.Name)"
        } else {
            Add-ADGroupMember -Identity $group.DistinguishedName -Members $user.DistinguishedName -ErrorAction Stop
            Write-Log "ADDED: $upn"
        }

        $added++
    } catch {
        # If user not found or other error
        $msg = $_.Exception.Message
        if ($msg -match "Cannot find an object") {
            Write-Log "NOT FOUND: $upn"
            $notFound++
        } else {
            Write-Log "ERROR: $upn | $msg"
            $errors++
        }
    }
}

Write-Log "Done."
Write-Log "Summary: Added=$added | AlreadyMember=$skippedAlreadyMember | NotFound=$notFound | Errors=$errors"
