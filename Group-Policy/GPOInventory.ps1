<#
.SYNOPSIS
    Exports a full Group Policy inventory to CSV files.

.DESCRIPTION
    Collects:
    - GPO summary
    - GPO links and link status
    - Security filtering / delegation entries
    - WMI filter assignment status

    Outputs:
    - GPO_Summary_<timestamp>.csv
    - GPO_Links_<timestamp>.csv
    - GPO_SecurityFiltering_<timestamp>.csv
    - GPO_Errors_<timestamp>.csv

.NOTES
    Requires:
    - GroupPolicy module
    - Permission to read GPOs
    - PowerShell 5.1 compatible

    Default output folder:
    C:\Users\<current-user>\Documents\GPOInventory
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = (Join-Path $env:USERPROFILE 'Documents\GPOInventory')
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Get-GpoStatusText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [int]$GpoStatus
    )

    switch ($GpoStatus) {
        0 { 'All settings enabled' }
        1 { 'Computer settings disabled' }
        2 { 'User settings disabled' }
        3 { 'All settings disabled' }
        default { "Unknown ($GpoStatus)" }
    }
}

function Convert-ToBoolean {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$Value
    )

    if ($null -eq $Value) {
        return $false
    }

    $text = [string]$Value
    return ($text -eq 'true' -or $text -eq 'True')
}

function Get-XmlNodeInnerText {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$XPath
    )

    try {
        $node = $Xml.SelectSingleNode($XPath)
        if ($null -ne $node -and -not [string]::IsNullOrWhiteSpace($node.InnerText)) {
            return $node.InnerText.Trim()
        }

        return $null
    }
    catch {
        return $null
    }
}

function Get-GpoLinksFromXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$GpoName,

        [Parameter(Mandatory = $true)]
        [string]$GpoId
    )

    $results = @()

    try {
        $somNodes = @($Xml.SelectNodes('//GPO/LinksTo/SOM'))

        foreach ($som in $somNodes) {
            if ($null -eq $som) {
                continue
            }

            $somName = $null
            $somPath = $null
            $enabled = $false
            $enforced = $false
            $linkOrder = $null

            if ($null -ne $som.SOMName) {
                $somName = [string]$som.SOMName
            }

            if ($null -ne $som.SOMPath) {
                $somPath = [string]$som.SOMPath
            }

            if ($null -ne $som.Enabled) {
                $enabled = Convert-ToBoolean -Value $som.Enabled
            }

            if ($null -ne $som.NoOverride) {
                $enforced = Convert-ToBoolean -Value $som.NoOverride
            }

            if ($null -ne $som.Order) {
                $linkOrder = [string]$som.Order
            }

            $results += [pscustomobject]@{
                GPOName     = $GpoName
                GPOId       = $GpoId
                SOMName     = $somName
                SOMPath     = $somPath
                LinkEnabled = $enabled
                Enforced    = $enforced
                LinkOrder   = $linkOrder
            }
        }
    }
    catch {
        # Return what we have; caller handles errors separately.
    }

    return $results
}

function Get-GpoSecurityFilteringFromXml {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [xml]$Xml,

        [Parameter(Mandatory = $true)]
        [string]$GpoName,

        [Parameter(Mandatory = $true)]
        [string]$GpoId
    )

    $results = @()

    try {
        $permissionNodes = @($Xml.SelectNodes('//GPO/SecurityDescriptor/Permissions/TrusteePermissions'))

        foreach ($perm in $permissionNodes) {
            if ($null -eq $perm) {
                continue
            }

            $trusteeName = $null
            $trusteeSid = $null
            $permissionType = $null
            $inherited = $false
            $deny = $false
            $readPermission = $false
            $applyGroupPolicy = $false

            try {
                if ($null -ne $perm.Trustee -and $null -ne $perm.Trustee.Name) {
                    if ($perm.Trustee.Name -is [string]) {
                        $trusteeName = [string]$perm.Trustee.Name
                    }
                    elseif ($null -ne $perm.Trustee.Name.'#text') {
                        $trusteeName = [string]$perm.Trustee.Name.'#text'
                    }
                    else {
                        $trusteeName = [string]$perm.Trustee.Name
                    }
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Trustee -and $null -ne $perm.Trustee.SID) {
                    if ($perm.Trustee.SID -is [string]) {
                        $trusteeSid = [string]$perm.Trustee.SID
                    }
                    elseif ($null -ne $perm.Trustee.SID.'#text') {
                        $trusteeSid = [string]$perm.Trustee.SID.'#text'
                    }
                    else {
                        $trusteeSid = [string]$perm.Trustee.SID
                    }
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Type -and $null -ne $perm.Type.PermissionType) {
                    $permissionType = [string]$perm.Type.PermissionType
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Inherited) {
                    $inherited = Convert-ToBoolean -Value $perm.Inherited
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Standard -and $null -ne $perm.Standard.Deny) {
                    $deny = Convert-ToBoolean -Value $perm.Standard.Deny
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Standard -and $null -ne $perm.Standard.Read) {
                    $readPermission = Convert-ToBoolean -Value $perm.Standard.Read
                }
            }
            catch { }

            try {
                if ($null -ne $perm.Standard -and $null -ne $perm.Standard.ApplyGroupPolicy) {
                    $applyGroupPolicy = Convert-ToBoolean -Value $perm.Standard.ApplyGroupPolicy
                }
            }
            catch { }

            $results += [pscustomobject]@{
                GPOName          = $GpoName
                GPOId            = $GpoId
                TrusteeName      = $trusteeName
                TrusteeSID       = $trusteeSid
                PermissionType   = $permissionType
                Inherited        = $inherited
                Deny             = $deny
                ReadPermission   = $readPermission
                ApplyGroupPolicy = $applyGroupPolicy
                IsSecurityFilter = ($readPermission -and $applyGroupPolicy -and -not $deny)
            }
        }
    }
    catch {
        # Return what we have; caller handles errors separately.
    }

    return $results
}

try {
    Write-Host "Starting GPO inventory export..." -ForegroundColor Cyan

    if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
        throw "The GroupPolicy module is not installed or not available. Install RSAT Group Policy Management tools first."
    }

    Import-Module GroupPolicy -ErrorAction Stop

    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    $summaryPath = Join-Path -Path $OutputFolder -ChildPath ("GPO_Summary_{0}.csv" -f $timestamp)
    $linksPath = Join-Path -Path $OutputFolder -ChildPath ("GPO_Links_{0}.csv" -f $timestamp)
    $securityPath = Join-Path -Path $OutputFolder -ChildPath ("GPO_SecurityFiltering_{0}.csv" -f $timestamp)
    $errorsPath = Join-Path -Path $OutputFolder -ChildPath ("GPO_Errors_{0}.csv" -f $timestamp)

    Write-Host ("Output folder: {0}" -f $OutputFolder) -ForegroundColor Yellow
    Write-Host "Collecting GPOs..." -ForegroundColor Cyan

    $allGpos = Get-GPO -All | Sort-Object DisplayName

    $summaryResults = @()
    $linkResults = @()
    $securityResults = @()
    $errorResults = @()

    foreach ($gpo in $allGpos) {
        try {
            Write-Host ("Processing: {0}" -f $gpo.DisplayName) -ForegroundColor Yellow

            $xmlText = Get-GPOReport -Guid $gpo.Id -ReportType Xml
            $xml = [xml]$xmlText

            $wmiFilterName = Get-XmlNodeInnerText -Xml $xml -XPath '//GPO/FilterName'
            $wmiFilterAssigned = -not [string]::IsNullOrWhiteSpace($wmiFilterName)

            $somNodes = @($xml.SelectNodes('//GPO/LinksTo/SOM'))
            $hasLinks = ($somNodes.Count -gt 0)

            $computerVersion = $null
            $userVersion = $null

            try {
                if ($null -ne $gpo.Computer -and $null -ne $gpo.Computer.DSVersion) {
                    $computerVersion = $gpo.Computer.DSVersion
                }
            }
            catch { }

            try {
                if ($null -ne $gpo.User -and $null -ne $gpo.User.DSVersion) {
                    $userVersion = $gpo.User.DSVersion
                }
            }
            catch { }

            $summaryResults += [pscustomobject]@{
                GPOName           = $gpo.DisplayName
                GPOId             = $gpo.Id.Guid
                Owner             = $gpo.Owner
                CreationTime      = $gpo.CreationTime
                ModificationTime  = $gpo.ModificationTime
                ComputerVersion   = $computerVersion
                UserVersion       = $userVersion
                GPOStatus         = Get-GpoStatusText -GpoStatus ([int]$gpo.GpoStatus)
                WMIFilterName     = $wmiFilterName
                WMIFilterAssigned = $wmiFilterAssigned
                HasLinks          = $hasLinks
            }

            $linkResults += Get-GpoLinksFromXml -Xml $xml -GpoName $gpo.DisplayName -GpoId $gpo.Id.Guid
            $securityResults += Get-GpoSecurityFilteringFromXml -Xml $xml -GpoName $gpo.DisplayName -GpoId $gpo.Id.Guid
        }
        catch {
            $errorResults += [pscustomobject]@{
                GPOName       = $gpo.DisplayName
                GPOId         = $gpo.Id.Guid
                ErrorMessage  = $_.Exception.Message
            }

            Write-Warning ("Failed to process GPO '{0}': {1}" -f $gpo.DisplayName, $_.Exception.Message)
        }
    }

    $summaryResults |
        Sort-Object GPOName |
        Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8

    $linkResults |
        Sort-Object GPOName, SOMPath, LinkOrder |
        Export-Csv -Path $linksPath -NoTypeInformation -Encoding UTF8

    $securityResults |
        Sort-Object GPOName, TrusteeName |
        Export-Csv -Path $securityPath -NoTypeInformation -Encoding UTF8

    $errorResults |
        Sort-Object GPOName |
        Export-Csv -Path $errorsPath -NoTypeInformation -Encoding UTF8

    Write-Host ""
    Write-Host "Export complete." -ForegroundColor Green
    Write-Host ("Summary file:           {0}" -f $summaryPath)
    Write-Host ("Links file:             {0}" -f $linksPath)
    Write-Host ("Security filtering:     {0}" -f $securityPath)
    Write-Host ("Error log:              {0}" -f $errorsPath)
    Write-Host ""
    Write-Host ("Total GPOs discovered:  {0}" -f @($allGpos).Count) -ForegroundColor Cyan
    Write-Host ("Summary rows:           {0}" -f @($summaryResults).Count)
    Write-Host ("Link rows:              {0}" -f @($linkResults).Count)
    Write-Host ("Security rows:          {0}" -f @($securityResults).Count)
    Write-Host ("Errors logged:          {0}" -f @($errorResults).Count)

    # Optional: open output folder after completion
    # Invoke-Item $OutputFolder
}
catch {
    Write-Error ("Script failed: {0}" -f $_.Exception.Message)
    throw
}