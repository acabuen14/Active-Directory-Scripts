<#
.SYNOPSIS
    Exports an AD computer inventory with category-specific outputs.

.DESCRIPTION
    Queries Active Directory for all computer accounts and returns the following fields:

    - Name
    - SamAccountName
    - DistinguishedName
    - Enabled
    - LastLogon_timestamp
    - ManagedBy
    - PasswordLastSet
    - PasswordExpired
    - WhenCreated
    - WhenChanged
    - OperatingSystem
    - IPv4Address
    - CanonicalName
    - Today
    - Days Since Lastlogon

    If the ImportExcel module is available, a single XLSX workbook is created with these worksheets:
    - All Computers
    - Windows 7
    - Windows 10
    - Windows 11
    - Windows Servers

    If the ImportExcel module is not available, separate CSV files are created for each category.

.NOTES
    PowerShell 5.1 compatible.
    Requires the ActiveDirectory module / RSAT tools.

    Default output folder:
    C:\Users\<current-user>\Documents\ADComputerInventory
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $false)]
    [string]$OutputFolder = (Join-Path -Path $env:USERPROFILE -ChildPath 'Documents\ADComputerInventory'),

    [Parameter(Mandatory = $false)]
    [string]$WorkbookName = ("ADComputerInventory_{0}.xlsx" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Convert-LastLogonTimestamp {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [object]$Value
    )

    try {
        if ($null -eq $Value) {
            return $null
        }

        if ($Value -is [datetime]) {
            return $Value
        }

        if ([string]::IsNullOrWhiteSpace([string]$Value)) {
            return $null
        }

        $int64Value = [int64]$Value
        if ($int64Value -le 0) {
            return $null
        }

        return [datetime]::FromFileTimeUtc($int64Value).ToLocalTime()
    }
    catch {
        return $null
    }
}

function Get-DaysSinceDate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [AllowNull()]
        [object]$DateValue
    )

    try {
        if ($null -eq $DateValue) {
            return $null
        }

        if (-not ($DateValue -is [datetime])) {
            return $null
        }

        return [int]((Get-Date) - $DateValue).TotalDays
    }
    catch {
        return $null
    }
}

try {
    Write-Host "Starting AD computer inventory export..." -ForegroundColor Cyan
    Write-Host ("Output folder: {0}" -f $OutputFolder) -ForegroundColor Yellow

    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw "The ActiveDirectory module is not installed or not available on this system."
    }

    Import-Module ActiveDirectory -ErrorAction Stop

    if (-not (Test-Path -Path $OutputFolder)) {
        New-Item -Path $OutputFolder -ItemType Directory -Force | Out-Null
    }

    $today = Get-Date
    $timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'

    Write-Host "Querying Active Directory for all computers..." -ForegroundColor Yellow

    $rawComputers = Get-ADComputer -Filter * -Properties `
        SamAccountName,
        DistinguishedName,
        Enabled,
        lastLogonTimestamp,
        ManagedBy,
        PasswordLastSet,
        PasswordExpired,
        WhenCreated,
        WhenChanged,
        OperatingSystem,
        IPv4Address,
        CanonicalName

    $allComputers = foreach ($computer in $rawComputers) {
        $lastLogon = Convert-LastLogonTimestamp -Value $computer.lastLogonTimestamp

        [pscustomobject]@{
            Name                   = $computer.Name
            SamAccountName         = $computer.SamAccountName
            DistinguishedName      = $computer.DistinguishedName
            Enabled                = $computer.Enabled
            LastLogon_timestamp    = $lastLogon
            ManagedBy              = $computer.ManagedBy
            PasswordLastSet        = $computer.PasswordLastSet
            PasswordExpired        = $computer.PasswordExpired
            WhenCreated            = $computer.WhenCreated
            WhenChanged            = $computer.WhenChanged
            OperatingSystem        = $computer.OperatingSystem
            IPv4Address            = $computer.IPv4Address
            CanonicalName          = $computer.CanonicalName
            Today                  = $today
            'Days Since Lastlogon' = Get-DaysSinceDate -DateValue $lastLogon
        }
    }

    $windows7 = $allComputers | Where-Object {
        $_.OperatingSystem -and $_.OperatingSystem -match '^Windows 7\b'
    }

    $windows10 = $allComputers | Where-Object {
        $_.OperatingSystem -and $_.OperatingSystem -match '^Windows 10\b'
    }

    $windows11 = $allComputers | Where-Object {
        $_.OperatingSystem -and $_.OperatingSystem -match '^Windows 11\b'
    }

    $windowsServers = $allComputers | Where-Object {
        $_.OperatingSystem -and $_.OperatingSystem -match 'Windows Server'
    }

    $importExcelAvailable = [bool](Get-Module -ListAvailable -Name ImportExcel)

    if ($importExcelAvailable) {
        Write-Host "ImportExcel detected. Creating workbook..." -ForegroundColor Green

        Import-Module ImportExcel -ErrorAction Stop

        $workbookPath = Join-Path -Path $OutputFolder -ChildPath $WorkbookName

        if (Test-Path -Path $workbookPath) {
            Remove-Item -Path $workbookPath -Force
        }

        $excelParams = @{
            AutoSize     = $true
            FreezeTopRow = $true
            BoldTopRow   = $true
            TableStyle   = 'Medium2'
            ClearSheet   = $true
        }

        $allComputers   | Sort-Object Name | Export-Excel -Path $workbookPath -WorksheetName 'All Computers'   @excelParams
        $windows7       | Sort-Object Name | Export-Excel -Path $workbookPath -WorksheetName 'Windows 7'       @excelParams
        $windows10      | Sort-Object Name | Export-Excel -Path $workbookPath -WorksheetName 'Windows 10'      @excelParams
        $windows11      | Sort-Object Name | Export-Excel -Path $workbookPath -WorksheetName 'Windows 11'      @excelParams
        $windowsServers | Sort-Object Name | Export-Excel -Path $workbookPath -WorksheetName 'Windows Servers' @excelParams

        Write-Host "Workbook created:" -ForegroundColor Cyan
        Write-Host $workbookPath
    }
    else {
        Write-Host "ImportExcel not detected. Creating CSV files..." -ForegroundColor Yellow

        $allPath     = Join-Path -Path $OutputFolder -ChildPath ("AllComputers_{0}.csv" -f $timestamp)
        $win7Path    = Join-Path -Path $OutputFolder -ChildPath ("Windows7_{0}.csv" -f $timestamp)
        $win10Path   = Join-Path -Path $OutputFolder -ChildPath ("Windows10_{0}.csv" -f $timestamp)
        $win11Path   = Join-Path -Path $OutputFolder -ChildPath ("Windows11_{0}.csv" -f $timestamp)
        $serversPath = Join-Path -Path $OutputFolder -ChildPath ("WindowsServers_{0}.csv" -f $timestamp)

        $allComputers   | Sort-Object Name | Export-Csv -Path $allPath     -NoTypeInformation -Encoding UTF8
        $windows7       | Sort-Object Name | Export-Csv -Path $win7Path    -NoTypeInformation -Encoding UTF8
        $windows10      | Sort-Object Name | Export-Csv -Path $win10Path   -NoTypeInformation -Encoding UTF8
        $windows11      | Sort-Object Name | Export-Csv -Path $win11Path   -NoTypeInformation -Encoding UTF8
        $windowsServers | Sort-Object Name | Export-Csv -Path $serversPath -NoTypeInformation -Encoding UTF8

        Write-Host "CSV files created:" -ForegroundColor Cyan
        Write-Host $allPath
        Write-Host $win7Path
        Write-Host $win10Path
        Write-Host $win11Path
        Write-Host $serversPath
    }

    Write-Host ""
    Write-Host "Summary:" -ForegroundColor Cyan
    Write-Host ("  All Computers   : {0}" -f @($allComputers).Count)
    Write-Host ("  Windows 7       : {0}" -f @($windows7).Count)
    Write-Host ("  Windows 10      : {0}" -f @($windows10).Count)
    Write-Host ("  Windows 11      : {0}" -f @($windows11).Count)
    Write-Host ("  Windows Servers : {0}" -f @($windowsServers).Count)

    Write-Host "Completed successfully." -ForegroundColor Green
}
catch {
    Write-Error ("Script failed: {0}" -f $_.Exception.Message)
    throw
}
