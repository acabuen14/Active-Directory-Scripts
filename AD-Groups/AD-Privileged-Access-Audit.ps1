[CmdletBinding()]
param(
    [string]$OutputRoot = 'C:\ProgramData\Datto\AD-Privileged-Access-Audit',
    [switch]$IncludeEntra,
    [switch]$IncludeExtendedGroups,
    [int]$KeepRuns = 4
)

$ErrorActionPreference = 'Stop'
New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
$RunFolder = Join-Path $OutputRoot (Get-Date -Format 'yyyyMMdd-HHmmss')
New-Item -Path $RunFolder -ItemType Directory -Force | Out-Null
$Transcript = Join-Path $RunFolder 'Collection-Transcript.txt'
Start-Transcript -Path $Transcript -Force | Out-Null
$TranscriptActive = $true

function Export-Safe { param($Data, [string]$Path)
    if ($null -eq $Data) { @() | Export-Csv $Path -NoTypeInformation -Encoding UTF8 }
    else { @($Data) | Export-Csv $Path -NoTypeInformation -Encoding UTF8 }
}

try {
    Write-Host "Starting read-only audit collection: $RunFolder"
    if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
        throw 'ActiveDirectory PowerShell module is not installed.'
    }
    Import-Module ActiveDirectory
    $domain = Get-ADDomain
    $forest = Get-ADForest
    [PSCustomObject]@{
        CollectionDateUtc = (Get-Date).ToUniversalTime().ToString('o')
        ComputerName = $env:COMPUTERNAME
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        DomainFqdn = $domain.DNSRoot
        DomainNetBIOS = $domain.NetBIOSName
        ForestFqdn = $forest.Name
        IncludeEntra = $IncludeEntra.IsPresent
        IncludeExtendedGroups = $IncludeExtendedGroups.IsPresent
    } | ConvertTo-Json | Set-Content (Join-Path $RunFolder 'Collection-Metadata.json') -Encoding UTF8

    Export-Safe (Get-ADDomainController -Filter * | Select-Object HostName,IPv4Address,Site,OperatingSystem,IsGlobalCatalog,IsReadOnly) (Join-Path $RunFolder 'Domain-Controllers.csv')

    $GroupNames = @('Domain Admins','Enterprise Admins','Administrators','Schema Admins')
    if ($IncludeExtendedGroups) { $GroupNames += @('Account Operators','Server Operators','Backup Operators','Print Operators','Group Policy Creator Owners','DnsAdmins') }
    $GroupNames = $GroupNames | Select-Object -Unique

    $GroupInventory = foreach ($name in $GroupNames) {
        try {
            $g = Get-ADGroup $name -Properties Description,ManagedBy,WhenCreated,WhenChanged
            [PSCustomObject]@{ GroupName=$g.Name; SamAccountName=$g.SamAccountName; GroupScope=$g.GroupScope; GroupCategory=$g.GroupCategory; DistinguishedName=$g.DistinguishedName; Description=$g.Description; ManagedBy=$g.ManagedBy; WhenCreated=$g.WhenCreated; WhenChanged=$g.WhenChanged; Status='Found' }
        } catch {
            [PSCustomObject]@{ GroupName=$name; Status='Not found'; DistinguishedName=$null }
        }
    }
    Export-Safe $GroupInventory (Join-Path $RunFolder 'AD-Privileged-Group-Inventory.csv')

    $Direct = foreach ($name in $GroupNames) {
        try {
            $g = Get-ADGroup $name
            foreach ($m in Get-ADGroupMember $g.DistinguishedName) {
                [PSCustomObject]@{ PrivilegedGroup=$g.Name; MembershipType='Direct'; MemberName=$m.Name; SamAccountName=$m.SamAccountName; ObjectClass=$m.objectClass; DistinguishedName=$m.DistinguishedName; CollectionStatus='Collected' }
            }
        } catch { [PSCustomObject]@{ PrivilegedGroup=$name; MembershipType='Direct'; CollectionStatus="Error: $($_.Exception.Message)" } }
    }
    Export-Safe $Direct (Join-Path $RunFolder 'AD-Privileged-Group-Direct-Membership.csv')

    $Recursive = foreach ($name in $GroupNames) {
        try {
            $g = Get-ADGroup $name
            foreach ($m in Get-ADGroupMember $g.DistinguishedName -Recursive) {
                [PSCustomObject]@{ PrivilegedGroup=$g.Name; MembershipType='Recursive'; MemberName=$m.Name; SamAccountName=$m.SamAccountName; ObjectClass=$m.objectClass; DistinguishedName=$m.DistinguishedName; CollectionStatus='Collected' }
            }
        } catch { [PSCustomObject]@{ PrivilegedGroup=$name; MembershipType='Recursive'; CollectionStatus="Error: $($_.Exception.Message)" } }
    }
    Export-Safe $Recursive (Join-Path $RunFolder 'AD-Privileged-Group-Recursive-Membership.csv')

    $UserDns = $Recursive | Where-Object { $_.ObjectClass -eq 'user' -and $_.DistinguishedName } | Select-Object -ExpandProperty DistinguishedName -Unique
    $UserDetails = foreach ($dn in $UserDns) {
        try {
            $u = Get-ADUser $dn -Properties Enabled,LockedOut,AccountExpirationDate,LastLogonDate,PasswordLastSet,PasswordNeverExpires,PasswordNotRequired,AdminCount,Description,Department,Title,Manager,ServicePrincipalName,SIDHistory,WhenCreated,WhenChanged
            [PSCustomObject]@{ SamAccountName=$u.SamAccountName; UserPrincipalName=$u.UserPrincipalName; DisplayName=$u.Name; Enabled=$u.Enabled; LockedOut=$u.LockedOut; AccountExpirationDate=$u.AccountExpirationDate; LastLogonDate=$u.LastLogonDate; PasswordLastSet=$u.PasswordLastSet; PasswordNeverExpires=$u.PasswordNeverExpires; PasswordNotRequired=$u.PasswordNotRequired; AdminCount=$u.AdminCount; HasServicePrincipalName=[bool]$u.ServicePrincipalName; HasSIDHistory=[bool]$u.SIDHistory; Description=$u.Description; Department=$u.Department; Title=$u.Title; Manager=$u.Manager; WhenCreated=$u.WhenCreated; WhenChanged=$u.WhenChanged; DistinguishedName=$u.DistinguishedName }
        } catch { [PSCustomObject]@{ DistinguishedName=$dn; CollectionStatus="Error: $($_.Exception.Message)" } }
    }
    Export-Safe $UserDetails (Join-Path $RunFolder 'AD-Privileged-User-Details.csv')

    Export-Safe (Get-ADUser -LDAPFilter '(adminCount=1)' -Properties Enabled,LastLogonDate,PasswordLastSet,PasswordNeverExpires,AdminCount,Description | Select-Object SamAccountName,UserPrincipalName,Name,Enabled,LastLogonDate,PasswordLastSet,PasswordNeverExpires,AdminCount,Description,DistinguishedName) (Join-Path $RunFolder 'AD-AdminCount-Users.csv')
    Export-Safe (Get-ADUser -LDAPFilter '(sIDHistory=*)' -Properties Enabled,LastLogonDate,SIDHistory | Select-Object SamAccountName,UserPrincipalName,Name,Enabled,LastLogonDate,SIDHistory,DistinguishedName) (Join-Path $RunFolder 'AD-Users-With-SIDHistory.csv')

    if ($IncludeEntra) {
        foreach ($module in @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement')) {
            if (-not (Get-Module -ListAvailable -Name $module)) { throw "$module is required when -IncludeEntra is used." }
        }
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Identity.DirectoryManagement
        Connect-MgGraph -Scopes 'RoleManagement.Read.Directory','Directory.Read.All' -NoWelcome
        $role = Get-MgRoleManagementDirectoryRoleDefinition -Filter "displayName eq 'Global Administrator'"
        if (-not $role) { throw 'Global Administrator role definition was not found.' }
        $active = Get-MgRoleManagementDirectoryRoleAssignment -Filter "roleDefinitionId eq '$($role.Id)'" -All
        $eligible = Get-MgRoleManagementDirectoryRoleEligibilityScheduleInstance -Filter "roleDefinitionId eq '$($role.Id)'" -All
        $scheduled = Get-MgRoleManagementDirectoryRoleAssignmentScheduleInstance -Filter "roleDefinitionId eq '$($role.Id)'" -All
        $out = foreach ($a in $active) { [PSCustomObject]@{AssignmentType='Active';RoleName='Global Administrator';RoleDefinitionId=$a.RoleDefinitionId;PrincipalId=$a.PrincipalId;DirectoryScopeId=$a.DirectoryScopeId;AssignmentId=$a.Id;StartDateTime=$null;EndDateTime=$null} }
        $out += foreach ($a in $eligible) { [PSCustomObject]@{AssignmentType='Eligible';RoleName='Global Administrator';RoleDefinitionId=$a.RoleDefinitionId;PrincipalId=$a.PrincipalId;DirectoryScopeId=$a.DirectoryScopeId;AssignmentId=$a.Id;StartDateTime=$a.StartDateTime;EndDateTime=$a.EndDateTime} }
        $out += foreach ($a in $scheduled) { [PSCustomObject]@{AssignmentType='Scheduled';RoleName='Global Administrator';RoleDefinitionId=$a.RoleDefinitionId;PrincipalId=$a.PrincipalId;DirectoryScopeId=$a.DirectoryScopeId;AssignmentId=$a.Id;StartDateTime=$a.AssignmentScheduleInfo.StartDateTime;EndDateTime=$a.AssignmentScheduleInfo.Expiration.DateTime} }
        Export-Safe $out (Join-Path $RunFolder 'Entra-Global-Administrator-Assignments.csv')
        Disconnect-MgGraph | Out-Null
    }
    # Stop the transcript before compression so Collection-Transcript.txt is no longer locked.
    if ($TranscriptActive) {
        Stop-Transcript | Out-Null
        $TranscriptActive = $false
    }

    $zipPath = Join-Path $OutputRoot ("AD-Privileged-Access-Audit_{0}.zip" -f (Split-Path $RunFolder -Leaf))
    Compress-Archive -Path (Join-Path $RunFolder '*') -DestinationPath $zipPath -Force

    if ($KeepRuns -gt 0) {
        Get-ChildItem -Path $OutputRoot -Directory |
            Where-Object { $_.Name -match '^\d{8}-\d{6}$' } |
            Sort-Object Name -Descending |
            Select-Object -Skip $KeepRuns |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
    }

    $csvCount = @(Get-ChildItem -Path $RunFolder -Filter '*.csv').Count
    Write-Host "AUDIT_STATUS=SUCCESS"
    Write-Host "AUDIT_CSV_COUNT=$csvCount"
    Write-Host "AUDIT_ARCHIVE=$zipPath"
}
finally {
    if ($TranscriptActive) { Stop-Transcript | Out-Null }
}
