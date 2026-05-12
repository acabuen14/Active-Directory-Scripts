<#
.SYNOPSIS
    Get members of Active Directory security groups.

.DESCRIPTION
    Retrieves all members from one or more AD security groups
    and exports the results to a CSV file.

.REQUIREMENTS
    - RSAT Active Directory module
    - Domain connectivity
    - Appropriate permissions

.EXAMPLE
    .\Get-SecurityGroupMembers.ps1 -GroupNames "Domain Admins"

.EXAMPLE
    .\Get-SecurityGroupMembers.ps1 -GroupNames "VPN Users","HR Team" -Recursive

#>

param (
    [Parameter(Mandatory = $true)]
    [string[]]$GroupNames,

    [switch]$Recursive,

    [string]$ExportPath = ".\SecurityGroupMembers.csv"
)

# Load AD module
Import-Module ActiveDirectory

# Store results
$Results = @()

foreach ($Group in $GroupNames) {

    Write-Host "Getting members of group: $Group" -ForegroundColor Cyan

    try {

        # Get members
        if ($Recursive) {
            $Members = Get-ADGroupMember -Identity $Group -Recursive
        }
        else {
            $Members = Get-ADGroupMember -Identity $Group
        }

        foreach ($Member in $Members) {

            # Get additional details depending on object type
            if ($Member.objectClass -eq "user") {

                $User = Get-ADUser $Member.SamAccountName -Properties DisplayName,Mail,Enabled

                $Results += [PSCustomObject]@{
                    GroupName       = $Group
                    Name            = $User.DisplayName
                    SamAccountName  = $User.SamAccountName
                    ObjectType      = "User"
                    Enabled         = $User.Enabled
                    EmailAddress    = $User.Mail
                }
            }
            elseif ($Member.objectClass -eq "computer") {

                $Computer = Get-ADComputer $Member.SamAccountName

                $Results += [PSCustomObject]@{
                    GroupName       = $Group
                    Name            = $Computer.Name
                    SamAccountName  = $Computer.SamAccountName
                    ObjectType      = "Computer"
                    Enabled         = ""
                    EmailAddress    = ""
                }
            }
            elseif ($Member.objectClass -eq "group") {

                $Results += [PSCustomObject]@{
                    GroupName       = $Group
                    Name            = $Member.Name
                    SamAccountName  = $Member.SamAccountName
                    ObjectType      = "Nested Group"
                    Enabled         = ""
                    EmailAddress    = ""
                }
            }
        }
    }
    catch {
        Write-Warning "Could not retrieve members for group: $Group"
        Write-Warning $_.Exception.Message
    }
}

# Export to CSV
$Results | Export-Csv -Path $ExportPath -NoTypeInformation -Encoding UTF8

Write-Host ""
Write-Host "Completed." -ForegroundColor Green
Write-Host "Results exported to: $ExportPath" -ForegroundColor Yellow
Write-Host "Total members found: $($Results.Count)"