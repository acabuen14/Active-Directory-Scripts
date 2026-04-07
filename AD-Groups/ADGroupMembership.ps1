# Parameters
$GroupName = "YourGroupNameHere"  # Replace with the name of your AD Security Group
$OutputFilePath = "C:\ADGroupMembers.csv"  # Path to save the CSV file

# Import Active Directory Module
Import-Module ActiveDirectory

# Retrieve Group Members
Write-Output "Fetching members of the group: $GroupName"
$GroupMembers = Get-ADGroupMember -Identity $GroupName -Recursive | Where-Object { $_.ObjectClass -eq 'user' }

# Initialize an array to store member details
$MemberDetails = @()

# Loop through each member and get their details
foreach ($Member in $GroupMembers) {
    $User = Get-ADUser -Identity $Member.SamAccountName -Properties DisplayName, EmailAddress, Department, Title, Enabled
    $MemberDetails += [PSCustomObject]@{
        UserName      = $User.SamAccountName
        DisplayName   = $User.DisplayName
        EmailAddress  = $User.EmailAddress
        Department    = $User.Department
        Title         = $User.Title
        IsDisabled    = -not $User.Enabled  # True if the account is disabled, False otherwise
    }
}

# Export to CSV
Write-Output "Exporting group members to file: $OutputFilePath"
$MemberDetails | Export-Csv -Path $OutputFilePath -NoTypeInformation -Encoding UTF8

Write-Output "Export completed successfully!"
