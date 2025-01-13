
# Import required module
Import-Module ActiveDirectory

# Function to calculate computer age
function Get-ComputerAge {
    param (
        [datetime]$InstallDate
    )

    # Calculate the time difference between the installation date and the current date
    $computerAge = (Get-Date) - $InstallDate

    # Return the computer age in years, months, and days
    return "{0} years, {1} months, and {2} days" -f [math]::Floor($computerAge.Days / 365), [math]::Floor(($computerAge.Days % 365) / 30), [math]::Floor($computerAge.Days % 30)
}

# Get all computers from Active Directory
$computers = Get-ADComputer -Filter * -Properties whenCreated, OperatingSystem

# Create an array to store the results
$computerAges = @()

# Loop through each computer and calculate its age
foreach ($computer in $computers) {
    $installDate = $computer.whenCreated
    $age = Get-ComputerAge -InstallDate $installDate
    $computerInfo = [PSCustomObject]@{
        Name        = $computer.Name
        OS          = $computer.OperatingSystem
        InstallDate = $installDate
        Age         = $age
    }
    $computerAges += $computerInfo
}

# Specify the path to export the CSV file
$csvFilePath = "C:\Path\To\Export\ComputerAges.csv"

# Export the results to a CSV file
$computerAges | Export-Csv -Path $csvFilePath -NoTypeInformation

Write-Host "Computer ages exported to $csvFilePath"
