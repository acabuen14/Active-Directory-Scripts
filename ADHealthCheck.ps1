# Define variables for log file path with dynamic date and hostname
$CurrentDate = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$Hostname = $env:COMPUTERNAME
$LogFilePath = "C:\Active Directory Logs\AD_Health_Check_${Hostname}_$CurrentDate.txt"

# Create log directory if it doesn't exist
if (!(Test-Path -Path "C:\Active Directory Logs")) {
    New-Item -ItemType Directory -Path "C:\Active Directory Logs" -Force
}

# Create log file with UTF-8 encoding if it doesn't exist
if (!(Test-Path -Path $LogFilePath)) {
    Out-File -FilePath $LogFilePath -Encoding UTF8
}

# Function to log messages
function Write-Log {
    param (
        [string]$Message
    )
    $Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $Message = "$Timestamp - $Message"
    Write-Output $Message | Out-File -FilePath $LogFilePath -Append -Encoding UTF8
    Write-Host $Message
}

# Start logging
Write-Log "Starting Active Directory Health Check on $Hostname..."

# Run dcdiag with comprehensive diagnostics
Write-Log "Running dcdiag with /c /v options..."
try {
    dcdiag /c /v /f:$LogFilePath
    Write-Log "dcdiag command completed successfully. Output saved to $LogFilePath."
} catch {
    Write-Log "Error while running dcdiag: $_"
    exit 1
}

# Additional commands for health check (optional)
Write-Log "Running repadmin /replsummary..."
try {
    $ReplSummary = repadmin /replsummary
    Write-Log "Replication summary completed. Results: `n$ReplSummary"
} catch {
    Write-Log "Error while running repadmin /replsummary: $_"
}

Write-Log "Running dcdiag /test:DNS..."
try {
    $DNSHealth = dcdiag /test:DNS
    Write-Log "DNS test completed. Results: `n$DNSHealth"
} catch {
    Write-Log "Error while running dcdiag /test:DNS: $_"
}

# Script completion
Write-Log "Active Directory Health Check completed on $Hostname."
Write-Log "Please review the log file at $LogFilePath for details."
