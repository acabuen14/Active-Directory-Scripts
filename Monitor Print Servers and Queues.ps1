function write-DRRMAlert ($message) {
    write-host '<-Start Result->'
    write-host "Alert=$message"
    write-host '<-End Result->'
    }
function write-DRMMDiag ($messages) {
    write-host  '<-Start Diagnostic->'
    foreach ($Message in $Messages) { $Message }
    write-host '<-End Diagnostic->'
}
$version = (Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion").CurrentVersion
if($Version -lt "6.2") {
write-DRRMAlert "Unsupported OS. Only Server 2012 and up are supported."
exit 1
}

$PrintStatus = @()
$Printers = get-printer
foreach($Printer in $Printers){
if($printer.PRinterStatus -ne "normal"){ $PrintStatus += "$($Printer.name) is in status: $($Printer.PrinterStatus)" }
$PrintJobs = get-PrintJob -PrinterObject $printer
foreach($job in $PrintJobs) {
    if($Job.JobStatus -ne "normal"){ $PrintStatus += "$($Printer.name) has a job with: $($Job.Jobstatus)." }
}
}

if(!$PrintStatus){
   write-DRRMAlert "Healthy" 
} else {
    write-DRMMDiag $PrintStatus
    write-DRRMAlert "Please check diagnostic info"
    exit 1
}