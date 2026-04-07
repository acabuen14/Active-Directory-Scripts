$root   = 'Drive Letter\Folder' #Source drive in C which is a network drive
$acct   = 'DOMAIN\username'
$log    = 'C:\Logs\Folder name\icacls_grant.log' #directory for logs
$csvBad = 'C:\Logs\Folder name\Failures.csv' #directory for logs
$csvAll = 'C:\Logs\Folder name\AllResults.csv'   # optional

# 1) Run and capture both stdout + stderr
icacls $root /grant "${acct}:(OI)(CI)M" /T /C > $log 2>&1

# 2) Parse FAILURES -> CSV
$fail = Select-String -Path $log -Pattern 'Access is denied' -Context 1,0 |
  ForEach-Object {
    # line before the error is the path icacls was working on
    $path = $_.Context.PreContext | Where-Object { $_ } | Select-Object -Last 1
    [pscustomobject]@{
      Path   = $path.Trim()
      Action = 'Grant Modify'
      Status = 'Failed'
      Reason = 'Access is denied'
    }
  }

$fail | Export-Csv $csvBad -NoTypeInformation -Encoding UTF8
Write-Host "Wrote failures to $csvBad"

# (Optional) also capture SUCCESSES into a single CSV
$ok = Select-String -Path $log -Pattern '^processed file:' |
  ForEach-Object {
    $path = ($_.Line -replace '^processed file:\s*','').Trim()
    [pscustomobject]@{
      Path   = $path
      Action = 'Grant Modify'
      Status = 'Success'
      Reason = ''
    }
  }

$all = @(); $all += $ok; $all += $fail
$all | Export-Csv $csvAll -NoTypeInformation -Encoding UTF8
Write-Host "Wrote all results to $csvAll"
