# =============================================================================
# Get-RemoteBatteryReport.ps1
# Generates a battery health report on a remote domain machine and downloads
# it to your Documents folder.
#
# Usage:
#   .\Get-RemoteBatteryReport.ps1 -RemotePC "HOSTNAME"
#   .\Get-RemoteBatteryReport.ps1 -RemotePC "HOSTNAME" -RemoteUser "domain\username"
# =============================================================================

param (
    [Parameter(Mandatory = $true)]
    [string]$RemotePC,

    [string]$RemoteUser = "",

    [string]$LocalDestination = "$env:USERPROFILE\Documents"
)

# --- Build session options ----------------------------------------------------
$sessionParams = @{ ComputerName = $RemotePC }

# If a specific remote username was provided, prompt for credentials
if ($RemoteUser -ne "") {
    $sessionParams.Credential = Get-Credential -UserName $RemoteUser -Message "Enter password for $RemoteUser"
}

# --- Test connectivity first --------------------------------------------------
Write-Host "`nTesting connection to $RemotePC..." -ForegroundColor Cyan

try {
    Test-WSMan @sessionParams -ErrorAction Stop | Out-Null
    Write-Host "  [OK] WinRM is reachable on $RemotePC" -ForegroundColor Green
} catch {
    Write-Host "  [FAIL] Could not reach $RemotePC via WinRM." -ForegroundColor Red
    Write-Host "         Either the machine is offline, WinRM is not enabled, or you lack permissions." -ForegroundColor Red
    Write-Host "         Ask IT to confirm WinRM is enabled via Group Policy.`n" -ForegroundColor Yellow
    exit 1
}

# --- Generate the report on the remote machine --------------------------------
$remoteTempPath = "C:\Temp\battery_report_$RemotePC.html"

Write-Host "`nGenerating battery report on $RemotePC..." -ForegroundColor Cyan

try {
    Invoke-Command @sessionParams -ScriptBlock {
        param($outPath)

        # Create C:\Temp if it doesn't exist
        if (-not (Test-Path "C:\Temp")) {
            New-Item -ItemType Directory -Path "C:\Temp" | Out-Null
        }

        # Generate the battery report
        $result = powercfg /batteryreport /output $outPath 2>&1

        # powercfg exits with 0 on success — check for the file
        if (-not (Test-Path $outPath)) {
            throw "powercfg ran but the report file was not created. Output: $result"
        }

    } -ArgumentList $remoteTempPath -ErrorAction Stop

    Write-Host "  [OK] Report generated at $remoteTempPath on $RemotePC" -ForegroundColor Green

} catch {
    Write-Host "  [FAIL] Could not generate battery report on $RemotePC" -ForegroundColor Red
    Write-Host "         $($_.Exception.Message)" -ForegroundColor Red
    exit 1
}

# --- Copy the file back to your machine ---------------------------------------
$localFileName = "BatteryReport_${RemotePC}_$(Get-Date -Format 'yyyy-MM-dd').html"
$localPath     = Join-Path $LocalDestination $localFileName

Write-Host "`nCopying report to $localPath..." -ForegroundColor Cyan

try {
    $session = New-PSSession @sessionParams -ErrorAction Stop

    Copy-Item -Path $remoteTempPath `
              -Destination $localPath `
              -FromSession $session `
              -ErrorAction Stop

    Write-Host "  [OK] File copied successfully" -ForegroundColor Green

} catch {
    Write-Host "  [FAIL] Could not copy file from $RemotePC" -ForegroundColor Red
    Write-Host "         $($_.Exception.Message)" -ForegroundColor Red
    exit 1

} finally {
    # Always clean up the session and the temp file on the remote machine
    if ($session) {
        Invoke-Command -Session $session -ScriptBlock {
            param($path)
            if (Test-Path $path) { Remove-Item $path -Force }
        } -ArgumentList $remoteTempPath

        Remove-PSSession $session
    }
}

# --- Done ---------------------------------------------------------------------
Write-Host "`nDone! Battery report saved to:" -ForegroundColor Cyan
Write-Host "  $localPath`n" -ForegroundColor Green

# Open in default browser
Start-Process $localPath
