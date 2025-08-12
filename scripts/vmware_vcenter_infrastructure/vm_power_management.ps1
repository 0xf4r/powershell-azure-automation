# VMPowerManagement.ps1
# Purpose: Power on or off VMs listed in vmList.txt in specified vCenters, with user-friendly prompts, progress feedback, CSV output, and error logging
# Author: 0xf4r
# Date: May 02, 2025

# Initialize script directory and paths
$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$vmListPath = Join-Path $scriptDir "vmList.txt"
$powerOnReportDir = Join-Path $scriptDir "Reports_Power_On"
$powerOffReportDir = Join-Path $scriptDir "Reports_Power_Off"
$logDir = Join-Path $scriptDir "Logs"

# Define vCenter server addresses
$vCenter1 = "vcenter_name_1" #First vCenter
$vCenter2 = "vccenter_name_2" #Second House vCenter

# Set PowerCLI configuration to ignore invalid certificates
Set-PowerCLIConfiguration -InvalidCertificateAction Ignore -Confirm:$false -ErrorAction SilentlyContinue

# Create "Logs" directory if it doesn't exist
if (-not (Test-Path $logDir)) {
    New-Item -ItemType Directory -Path $logDir -ErrorAction Stop | Out-Null
}

# Create "Reports_Power_On" directory if it doesn't exist
if (-not (Test-Path $powerOnReportDir)) {
    New-Item -ItemType Directory -Path $powerOnReportDir -ErrorAction Stop | Out-Null
}

# Create "Reports_Power_Off" directory if it doesn't exist
if (-not (Test-Path $powerOffReportDir)) {
    New-Item -ItemType Directory -Path $powerOffReportDir -ErrorAction Stop | Out-Null
}

# Generate timestamp for CSV and log filenames
$timestamp = Get-Date -Format "yyyy_MM_dd_HHmmss"
$errorLogPath = Join-Path $logDir "${timestamp}_errors_log.txt"

# Function to log errors to file
function Log-Error {
    param ($Message)
    $logMessage = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss'): $Message"
    Add-Content -Path $errorLogPath -Value $logMessage
}

# Main menu
Clear-Host
Write-Host "`n======================" -ForegroundColor Cyan
Write-Host "    VM Power Management" -ForegroundColor Cyan
Write-Host "======================" -ForegroundColor Cyan
Write-Host "[1] Power On VMs (from $vmListPath)"
Write-Host "[2] Power Off VMs (from $vmListPath)"
Write-Host "[0] Exit"
Write-Host ""
$operation = Read-Host "Enter your choice (0-2)"

# Validate main menu input
if ($operation -notin @("0", "1", "2")) {
    $errorMsg = "Invalid choice: $operation. Exiting script."
    Write-Host $errorMsg -ForegroundColor Red
    Log-Error $errorMsg
    exit
}

if ($operation -eq "0") {
    Write-Host "Exiting script." -ForegroundColor Yellow
    exit
}

# Initialize vCenter selections
$selectedVCenters = @()

# Power On sub-menu
if ($operation -eq "1") {
    Clear-Host
    Write-Host "`n==================" -ForegroundColor Cyan
    Write-Host "    Power On VMs" -ForegroundColor Cyan
    Write-Host "==================" -ForegroundColor Cyan
    Write-Host "Select vCenter(s) to power on VMs listed in ($vmListPath):"
    Write-Host "[1] Both vCenters ($vCenter1, $vCenter2)"
    Write-Host "[2] First vCenter ($vCenter1)"
    Write-Host "[3] Second vCenter ($vCenter2)"
    Write-Host "[0] Back to Main Menu"
    Write-Host ""
    $powerOnChoice = Read-Host "Enter your choice (0-3)"
    
    if ($powerOnChoice -eq "0") {
        Write-Host "Returning to main menu..." -ForegroundColor Yellow
        exit
    }
    if ($powerOnChoice -notin @("1", "2", "3")) {
        $errorMsg = "Invalid Power On choice: $powerOnChoice. Exiting script."
        Write-Host $errorMsg -ForegroundColor Red
        Log-Error $errorMsg
        exit
    }
    if ($powerOnChoice -eq "1") { $selectedVCenters = @($vCenter1, $vCenter2) }
    if ($powerOnChoice -eq "2") { $selectedVCenters = @($vCenter1) }
    if ($powerOnChoice -eq "3") { $selectedVCenters = @($vCenter2) }
}

# Power Off sub-menu
if ($operation -eq "2") {
    Clear-Host
    Write-Host "`n===================" -ForegroundColor Cyan
    Write-Host "    Power Off VMs" -ForegroundColor Cyan
    Write-Host "===================" -ForegroundColor Cyan
    Write-Host "Select vCenter(s) to power off VMs listed in $vmListPath (using guest shutdown):"
    Write-Host "[1] Both vCenters ($vCenter1, $vCenter2)"
    Write-Host "[2] First vCenter ($vCenter1)"
    Write-Host "[3] Second vCenter ($vCenter2)"
    Write-Host "[0] Back to Main Menu"
    Write-Host ""
    $powerOffChoice = Read-Host "Enter your choice (0-3)"
    
    if ($powerOffChoice -eq "0") {
        Write-Host "Returning to main menu..." -ForegroundColor Yellow
        exit
    }
    if ($powerOffChoice -notin @("1", "2", "3")) {
        $errorMsg = "Invalid Power Off choice: $powerOffChoice. Exiting script."
        Write-Host $errorMsg -ForegroundColor Red
        Log-Error $errorMsg
        exit
    }
    if ($powerOffChoice -eq "1") { $selectedVCenters = @($vCenter1, $vCenter2) }
    if ($powerOffChoice -eq "2") { $selectedVCenters = @($vCenter1) }
    if ($powerOffChoice -eq "3") { $selectedVCenters = @($vCenter2) }
}

# Get credentials for vCenter access
$credential = Get-Credential -Message "Enter vCenter credentials for $selectedVCenters"

# Connect to selected vCenters
foreach ($vc in $selectedVCenters) {
    Write-Host "Connecting to $vc..." -ForegroundColor Yellow
    try {
        Connect-VIServer -Server $vc -Credential $credential -ErrorAction Stop
    } catch {
        $errorMsg = "Failed to connect to $vc. Error: $($_.Exception.Message)"
        Write-Host $errorMsg -ForegroundColor Red
        Log-Error $errorMsg
        exit
    }
}

# Initialize result collections
$powerOnResults = @()
$powerOffResults = @()
$notFoundVMs = @()

# Power On Operation
if ($operation -eq "1") {
    # Read VMs from text file, ignoring comments and blank lines
    $VMs = Get-Content -Path $vmListPath -ErrorAction SilentlyContinue | Where-Object { $_.Trim() -notlike '#*' -and $_.Trim() -ne '' }
    $totalVMs = $VMs.Count
    $currentVM = 0
    
    foreach ($VM in $VMs) {
        $currentVM++
        # Show progress bar for power-on operation
        Write-Progress -Activity "Processing Power On for VMs" -Status "VM: $VM ($currentVM/$totalVMs)" -PercentComplete (($currentVM / $totalVMs) * 100)
        $vmFound = $false
        foreach ($vc in $selectedVCenters) {
            try {
                $vmObject = Get-VM -Name $VM -Server $vc -ErrorAction SilentlyContinue
                if ($vmObject) {
                    $vmFound = $true
                    $powerStateBefore = $vmObject.PowerState
                    Write-Host "Initiating Power On for VM $VM on $vc" -ForegroundColor DarkGreen
                    Start-VM -VM $vmObject -Confirm:$false -ErrorAction SilentlyContinue
                    Start-Sleep -Seconds 10
                    $powerStateNow = (Get-VM -Name $VM -Server $vc).PowerState
                    $powerOnResults += [PSCustomObject]@{
                        VMname          = $VM
                        vCentre         = $vc
                        PowerStateBefore = $powerStateBefore
                        PowerStateNow   = $powerStateNow
                        DNSName         = $null
                        IP              = $null
                    }
                }
            } catch {
                $errorMsg = "Error checking VM $VM on $vc. Error: $($_.Exception.Message)"
                Write-Host $errorMsg -ForegroundColor Red
                Log-Error $errorMsg
            }
        }
        if (-not $vmFound) {
            $notFoundVMs += $VM
        }
    }
    Write-Progress -Activity "Processing Power On for VMs" -Completed

    # Wait 60 seconds for VMs to stabilize before fetching DNS/IP
    if ($powerOnResults.Count -gt 0) {
        Write-Host "Waiting 60 seconds for powered-on VMs to stabilize..." -ForegroundColor Yellow
        Start-Sleep -Seconds 60

        # Fetch DNS and IP for powered-on VMs
        $currentVM = 0
        foreach ($vm in $powerOnResults) {
            $currentVM++
            Write-Progress -Activity "Fetching DNS/IP for Powered-On VMs" -Status "VM: $($vm.VMname) ($currentVM/$($powerOnResults.Count))" -PercentComplete (($currentVM / $powerOnResults.Count) * 100)
            try {
                $vmObject = Get-VM -Name $vm.VMname -Server $vm.vCentre -ErrorAction SilentlyContinue
                if ($vmObject -and $vmObject.Guest) {
                    $vm.DNSName = if ($null -eq $vmObject.Guest.HostName) { "N/A" } else { $vmObject.Guest.HostName }
                    $vm.IP = if ($null -eq $vmObject.Guest.IPAddress) { "N/A" } else { $vmObject.Guest.IPAddress -join ", " }
                } else {
                    $vm.DNSName = "N/A"
                    $vm.IP = "N/A"
                }
            } catch {
                $errorMsg = "Error fetching DNS/IP for VM $($vm.VMname) on $($vm.vCentre). Error: $($_.Exception.Message)"
                $vm.DNSName = "Error"
                $vm.IP = "Error"
                Log-Error $errorMsg
            }
        }
        Write-Progress -Activity "Fetching DNS/IP for Powered-On VMs" -Completed
    }

    # Display Power On results on screen
    Write-Host "`n=== Power On Results ===" -ForegroundColor Cyan
    if ($powerOnResults.Count -eq 0) {
        Write-Host "No VMs were powered on." -ForegroundColor Yellow
    } else {
        $powerOnResults | Sort-Object VMname, vCentre | Format-Table -Property VMname, vCentre, PowerStateBefore, PowerStateNow, DNSName, IP -AutoSize
    }

    # Export Power On results to CSV
    $powerOnLogPath = Join-Path $powerOnReportDir "${timestamp}_PowerOn.csv"
    Write-Host "Preparing to export Power On results to: $powerOnLogPath" -ForegroundColor Yellow
    if ($powerOnResults.Count -gt 0) {
        try {
            $powerOnResults | Export-Csv -Path $powerOnLogPath -NoTypeInformation -ErrorAction Stop
            Write-Host "Power On results exported to $powerOnLogPath" -ForegroundColor Green
        } catch {
            $errorMsg = "Failed to export Power On results to $powerOnLogPath. Error: $($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            Log-Error $errorMsg
        }
    }

    # Display not-found VMs
    Write-Host "=== VMs Not Found in Selected vCenter(s) ===" -ForegroundColor Yellow
    if ($notFoundVMs.Count -eq 0) {
        Write-Host "All VMs were found." -ForegroundColor Green
    } else {
        $notFoundVMs | ForEach-Object { Write-Host $_ }
    }
}

# Power Off Operation
if ($operation -eq "2") {
    # Confirm power-off operation
    Write-Host "`nWARNING: This will power off VMs listed in $vmListPath using guest shutdown." -ForegroundColor Red
    $confirmation = Read-Host "Are you sure you want to proceed? (Y/N)"
    if ($confirmation -notlike "Y*") {
        Write-Host "Power Off operation cancelled." -ForegroundColor Yellow
        exit
    }

    # Read VMs from text file, ignoring comments and blank lines
    $VMs = Get-Content -Path $vmListPath -ErrorAction SilentlyContinue | Where-Object { $_.Trim() -notlike '#*' -and $_.Trim() -ne '' }
    $totalVMs = $VMs.Count
    $currentVM = 0
    
    foreach ($VM in $VMs) {
        $currentVM++
        # Show progress bar for power-off operation
        Write-Progress -Activity "Processing Power Off for VMs" -Status "VM: $VM ($currentVM/$totalVMs)" -PercentComplete (($currentVM / $totalVMs) * 100)
        $vmFound = $false
        foreach ($vc in $selectedVCenters) {
            try {
                $vmObject = Get-VM -Name $VM -Server $vc -ErrorAction SilentlyContinue
                if ($vmObject) {
                    $vmFound = $true
                    $powerStateBefore = $vmObject.PowerState
                    $result = "N/A"
                    if ($powerStateBefore -eq "PoweredOn") {
                        try {
                            Write-Host "Powering off VM $VM on $vc (Guest Shutdown)" -ForegroundColor DarkGreen
                            Stop-VM -VM $vmObject -Confirm:$false -ErrorAction Stop
                            Start-Sleep -Seconds 5
                            $powerStateAfter = (Get-VM -Name $VM -Server $vc).PowerState
                            $result = "Success"
                        } catch {
                            $powerStateAfter = (Get-VM -Name $VM -Server $vc).PowerState
                            $result = "Failed"
                            $errorMsg = "Failed to power off VM $VM on $vc. Error: $($_.Exception.Message)"
                            Write-Host $errorMsg -ForegroundColor Red
                            Log-Error $errorMsg
                        }
                    } else {
                        $powerStateAfter = $powerStateBefore
                        $result = "Already Off"
                    }
                    $powerOffResults += [PSCustomObject]@{
                        VMname           = $VM
                        vCentre          = $vc
                        PowerStateBefore = $powerStateBefore
                        PowerStateAfter  = $powerStateAfter
                        Result           = $result
                    }
                }
            } catch {
                $errorMsg = "Error checking VM $VM on $vc. Error: $($_.Exception.Message)"
                Write-Host $errorMsg -ForegroundColor Red
                Log-Error $errorMsg
            }
        }
        if (-not $vmFound) {
            $notFoundVMs += $VM
        }
    }
    Write-Progress -Activity "Processing Power Off for VMs" -Completed

    # Display Power Off results on screen
    Write-Host "`n=== Power Off Results ===" -ForegroundColor Cyan
    if ($powerOffResults.Count -eq 0) {
        Write-Host "No VMs were powered off." -ForegroundColor Yellow
    } else {
        $powerOffResults | Sort-Object VMname, vCentre | Format-Table -Property VMname, vCentre, PowerStateBefore, PowerStateAfter, Result -AutoSize
    }

    # Export Power Off results to CSV
    $powerOffLogPath = Join-Path $powerOffReportDir "${timestamp}_PowerOff.csv"
    Write-Host "Preparing to export Power Off results to: $powerOffLogPath" -ForegroundColor Yellow
    if ($powerOffResults.Count -gt 0) {
        try {
            $powerOffResults | Export-Csv -Path $powerOffLogPath -NoTypeInformation -ErrorAction Stop
            Write-Host "Power Off results exported to $powerOffLogPath" -ForegroundColor Green
        } catch {
            $errorMsg = "Failed to export Power Off results to $powerOffLogPath. Error: $($_.Exception.Message)"
            Write-Host $errorMsg -ForegroundColor Red
            Log-Error $errorMsg
        }
    }

    # Display not-found VMs
    Write-Host "=== VMs Not Found in Selected vCenter(s) ===" -ForegroundColor Yellow
    if ($notFoundVMs.Count -eq 0) {
        Write-Host "All VMs were found." -ForegroundColor Green
    } else {
        $notFoundVMs | ForEach-Object { Write-Host $_ }
    }
}

# Pause for user review
Write-Host "`nPress Enter to exit..." -ForegroundColor Cyan
Pause

# Disconnect from all vCenters
Disconnect-VIServer -Server * -Confirm:$false -ErrorAction SilentlyContinue