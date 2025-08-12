<#
.SYNOPSIS
    Decommissions shared mailboxes by adding "del_" to DisplayName and Alias, and hiding from GAL.

.DESCRIPTION
    Processes a headerless CSV (mailbox.csv) in the script directory, updates shared mailboxes
    in a hybrid Exchange environment, and exports results to output\<today’s_date>_result.csv.
    Only updates unchanged attributes and provides detailed status.

.AUTHOR
    0xf4r

.VERSION
    v1.0 Basic

.DATE CREATED
    May 19, 2025

.REQUIREMENTS
    - PowerShell 5.1 or later
    - Exchange Management Shell (Exchange 2016/2019 or Exchange Online PowerShell module for hybrid environments)
    
.NOTES
    Run in Exchange Management Shell on-premises with permissions for on-premises Exchange and Exchange Online.
#>

# Get script directory
$scriptPath = $PSScriptRoot

# Define input and output paths
$inputFile = Join-Path -Path $scriptPath -ChildPath "mailbox.csv"
$outputDir = Join-Path -Path $scriptPath -ChildPath "Output"
$outputFile = Join-Path -Path $outputDir -ChildPath ("$(Get-Date -Format 'yyyyMMdd')_result.csv")

# Validate input file
if (-not (Test-Path -Path $inputFile)) {
    Write-Error "Input file not found at $inputFile"
    exit
}

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory | Out-Null
}

# Read input file (no header)
$identities = Get-Content -Path $inputFile

# Initialize results array
$results = @()

# Process each mailbox
foreach ($identity in $identities) {
    try {
        # Get recipient details
        $recipient = Get-Recipient -Identity $identity -ErrorAction Stop
        
        # Check if it's a shared mailbox
        if ($recipient.RecipientTypeDetails -eq 'SharedMailbox' -or $recipient.RecipientTypeDetails -eq 'RemoteSharedMailbox') {
            # Capture original attributes
            $originalDisplayName = $recipient.DisplayName
            $originalAlias = $recipient.Alias
            $currentHiddenFromGAL = $recipient.HiddenFromAddressListsEnabled
            $location = if ($recipient.RecipientTypeDetails -eq 'SharedMailbox') { "On-Prem" } else { "Cloud" }
            
            # Initialize change tracking
            $changesMade = @()
            $unchangedItems = @()
            $newDisplayName = $originalDisplayName
            $newAlias = $originalAlias
            $newHiddenFromGAL = $currentHiddenFromGAL
            
            # Check what needs to be changed
            $updateDisplayName = $originalDisplayName -notlike "del_*"
            $updateAlias = $originalAlias -notlike "del_*"
            $updateHiddenFromGAL = -not $currentHiddenFromGAL
            
            # Prepare changes
            if ($updateDisplayName) {
                $newDisplayName = "del_$originalDisplayName"
            } else {
                $newDisplayName = "Unchanged"
                $unchangedItems += "Name: Already prefixed"
            }
            
            if ($updateAlias) {
                $newAlias = "del_$originalAlias"
            } else {
                $newAlias = "Unchanged"
                $unchangedItems += "Alias: Already prefixed"
            }
            
            if ($updateHiddenFromGAL) {
                $newHiddenFromGAL = $true
            } else {
                $unchangedItems += "GAL: Already hidden"
            }
            
            # Apply changes if needed
            if ($updateDisplayName -or $updateAlias -or $updateHiddenFromGAL) {
                if ($location -eq "On-Prem") {
                    $params = @{
                        Identity = $identity
                        HiddenFromAddressListsEnabled = $true
                        ErrorAction = 'Stop'
                    }
                    if ($updateDisplayName) { $params.DisplayName = $newDisplayName }
                    if ($updateAlias) { $params.Alias = $newAlias }
                    Set-Mailbox @params
                } else {
                    $params = @{
                        Identity = $identity
                        HiddenFromAddressListsEnabled = $true
                        ErrorAction = 'Stop'
                    }
                    if ($updateDisplayName) { $params.DisplayName = $newDisplayName }
                    if ($updateAlias) { $params.Alias = $newAlias }
                    Set-RemoteMailbox @params
                }
                
                # Record changes made
                if ($updateDisplayName) { $changesMade += "Name" }
                if ($updateAlias) { $changesMade += "Alias" }
                if ($updateHiddenFromGAL) { $changesMade += "GAL" }
            } else {
                $unchangedItems += "None: No changes needed"
            }
            
            # Store result
            $result = [PSCustomObject]@{
                PrimarySMTP       = $recipient.PrimarySmtpAddress
                OldDisplayName    = $originalDisplayName
                NewDisplayName    = $newDisplayName
                OldAlias          = $originalAlias
                NewAlias          = $newAlias
                Location          = $location
                'GAL Status'      = if ($newHiddenFromGAL) { "Hidden" } else { "Not Hidden" }
                UpdatedFields     = if ($changesMade) { $changesMade -join ', ' } else { "None" }
                SkippedFields     = if ($unchangedItems) { $unchangedItems -join ', ' } else { "None" }
            }
            $results += $result
        } else {
            # Log non-shared mailboxes
            $result = [PSCustomObject]@{
                PrimarySMTP       = $recipient.PrimarySmtpAddress
                OldDisplayName    = $recipient.DisplayName
                NewDisplayName    = $recipient.DisplayName
                OldAlias          = $recipient.Alias
                NewAlias          = $recipient.Alias
                Location          = "N/A"
                'GAL Status'      = if ($recipient.HiddenFromAddressListsEnabled) { "Hidden" } else { "Not Hidden" }
                UpdatedFields     = "None"
                SkippedFields     = "Not shared mailbox"
            }
            $results += $result
            Write-Warning "Recipient $identity is not a shared mailbox."
        }
    } catch {
        # Log errors
        $result = [PSCustomObject]@{
            PrimarySMTP       = $identity
            OldDisplayName    = "N/A"
            NewDisplayName    = "N/A"
            OldAlias          = "N/A"
            NewAlias          = "N/A"
            Location          = "N/A"
            'GAL Status'      = "Not Hidden"
            UpdatedFields     = "None"
            SkippedFields     = "Error: $($_.Exception.Message)"
        }
        $results += $result
        Write-Error "Error processing $identity : $($_.Exception.Message)"
    }
}

# Display results in console
$results | Format-Table -AutoSize

# Export results to CSV
$results | Export-Csv -Path $outputFile -NoTypeInformation
