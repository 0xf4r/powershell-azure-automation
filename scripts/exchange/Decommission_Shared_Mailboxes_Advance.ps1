<#
.SYNOPSIS
    Decommissions shared mailboxes by adding "del_" to DisplayName and Alias, and hiding from GAL.

.DESCRIPTION
    Processes a headerless CSV (mailbox.csv) in the script directory to update shared mailboxes
    in a hybrid Exchange environment. Adds "del_" to DisplayName and Alias if not present, hides
    mailboxes from GAL, checks for alias conflicts, and exports results to output\<today’s_date>_result.csv.
    Handles disabled AD accounts and provides detailed status.

.AUTHOR
    0xf4r

.VERSION
    v1.0 Advance

.DATE CREATED
    May 19, 2025

.REQUIREMENTS
    - PowerShell 5.1 or later
    - Exchange Management Shell (Exchange 2016/2019 or Exchange Online PowerShell module for hybrid environments)
    - ActiveDirectory module (for Get-ADUser)
    
.NOTES
    Test in a non-production environment first, as the script modifies mailbox attributes.
#>

# Import ActiveDirectory module
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

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
        # Initialize variables
        $recipient = $null
        $originalDisplayName = "N/A"
        $originalAlias = "N/A"
        $currentHiddenFromGAL = $false
        $location = "N/A"
        $newDisplayName = "N/A"
        $newAlias = "N/A"
        $newHiddenFromGAL = $false
        $changesMade = @()
        $unchangedItems = @()
        $aliasConflict = $false

        # Check AD for the object
        $adUser = Get-ADUser -Filter {mail -eq $identity -or UserPrincipalName -eq $identity} -Properties mail, proxyAddresses, Enabled -ErrorAction SilentlyContinue
        if ($adUser) {
            if (-not $adUser.Enabled) {
                Write-Warning "Processing disabled AD account for $identity"
            }

            # Try to get recipient information
            $remoteMailbox = Get-RemoteMailbox -Identity $identity -ErrorAction SilentlyContinue
            if ($remoteMailbox) {
                $recipient = $remoteMailbox
                $location = "Cloud"
            } else {
                $mailbox = Get-Mailbox -Identity $identity -ErrorAction SilentlyContinue
                if ($mailbox) {
                    $recipient = $mailbox
                    $location = "On-Prem"
                }
            }

            if (-not $recipient) {
                throw "No valid mailbox or remote mailbox found for $identity."
            }

            # Capture original attributes
            $originalDisplayName = $recipient.DisplayName
            $originalAlias = $recipient.Alias
            $currentHiddenFromGAL = $recipient.HiddenFromAddressListsEnabled

            # Check if it's a shared mailbox
            if ($recipient.RecipientTypeDetails -eq 'SharedMailbox' -or $recipient.RecipientTypeDetails -eq 'RemoteSharedMailbox') {
                # Determine what needs to be changed
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
                    $proposedNewAlias = "del_$originalAlias"
                    # Check for alias conflict
                    $existingRecipient = Get-Recipient -Identity $proposedNewAlias -ErrorAction SilentlyContinue
                    if ($existingRecipient) {
                        $aliasConflict = $true
                        $unchangedItems += "Alias: Conflict with existing alias '$proposedNewAlias'"
                        $newAlias = $originalAlias
                    } else {
                        $newAlias = $proposedNewAlias
                    }
                } else {
                    $newAlias = $originalAlias
                    $unchangedItems += "Alias: Already prefixed"
                }

                if ($updateHiddenFromGAL) {
                    $newHiddenFromGAL = $true
                } else {
                    $unchangedItems += "GAL: Already hidden"
                }

                # Apply changes if needed
                if ($updateDisplayName -or ($updateAlias -and -not $aliasConflict) -or $updateHiddenFromGAL) {
                    if ($location -eq "On-Prem") {
                        $params = @{
                            Identity = $identity
                            HiddenFromAddressListsEnabled = $true
                            ErrorAction = 'Stop'
                        }
                        if ($updateDisplayName) { $params.DisplayName = $newDisplayName }
                        if ($updateAlias -and -not $aliasConflict) { $params.Alias = $newAlias }
                        Set-Mailbox @params
                    } else {
                        $params = @{
                            Identity = $identity
                            HiddenFromAddressListsEnabled = $true
                            ErrorAction = 'Stop'
                        }
                        if ($updateDisplayName) { $params.DisplayName = $newDisplayName }
                        if ($updateAlias -and -not $aliasConflict) { $params.Alias = $newAlias }
                        Set-RemoteMailbox @params
                    }

                    # Record changes made
                    if ($updateDisplayName) { $changesMade += "Name" }
                    if ($updateAlias -and -not $aliasConflict) { $changesMade += "Alias" }
                    if ($updateHiddenFromGAL) { $changesMade += "GAL" }
                } else {
                    $unchangedItems += "None: No changes needed"
                }
            } else {
                $unchangedItems = @("Not shared mailbox")
                $newDisplayName = $originalDisplayName
                $newAlias = $originalAlias
                $newHiddenFromGAL = $currentHiddenFromGAL
            }
        } else {
            throw "No AD object found for $identity."
        }

        # Store result
        $result = [PSCustomObject]@{
            PrimarySMTP       = $identity
            OldDisplayName    = $originalDisplayName
            NewDisplayName    = if ($updateDisplayName) { $newDisplayName } else { "Unchanged" }
            OldAlias          = $originalAlias
            NewAlias          = if ($updateAlias -and -not $aliasConflict) { $newAlias } elseif ($aliasConflict) { "Conflict: $proposedNewAlias" } else { "Unchanged" }
            Location          = $location
            'GAL Status'      = if ($newHiddenFromGAL) { "Hidden" } else { "Not Hidden" }
            UpdatedFields     = if ($changesMade) { $changesMade -join ', ' } else { "None" }
            SkippedFields     = if ($unchangedItems) { $unchangedItems -join ', ' } else { "None" }
        }
        $results += $result
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