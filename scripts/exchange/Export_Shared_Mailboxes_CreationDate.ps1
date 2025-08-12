# Exchange Online shared mailbox export
# Connect to Exchange Online (Run this from a machine with the Exchange Online PowerShell Module installed)

$UserCredential = Get-Credential
Connect-ExchangeOnline -UserPrincipalName $UserCredential.UserName -Password $UserCredential.Password

# Fetch shared mailboxes and export to CSV
$SharedMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $Owners = Get-MailboxPermission $_.PrimarySmtpAddress | Where-Object { $_.AccessRights -eq "FullAccess" } | ForEach-Object { $_.User }
    $Members = Get-RecipientPermission $_.PrimarySmtpAddress | ForEach-Object { $_.Trustee }
    $LastActivity = Get-MailboxStatistics $_.PrimarySmtpAddress | Select-Object -ExpandProperty LastLogonTime

    [PSCustomObject]@{
        MailboxName    = $_.DisplayName
        SMTPAddress    = $_.PrimarySmtpAddress
        Alias          = $_.Alias
        Owners         = ($Owners -join ",")
        Members        = ($Members -join ",")
        CreationDate   = $_.WhenCreated
        LastActivity   = $LastActivity
    }
}

$SharedMailboxes | Export-Csv -Path "C:\Temp\SharedMailboxes_Online.csv" -NoTypeInformation

# Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false

# Exchange On-Prem shared mailbox export (Run this from an Exchange Management Shell on-premise)
$SharedMailboxesOnPrem = Get-Mailbox -RecipientTypeDetails SharedMailbox | ForEach-Object {
    $Owners = Get-MailboxPermission $_.PrimarySmtpAddress | Where-Object { $_.AccessRights -eq "FullAccess" } | ForEach-Object { $_.User }
    $Members = Get-RecipientPermission $_.PrimarySmtpAddress | ForEach-Object { $_.Trustee }
    $CreationDate = Get-MailboxStatistics $_.PrimarySmtpAddress | Select-Object -ExpandProperty WhenMailboxCreated
    $LastActivity = Get-MailboxStatistics $_.PrimarySmtpAddress | Select-Object -ExpandProperty LastLogonTime

    [PSCustomObject]@{
        MailboxName    = $_.DisplayName
        SMTPAddress    = $_.PrimarySmtpAddress
        Alias          = $_.Alias
        Owners         = ($Owners -join ",")
        Members        = ($Members -join ",")
        CreationDate   = $CreationDate
        LastActivity   = $LastActivity
    }
}

$SharedMailboxesOnPrem | Export-Csv -Path "C:\Temp\SharedMailboxes_OnPrem.csv" -NoTypeInformation


Connect-ExchangeOnline

Get-Mailbox -ResultSize Unlimited | select Displayname, PrimarySmtpAddress, RecipientTypeDetails | Export-Csv "C:\mailbox_report.csv" -NoTypeInformation


# Exchange Online shared mailbox export
$UserCredential = Get-Credential
Connect-ExchangeOnline -UserPrincipalName $UserCredential.UserName -Password $UserCredential.Password

# Fetch shared mailboxes and distribution lists
$Mailboxes = Get-Mailbox -ResultSize Unlimited -RecipientTypeDetails SharedMailbox, MailUniversalDistributionGroup | ForEach-Object {
    
    # Fetching mailbox permissions (Owners)
    $Owners = Get-MailboxPermission $_.PrimarySmtpAddress | Where-Object { $_.AccessRights -eq "FullAccess" -and $_.IsInherited -eq $false } | ForEach-Object { $_.User }

    # Fetching members (only applicable for Distribution Lists)
    $Members = if ($_.RecipientTypeDetails -eq 'MailUniversalDistributionGroup') {
        Get-DistributionGroupMember $_.PrimarySmtpAddress | ForEach-Object { $_.PrimarySmtpAddress }
    } else {
        @() # No members for shared mailboxes
    }

    # Fetching last activity and creation date
    $MailboxStats = Get-MailboxStatistics $_.PrimarySmtpAddress
    $LastActivity = $MailboxStats.LastLogonTime
    $CreationDate = $MailboxStats.WhenMailboxCreated

    # Create the output object
    [PSCustomObject]@{
        MailboxName    = $_.DisplayName
        SMTPAddress    = $_.PrimarySmtpAddress
        RecipientType  = $_.RecipientTypeDetails
        Owners         = ($Owners -join ",")
        Members        = ($Members -join ",")
        CreationDate   = $CreationDate
        LastActivity   = $LastActivity
    }
}

# Export the results to CSV
$Mailboxes | Export-Csv -Path "C:\Temp\SharedMailboxes_DistributionLists.csv" -NoTypeInformation

# Disconnect Exchange Online session
Disconnect-ExchangeOnline -Confirm:$false

