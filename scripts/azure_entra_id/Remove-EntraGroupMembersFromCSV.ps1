# Connect to Microsoft 365 (run once per session)
Connect-MgGraph -Scopes "GroupMember.ReadWrite.All", "User.Read.All"

# Define variables
$groupName = "AZ_GRP_APP_ABC_ReadOnly"  # Replace with your group name
$csvPath = "C:\temp\members-list.csv"  # Replace with your CSV file path

# Read the CSV file (no header, assuming one UPN per line)
$members = Get-Content -Path $csvPath

# Get the group ID
$group = Get-MgGroup -Filter "displayName eq '$groupName'"
if (-not $group) {
    Write-Host "Group $groupName not found."
    exit
}

# Loop through each member and remove from the group
foreach ($member in $members) {
    try {
        # Trim any whitespace
        $member = $member.Trim()
        
        # Get the user ID by UPN
        $user = Get-MgUser -Filter "userPrincipalName eq '$member'"
        if (-not $user) {
            Write-Host "User $member not found."
            continue
        }
        
        # Remove member from the group
        Remove-MgGroupMemberByRef -GroupId $group.Id -DirectoryObjectId $user.Id
        Write-Host "Removed $member from $groupName successfully."
    }
    catch {
        Write-Host "Failed to remove $member from $groupName. Error: $_"
    }
}