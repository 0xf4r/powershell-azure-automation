# Ensure you're authenticated
Connect-AzAccount

# Read input CSV
$groups = Import-Csv -Path "C:\temp\groups-input.csv" -Header "GroupName"
$results = @()

foreach ($group in $groups) {
    $groupName = $group.GroupName
    try {
        $azGroup = Get-AzADGroup -Filter "displayName eq '$groupName'" -ErrorAction Stop
        if ($azGroup) {
            $results += [PSCustomObject]@{
                GroupName = $groupName
                GroupID   = $azGroup.Id
            }
        } else {
            $results += [PSCustomObject]@{
                GroupName = $groupName
                GroupID   = "Not Found"
            }
        }
    } catch {
        $results += [PSCustomObject]@{
            GroupName = $groupName
            GroupID   = "Error: $($_.Exception.Message)"
        }
    }
}

# Write to output CSV
$results | Export-Csv -Path "C:\temp\groups-output.csv" -NoTypeInformation