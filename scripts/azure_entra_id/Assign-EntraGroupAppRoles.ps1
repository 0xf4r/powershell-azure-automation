# Import the Microsoft Graph module (if not already imported)
# Import-Module Microsoft.Graph

# Connect to Microsoft Graph with the necessary scopes
Connect-MgGraph -Scopes "AppRoleAssignment.ReadWrite.All", "Application.Read.All", "Group.Read.All"

# Define the assignments based on your input
$assignments = @(
    @{
        appName = "APP_DEV_ABC"
        groupName = "SEC_DEV_Users"
        roles = @("Role1.Read","Role2.Write","Role3.Update","Role4.Delete")
    },
    @{
        appName = "APP_TEST_XYZ"
        groupName = "SEC_TEST_Users"
        roles = @("Role1.Read","Role2.Write")
    },
    @{
        appName = "APP_UAT_123"
        groupName = "SEC_UAT_Users"
        roles = @("Role1.Read")
    }
)

# Initialize caches for service principals and groups to avoid redundant API calls
$servicePrincipals = @{}
$groups = @{}

# Process each app-group-roles assignment
foreach ($assignment in $assignments) {
    $appName = $assignment.appName
    $groupName = $assignment.groupName
    $roles = $assignment.roles

    # Retrieve or get cached service principal
    if (-not $servicePrincipals.ContainsKey($appName)) {
        $sp = Get-MgServicePrincipal -Filter "displayName eq '$appName'"
        if ($sp -eq $null -or $sp.Count -eq 0) {
            Write-Error "Service principal '$appName' not found"
            continue
        }
        $servicePrincipals[$appName] = $sp
    }
    $servicePrincipal = $servicePrincipals[$appName]

    # Retrieve or get cached group
    if (-not $groups.ContainsKey($groupName)) {
        $grp = Get-MgGroup -Filter "displayName eq '$groupName'"
        if ($grp -eq $null -or $grp.Count -eq 0) {
            Write-Error "Group '$groupName' not found"
            continue
        }
        $groups[$groupName] = $grp
    }
    $group = $groups[$groupName]

    # Process each role for the current app and group
    foreach ($roleName in $roles) {
        # Find the app role within the service principal
        $appRole = $servicePrincipal.AppRoles | Where-Object { $_.Value -eq $roleName }
        if ($appRole -eq $null) {
            Write-Error "App role '$roleName' not found in application '$appName'"
            continue
        }
        $appRoleId = $appRole.Id

        # Create the app role assignment
        $assignmentBody = @{
            "principalId" = $group.Id
            "resourceId" = $servicePrincipal.Id
            "appRoleId" = $appRoleId
        }

        try {
            New-MgGroupAppRoleAssignment -GroupId $group.Id -BodyParameter $assignmentBody
            Write-Output "Successfully assigned role '$roleName' to group '$groupName' for app '$appName'"
        } catch {
            Write-Error "Failed to assign role '$roleName' to group '$groupName' for app '$appName': $_"
        }
    }
}
