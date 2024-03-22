#Requires -Modules Microsoft.Graph.Identity.SignIns, Microsoft.Graph.Users

<#
.SYNOPSIS
    Creates the users and groups in Entra ID, required to run the unit tests in EntraCP project
.DESCRIPTION
    It creates the objects only if they do not exist (no overwrite)
.LINK
    https://github.com/Yvand/EntraCP/
#>

Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All" -UseDeviceCode
$tenantName = (Get-MgOrganization).VerifiedDomains[0].Name

$memberUsersNamePrefix = "testEntraCPUser_"
$guestUsersNamePrefix = "testEntraCPGuestUser_"
$groupNamePrefix = "testEntraCPGroup_"

$confirmation = Read-Host "Connected to tenant '$tenantName' and about to process users starting with '$memberUsersNamePrefix' and groups starting with '$groupNamePrefix'. Are you sure you want to proceed? [y/n]"
if ($confirmation -ne 'y') {
    Write-Warning -Message "Aborted."
    return
}

# Set specific attributes for some users
$usersWithSpecificSettings = @( 
    @{ UserPrincipalName = "$($memberUsersNamePrefix)001@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)002@$($tenantName)"; UserAttributes = @{ "GivenName" = "firstname 002" } }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)010@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)011@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)012@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)013@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)014@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)015@$($tenantName)"; IsMemberOfAllGroups = $true }
)
$guestUsersList = @(
    @{ Mail = "$($guestUsersNamePrefix)001@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)002@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)003@contoso.local"; Id = ""; UserPrincipalName = "" }
)
#$guestUsers = @("$($guestUsersNamePrefix)001@contoso.local", "$($guestUsersNamePrefix)002@contoso.local", "$($guestUsersNamePrefix)003@contoso.local")
$groupsWithSpecificSettings = @(
    @{
        GroupName              = "$($groupNamePrefix)001"
        SecurityEnabled        = $false
        AllTestUsersAreMembers = $true
    },
    @{
        GroupName              = "$($groupNamePrefix)005"
        AllTestUsersAreMembers = $true
    },
    @{
        GroupName       = "$($groupNamePrefix)008"
        SecurityEnabled = $false
    },
    @{
        GroupName              = "$($groupNamePrefix)018"
        SecurityEnabled        = $false
        AllTestUsersAreMembers = $true
    },
    @{
        GroupName              = "$($groupNamePrefix)025"
        AllTestUsersAreMembers = $true
    },
    @{
        GroupName       = "$($groupNamePrefix)028"
        SecurityEnabled = $false
    },
    @{
        GroupName              = "$($groupNamePrefix)038"
        SecurityEnabled        = $false
        AllTestUsersAreMembers = $true
    },
    @{
        GroupName       = "$($groupNamePrefix)048"
        SecurityEnabled = $false
    }
)

$temporaryPassword = @(
    (0..9 | Get-Random ),
    ('!', '@', '#', '$', '%', '^', '&', '*', '?', ';', '+' | Get-Random),
    (0..9 | Get-Random ),
    [char](65..90 | Get-Random),
    (0..9 | Get-Random ),
    [char](97..122 | Get-Random),
    [char](97..122 | Get-Random),
    (0..9 | Get-Random ),
    [char](97..122 | Get-Random)
) -Join ''
$passwordProfile = @{
    Password                      = $temporaryPassword
    ForceChangePasswordNextSignIn = $true
}

# Bulk add users
$totalUsers = 50
for ($i = 1; $i -le $totalUsers; $i++) {
    $accountName = "$($memberUsersNamePrefix)$("{0:D3}" -f $i)"
    $userPrincipalName = "$($accountName)@$($tenantName)"
    $user = Get-MgUser -Filter "UserPrincipalName eq '$userPrincipalName'"
    if ($null -eq $user) {
        $additionalUserAttributes = New-Object -TypeName HashTable
        $userHasSpecificAttributes = [System.Linq.Enumerable]::FirstOrDefault($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $userPrincipalName })
        if ($null -ne $userHasSpecificAttributes.UserAttributes) {
            $additionalUserAttributes = $userHasSpecificAttributes.UserAttributes
        }

        New-MgUser -DisplayName $accountName -PasswordProfile $passwordProfile -AccountEnabled -MailNickName $accountName -UserPrincipalName $userPrincipalName @additionalUserAttributes
        Write-Host "Created user '$userPrincipalName'" -ForegroundColor Green
    }
}

# Add the guest users
foreach ($guestUser in $guestUsersList) {
    $guestUserInGraph = Get-MgUser -Filter "Mail eq '$($guestUser.Mail)'"
    if ($null -eq $guestUserInGraph) {
        $invitedUser = New-MgInvitation -InvitedUserDisplayName $guestUser.Mail -InvitedUserEmailAddress $guestUser.Mail -SendInvitationMessage:$false -InviteRedirectUrl "https://myapplications.microsoft.com"
        Write-Host "Invited guest user $($invitedUser.InvitedUserEmailAddress)" -ForegroundColor Green
        $guestUserInGraph = $invitedUser.InvitedUser
    }
    $guestUser.Id = $guestUserInGraph.Id
    $guestUser.UserPrincipalName = $guestUserInGraph.UserPrincipalName
}

# groups
$allMemberUsersInEntra = Get-MgUser -ConsistencyLevel eventual -Count userCount -Filter "startsWith(DisplayName, '$($memberUsersNamePrefix)')" -OrderBy UserPrincipalName
$usersMemberOfAllGroups = [System.Linq.Enumerable]::Where($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.IsMemberOfAllGroups -eq $true })

# Bulk add groups
$totalGroups = 50
for ($i = 1; $i -le $totalGroups; $i++) {
    $groupName = "$($groupNamePrefix)$("{0:D3}" -f $i)"
    $groupSettings = [System.Linq.Enumerable]::FirstOrDefault($groupsWithSpecificSettings, [Func[object, bool]] { param($x) $x.GroupName -like $groupName })
    $entraGroup = Get-MgGroup -Filter "DisplayName eq '$($groupName)'"
    $entraGroupJustCreated = $false
    if ($null -eq $entraGroup) {
        $newGroupCmdletParameters = New-Object -TypeName HashTable
        $newGroupCmdletParameters.add("SecurityEnabled", $true)
        if ($null -ne $groupSettings) {
            if ($groupSettings.ContainsKey("SecurityEnabled") -and $groupSettings["SecurityEnabled"] -eq $false) {
                $newGroupCmdletParameters["SecurityEnabled"] = $false
            }
        }

        $entraGroup = New-MgGroup -GroupTypes "Unified" -DisplayName $groupName -MailNickName $groupName -MailEnabled:$False @newGroupCmdletParameters
        Write-Host "Created group $groupName" -ForegroundColor Green
        $entraGroupJustCreated = $true
    }

    if ($false -eq $entraGroupJustCreated) {
        # Remove all members
        $existingGroupMembers = Get-MgGroupMember -GroupId $entraGroup.Id
        foreach ($groupMember in $existingGroupMembers) {
            Remove-MgGroupMemberByRef -GroupId $entraGroup.Id -DirectoryObjectId $groupMember.Id
        }
        Write-Host "Removed all members of existing group $($entraGroup.DisplayName)." -ForegroundColor Green
    }

    # Set membership
    $newGroupMembers = $usersMemberOfAllGroups | Select-Object -ExpandProperty UserPrincipalName
    $newGroupMemberIds = New-Object -TypeName "System.Collections.Generic.List[System.String]"
    if ($null -ne $groupSettings -and $groupSettings.ContainsKey("AllTestUsersAreMembers") -and $groupSettings["AllTestUsersAreMembers"] -eq $true) {
        $newGroupMembers = $allMemberUsersInEntra.UserPrincipalName

        foreach ($guestUser in $guestUsersList) {
            $newGroupMemberIds.Add($guestUser.Id)
        }
    }

    foreach ($groupMember in $newGroupMembers) {
        $entraUser = [System.Linq.Enumerable]::FirstOrDefault($allMemberUsersInEntra, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $groupMember })
        $newGroupMemberIds.Add($entraUser.Id)
    }

    # $newGroupMemberIds = $newGroupMemberIds | Select-Object -Unique
    foreach ($groupMemberId in $newGroupMemberIds) {
        New-MgGroupMember -GroupId $entraGroup.Id -DirectoryObjectId $groupMemberId
    }
    Write-Host "Added $($newGroupMemberIds.Count) member(s) to group $($entraGroup.DisplayName)" -ForegroundColor Green
}
