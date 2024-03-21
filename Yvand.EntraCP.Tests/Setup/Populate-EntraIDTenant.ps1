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

$accountNamePrefix = "testEntraCPUser_"
$groupNamePrefix = "testEntraCPGroup_"
$confirmation = Read-Host "Connected to tenant '$tenantName' and about to process users starting with '$accountNamePrefix' and groups starting with '$groupNamePrefix'. Are you sure you want to proceed? [y/n]"
if ($confirmation -ne 'y') {
    Write-Warning -Message "Aborted."
    return
}

# Set specific attributes for some users
$usersWithSpecificSettings = @( 
    @{ UserPrincipalName = "$($accountNamePrefix)001@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)002@$($tenantName)"; UserAttributes = @{ "GivenName" = "firstname 002" } }
    @{ UserPrincipalName = "$($accountNamePrefix)010@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)011@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)012@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)013@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)014@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($accountNamePrefix)015@$($tenantName)"; IsMemberOfAllGroups = $true }
)

$dtlGroupsSettings = @(
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

$guestUsers = @("testEntraCPGuestUser_001@contoso.local", "testEntraCPGuestUser_002@contoso.local", "testEntraCPGuestUser_003@contoso.local")
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
    $accountName = "$($accountNamePrefix)$("{0:D3}" -f $i)"
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
foreach ($guestUser in $guestUsers) {
    $user = Get-MgUser -Filter "Mail eq '$($guestUser)'"
    if ($null -eq $user) {
        New-MgInvitation -InvitedUserDisplayName $guestUser -InvitedUserEmailAddress $guestUser -SendInvitationMessage:$false -InviteRedirectUrl "https://myapplications.microsoft.com"
        Write-Host "Created guest user $guestUser" -ForegroundColor Green
    }
}

# groups
$allTestEntraUsers = Get-MgUser -ConsistencyLevel eventual -Count userCount -Filter "startsWith(DisplayName, '$($accountNamePrefix)')" -OrderBy UserPrincipalName
$usersMemberOfAllGroups = [System.Linq.Enumerable]::Where($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.IsMemberOfAllGroups -eq $true })

# Bulk add groups
$totalGroups = 50
for ($i = 1; $i -le $totalGroups; $i++) {
    $groupName = "$($groupNamePrefix)$("{0:D3}" -f $i)"
    $entraGroup = Get-MgGroup -Filter "DisplayName eq '$($groupName)'"
    if ($null -eq $entraGroup) {
        $newGroupCmdletParameters = New-Object -TypeName HashTable
        $newGroupCmdletParameters.add("SecurityEnabled", $true)
        $groupSettings = [System.Linq.Enumerable]::FirstOrDefault($dtlGroupsSettings, [Func[object, bool]] { param($x) $x.GroupName -like $groupName })
        if ($null -ne $groupSettings) {
            if ($groupSettings.ContainsKey("SecurityEnabled") -and $groupSettings["SecurityEnabled"] -eq $false) {
                $newGroupCmdletParameters["SecurityEnabled"] = $false
            }
        }

        $entraGroup = New-MgGroup -GroupTypes "Unified" -DisplayName $groupName -MailNickName $groupName -MailEnabled:$False @newGroupCmdletParameters
        Write-Host "Created group $groupName" -ForegroundColor Green

        # Set membership
        $groupMembers = $usersMemberOfAllGroups | Select-Object -ExpandProperty UserPrincipalName
        if ($null -ne $groupSettings -and $groupSettings.ContainsKey("AllTestUsersAreMembers") -and $groupSettings["AllTestUsersAreMembers"] -eq $true) {
            $groupMembers = $allTestEntraUsers.UserPrincipalName
        }

        $groupMemberIds = New-Object -TypeName "System.Collections.Generic.List[System.String]"
        foreach ($groupMember in $groupMembers) {
            $entraUser = [System.Linq.Enumerable]::FirstOrDefault($allTestEntraUsers, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $groupMember })
            $groupMemberIds.Add($entraUser.Id)
        }

        # $groupMemberIds = $groupMemberIds | Select-Object -Unique
        foreach ($groupMemberId in $groupMemberIds) {
            New-MgGroupMember -GroupId $entraGroup.Id -DirectoryObjectId $groupMemberId
        }
        Write-Host "Added $($groupMemberIds.Count) member(s) to group $groupName" -ForegroundColor Green
    }
}
