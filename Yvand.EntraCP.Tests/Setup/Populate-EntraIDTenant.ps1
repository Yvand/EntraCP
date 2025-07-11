#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups, Microsoft.Graph.Identity.SignIns

<#
.SYNOPSIS
    Creates the users and groups in Entra ID, required to run the unit tests in EntraCP project
.DESCRIPTION
    It creates the objects only if they do not exist (no overwrite)
.LINK
    https://github.com/Yvand/EntraCP/
#>

Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"
$tenantName = (Get-MgOrganization).VerifiedDomains[0].Name

$exportedUsersFullFilePath = "C:\YvanData\dev\EntraCP_Tests_Users.csv"
$exportedGroupsFullFilePath = "C:\YvanData\dev\EntraCP_Tests_Groups.csv"

$memberUsersNamePrefix = "testEntraCPUser_"
$guestUsersNamePrefix = "testEntraCPGuestUser_"
$groupNamePrefix = "testEntraCPGroup_"
$usersCount = 999
$groupsCount = 50

$confirmation = Read-Host "Connected to tenant '$tenantName', about to create $usersCount users starting with '$memberUsersNamePrefix' and $groupsCount groups starting with '$groupNamePrefix'. Are you sure you want to proceed? [y/n]"
if ($confirmation -ne 'y') {
    Write-Warning -Message "Aborted."
    return
}

$usersWithSpecificSettings = @( 
    @{ UserPrincipalName = "$($memberUsersNamePrefix)001@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)002@$($tenantName)"; UserAttributes = @{ "GivenName" = "firstname 002" } }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)010@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)011@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)012@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)013@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)014@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)015@$($tenantName)"; IsMemberOfAllGroups = $true }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)031@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)032@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)033@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)034@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)035@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)036@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)037@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)038@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)039@$($tenantName)"; AccountEnabled = $false }
    @{ UserPrincipalName = "$($memberUsersNamePrefix)040@$($tenantName)"; AccountEnabled = $false }
)
$guestUsersList = @(
    @{ Mail = "$($guestUsersNamePrefix)001@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)002@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)003@contoso.local"; Id = ""; UserPrincipalName = "" }
)
$usersMemberOfAllGroups = [System.Linq.Enumerable]::Where($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.IsMemberOfAllGroups -eq $true })
$groupsWithSpecificSettings = @(
    @{
        GroupName              = "$($groupNamePrefix)001"
        SecurityEnabled        = $false
        EveryoneIsMember = $true
    },
    @{
        GroupName              = "$($groupNamePrefix)005"
        EveryoneIsMember = $true
    },
    @{
        GroupName       = "$($groupNamePrefix)008"
        SecurityEnabled = $false
    },
    @{
        GroupName              = "$($groupNamePrefix)018"
        SecurityEnabled        = $false
        EveryoneIsMember = $true
    },
    @{
        GroupName              = "$($groupNamePrefix)025"
        EveryoneIsMember = $true
    },
    @{
        GroupName       = "$($groupNamePrefix)028"
        SecurityEnabled = $false
    },
    @{
        GroupName              = "$($groupNamePrefix)038"
        SecurityEnabled        = $false
        EveryoneIsMember = $true
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

# Bulk add member users
$allUsersInEntra = @()
for ($i = 1; $i -le $usersCount; $i++) {
    $accountName = "$($memberUsersNamePrefix)$("{0:D3}" -f $i)"
    $userPrincipalName = "$($accountName)@$($tenantName)"
    $user = Get-MgUser -Filter "UserPrincipalName eq '$userPrincipalName'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    if ($null -eq $user) {
        $additionalUserAttributes = New-Object -TypeName HashTable
        $userHasSpecificAttributes = [System.Linq.Enumerable]::FirstOrDefault($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $userPrincipalName })
        if ($null -ne $userHasSpecificAttributes.UserAttributes) {
            $additionalUserAttributes = $userHasSpecificAttributes.UserAttributes
        }
        $accountEnabled = $true
        if ($null -ne $userHasSpecificAttributes.AccountEnabled) {
            $accountEnabled = $userHasSpecificAttributes.AccountEnabled
        }

        New-MgUser -UserPrincipalName $userPrincipalName -DisplayName $accountName -PasswordProfile $passwordProfile -AccountEnabled:$accountEnabled -MailNickName $accountName @additionalUserAttributes | Out-Null
        Write-Host "Created user '$userPrincipalName'" -ForegroundColor Green
        $user = Get-MgUser -Filter "UserPrincipalName eq '$userPrincipalName'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    }
    $allUsersInEntra += $user
}

# Bulk add guest users
foreach ($guestUser in $guestUsersList) {
    $user = Get-MgUser -Filter "Mail eq '$($guestUser.Mail)'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    if ($null -eq $user) {
        $invitedUser = New-MgInvitation -InvitedUserDisplayName $guestUser.Mail -InvitedUserEmailAddress $guestUser.Mail -SendInvitationMessage:$false -InviteRedirectUrl "https://myapplications.microsoft.com"
        Write-Host "Invited guest user $($invitedUser.InvitedUserEmailAddress)" -ForegroundColor Green
        $user = $invitedUser.InvitedUser
        $user = Get-MgUser -Filter "Mail eq '$($guestUser.Mail)'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    }
    $allUsersInEntra += $user
}

# Bulk add groups and set their membership
$allGroupsInEntra = @()
for ($i = 1; $i -le $groupsCount; $i++) {
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
    $allGroupsInEntra += $entraGroup

    if ($false -eq $entraGroupJustCreated) {
        # Remove all members
        $existingGroupMembers = Get-MgGroupMember -GroupId $entraGroup.Id -All
        foreach ($groupMember in $existingGroupMembers) {
            Remove-MgGroupMemberByRef -GroupId $entraGroup.Id -DirectoryObjectId $groupMember.Id
        }
        Write-Host "Removed all members of existing group $($entraGroup.DisplayName)." -ForegroundColor Green
    }

    # Set group membership
    $newGroupMemberIds = New-Object -TypeName "System.Collections.Generic.List[System.String]"
    if ($null -ne $groupSettings -and $groupSettings.ContainsKey("EveryoneIsMember") -and $groupSettings["EveryoneIsMember"] -eq $true) {
        # Everyone is mmember of this group
        foreach($userInEntra in $allUsersInEntra) {
            $newGroupMemberIds.Add($userInEntra.Id)
        }
    } else {
        # Only users with IsMemberOfAllGroups true are members of this group
        foreach($upnOfUserMemberOfAllGroups in $usersMemberOfAllGroups | Select-Object -ExpandProperty UserPrincipalName) {
            $upnOfUserMemberOfAllGroups
            $userInEntra = [System.Linq.Enumerable]::First($allUsersInEntra, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $upnOfUserMemberOfAllGroups })
            $newGroupMemberIds.Add($userInEntra.Id)
        }
    }

    # $newGroupMemberIds = $newGroupMemberIds | Select-Object -Unique
    foreach ($groupMemberId in $newGroupMemberIds) {
        New-MgGroupMember -GroupId $entraGroup.Id -DirectoryObjectId $groupMemberId
    }
    Write-Host "Added $($newGroupMemberIds.Count) member(s) to group $($entraGroup.DisplayName)" -ForegroundColor Green
}

# export users and groups to their CSV file
$allUsersInEntra | 
Select-Object -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled, @{ Name = "IsMemberOfAllGroups"; Expression = { if ([System.Linq.Enumerable]::FirstOrDefault($usersWithSpecificSettings, [Func[object, bool]] { param($x) $x.UserPrincipalName -like $_.UserPrincipalName }).IsMemberOfAllGroups) { $true } else { $false } } } |
Export-Csv -Path $exportedUsersFullFilePath -NoTypeInformation
Write-Host "Exported Entra users to CSV file $($exportedUsersFullFilePath)" -ForegroundColor Green

$allGroupsInEntra | 
Select-Object -Property Id, DisplayName, SecurityEnabled, 
@{ Name = "EveryoneIsMember"; Expression = { if ([System.Linq.Enumerable]::FirstOrDefault($groupsWithSpecificSettings, [Func[object, bool]] { param($x) $x.GroupName -like $_.DisplayName }).EveryoneIsMember) { $true } else { $false } } }, 
@{ Name = "GroupType"; Expression = { $_.GroupTypes[0] } } |
Export-Csv -Path $exportedGroupsFullFilePath -NoTypeInformation
Write-Host "Exported Entra groups to CSV file $($exportedGroupsFullFilePath)" -ForegroundColor Green
