#Requires -Modules Microsoft.Graph.Authentication, Microsoft.Graph.Identity.DirectoryManagement, Microsoft.Graph.Users, Microsoft.Graph.Groups

<#
.SYNOPSIS
    Remove all test users and groups in Entra ID, created to run the unit tests for EntraCP project
.LINK
    https://github.com/Yvand/EntraCP/
#>

Connect-MgGraph -Scopes "User.ReadWrite.All", "Group.ReadWrite.All"
$tenantName = (Get-MgOrganization).VerifiedDomains[0].Name

$memberUsersNamePrefix = "testEntraCPUser_"
$guestUsersNamePrefix = "testEntraCPGuestUser_"
$groupNamePrefix = "testEntraCPGroup_"
$usersCount = 999
$groupsCount = 50

$confirmation = Read-Host "Connected to tenant '$tenantName', about to remove $usersCount users starting with '$memberUsersNamePrefix' and $groupsCount groups starting with '$groupNamePrefix'. Are you sure you want to proceed? [y/n]"
if ($confirmation -ne 'y') {
    Write-Warning -Message "Aborted."
    return
}

$guestUsersList = @(
    @{ Mail = "$($guestUsersNamePrefix)001@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)002@contoso.local"; Id = ""; UserPrincipalName = "" }
    @{ Mail = "$($guestUsersNamePrefix)003@contoso.local"; Id = ""; UserPrincipalName = "" }
)

# Bulk remove member users
for ($i = 1; $i -le $usersCount; $i++) {
    $accountName = "$($memberUsersNamePrefix)$("{0:D3}" -f $i)"
    $userPrincipalName = "$($accountName)@$($tenantName)"
    $user = Get-MgUser -Filter "UserPrincipalName eq '$userPrincipalName'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    if ($null -ne $user) {
        Remove-MgUser -UserId $user.Id
        Write-Host "Removed user '$($user.UserPrincipalName)'" -ForegroundColor Green
    }
}

# Bulk remove guest users
foreach ($guestUser in $guestUsersList) {
    $user = Get-MgUser -Filter "Mail eq '$($guestUser.Mail)'" -Property Id, UserPrincipalName, Mail, UserType, DisplayName, GivenName, AccountEnabled
    if ($null -ne $user) {
        Remove-MgUser -UserId $user.Id
        Write-Host "Removed user '$($user.UserPrincipalName)'" -ForegroundColor Green
    }
}

# Bulk remove groups
for ($i = 1; $i -le $groupsCount; $i++) {
    $groupName = "$($groupNamePrefix)$("{0:D3}" -f $i)"
    $entraGroup = Get-MgGroup -Filter "DisplayName eq '$($groupName)'"
    if ($null -ne $entraGroup) {
        Remove-MgGroup -GroupId $entraGroup.Id
        Write-Host "Removed group $groupName" -ForegroundColor Green
    }    
}

Write-Host "Finished." -ForegroundColor Green
return
