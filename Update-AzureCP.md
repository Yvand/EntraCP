## How to update AzureCP

> **Important:**  
> Start a **new PowerShell console** to ensure the use of up to date persisted objects, which avoids concurrency update errors.  
> Version 12 has breaking changes, please read below if you update from an earlier version.  
> If some SharePoint servers do not run SharePoint service "Microsoft SharePoint Foundation Web Application", azurecp.dll must be manually updated in their GAC as [documented here](Install-AzureCP.html).  

- Update solution

Run this cmdlet on the server running central administration:

```powershell
# This will start a timer job that will deploy the update on SharePoint servers. Central administration will restart during the process
Update-SPSolution -GACDeployment -Identity "AzureCP.wsp" -LiteralPath "F:\Data\Dev\AzureCP.wsp"
```

- Restart IIS service on each SharePoint server

## Updating from a version earlier than v12

Version 12 is a major update that has breaking changes. If you update from an earlier version:

- Claim type configuration list will be reset and all customization made to that list will be lost.
- Starting with v12, AzureCP creates group entities (permissions) using the Id of the groups rather than their DisplayName. There are 2 reasons for this change:
  - Group Id is unique, DisplayName is not
  - With group Id, AzureCP can get nested groups during augmentation, which is not possible with the DisplayName

As a consequence of this change, permissions granted to Azure AD groups before v12 will stop working, because the group value in the SAML token of AAD users (set with the Id) won't match the group value of group permission in the sites (set with the DisplayName).

To fix this, group permissions must be migrated to change their value with the Id. This can be done by calling method [SPFarm.MigrateGroup()](https://msdn.microsoft.com/en-us/library/office/microsoft.sharepoint.administration.spfarm.migrategroup.aspx) for each group to migrate:

```powershell
# SPFarm.MigrateGroup() will migrate group "c:0-.t|contoso.local|AzureGroupDisplayName" to "c:0-.t|contoso.local|a5e76528-a305-4345-8481-af345ea56032" in the whole farm
$oldlogin = "c:0-.t|contoso.local|AzureGroupDisplayName";
$newlogin = "c:0-.t|contoso.local|a5e76528-a305-4345-8481-af345ea56032";
[Microsoft.SharePoint.Administration.SPFarm]::Local.MigrateGroup($oldlogin, $newlogin);
```

> **Important:** This operation is farm wide and must be tested carefully before applying it in production environment.

Alternatively, administrators can configure AzureCP to use the property DisplayName for the groups, instead of the Id:

```powershell
Add-Type -AssemblyName "AzureCP, Version=1.0.0.0, Culture=neutral, PublicKeyToken=65dc6b5903b51636"
$config = [azurecp.AzureCPConfig]::GetConfiguration("AzureCPConfig")
# Get the ClaimTypeConfig used for groups and set property DirectoryObjectProperty to DisplayName
$ctConfig = $config.ClaimTypes| ?{$_.ClaimType -eq "http://schemas.microsoft.com/ws/2008/06/identity/claims/role"}
$ctConfig.DirectoryObjectProperty = [azurecp.AzureADObjectProperty]::DisplayName
$config.Update()
```
