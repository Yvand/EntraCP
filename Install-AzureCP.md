## How to install AzureCP

Download [latest release](https://github.com/Yvand/AzureCP/releases/latest).
Install and deploy the solution (that will automatically activate the "AzureCP" farm feature):

```powershell
# Run this on a new PowerShell console (it tends to avoid issues with local cache of persisted objects, that could cause errors on such operations)
Add-SPSolution "PATH TO WSP FILE"
Install-SPSolution -Identity "AzureCP.wsp" -GACDeployment
```

At this point AzureCP is inactive and it must be associated to a SPTrustedIdentityTokenIssuer:

```powershell
$trust = Get-SPTrustedIdentityTokenIssuer "SPTRUST NAME"
$trust.ClaimProviderName = "AzureCP"
$trust.Update()
```

Finally, AzureCP must be registered as an application in Azure Active Directory to be able to query it, with permission "Read directory data". Check [this page](Create-App-In-AAD.html) to create the app.

Once the app is created in Azure AD, information can be entered in Central Administration > Security > AzureCP Glogal configuration > "New Azure tenant" section.

## Important - Limitations

Due to limitations of SharePoint API, do not associate AzureCP with more than 1 SPTrustedIdentityTokenIssuer.

You must manually deploy AzureCP.dll on SharePoint servers that do not have SharePoint service "Microsoft SharePoint Foundation Web Application" started. You can use this PowerShell script:

```powershell
[System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
$publish = New-Object System.EnterpriseServices.Internal.Publish
$publish.GacInstall("C:\Data\Dev\AzureCP.dll")
```
