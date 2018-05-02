## How to install AzureCP

Download [latest release](https://github.com/Yvand/AzureCP/releases/latest).
Install and deploy the solution (that will automatically activate the "AzureCP" farm feature):

> **Important:**
> - Always start a new PowerShell console to ensure it uses up to date persisted objects and avoid concurrency update errors.
> - If some SharePoint servers do not have SharePoint service “Microsoft SharePoint Foundation Web Application” started, you need to deploy AzureCP.dll on their GAC manually. Read "Important - Limitations" below.

```powershell
Add-SPSolution "PATH TO WSP FILE"
Install-SPSolution -Identity "AzureCP.wsp" -GACDeployment
```

At this point AzureCP is inactive and it must be associated to a SPTrustedIdentityTokenIssuer:

```powershell
$trust = Get-SPTrustedIdentityTokenIssuer "SPTRUST NAME"
$trust.ClaimProviderName = "AzureCP"
$trust.Update()
```

Finally, AzureCP must be registered as an application in Azure Active Directory. Check [this page](Create-App-In-AAD.html) to create the app.

## Important - Limitations

What you need to know:
- Due to limitations of SharePoint API, do not associate AzureCP with more than 1 SPTrustedIdentityTokenIssuer.
- You must install AzureCP.dll manually in the GAC of SharePoint servers that do not have SharePoint service "Microsoft SharePoint Foundation Web Application" started.  
You can extract AzureCP.dll from AzureCP.wsp by opening it with [7-zip](https://www.7-zip.org/) and install it in the GAC with this PowerShell script:

```powershell
# Manually install AzureCP.dll in the GAC of a server
[System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
$publish = New-Object System.EnterpriseServices.Internal.Publish
$publish.GacInstall("C:\Data\Dev\AzureCP.dll")
```
