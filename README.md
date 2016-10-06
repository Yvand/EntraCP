This claims provider for SharePoint 2013 leverages [Azure AD Graph Client Library](http://www.nuget.org/packages/Microsoft.Azure.ActiveDirectory.GraphClient/) to query Azure Active Directory from the people picker. It also gets the groups of Azure users to augment their SAML token, so that permission can be granted on Azure groups.

![People picker with AzureCP](https://cloud.githubusercontent.com/assets/8788631/9786028/28938be0-57b7-11e5-8119-ea759f5c508e.png)

## Features
- Easy to configure with administration pages added in Central administration > Security.
- Connect to multiple Azure AD tenants in parallel (multi-threaded queries).
- Populate properties upon permission creation, e.g. email to allow email invitations to be sent.
- Supports rehydration for provider hosted apps. 
- Implements SharePoint logging infrastructure and logs messages in Area/Product "AzureCP". 
- Augment Azure AD users with their group membership.

## Customization capabilities
- Customize list of claim types, and their mapping with Azure AD users or groups. 
- Enable/disable augmentation.
- Enable/disable Azure AD lookup (to keep people picker returning results even if connectivity to Azure tenant is lost).
- Customize display of permissions. 
- Set a keyword to bypass Azure AD lookup. E.g. input "extuser:partner@contoso.com" directly creates permission "partner@contoso.com" on claim type set for this.
- Developers can easily customize it by inheriting AzureCP class and override many methods.

## Important - Limitations
Due to limitations of SharePoint API, do not associate AzureCP with more than 1 SPTrustedIdentityTokenIssuer.

You must manually deploy AzureCP.dll on SharePoint servers that do not have SharePoint service "Microsoft SharePoint Foundation Web Application" started. You can use this PowerShell script:
```powershell
[System.Reflection.Assembly]::Load("System.EnterpriseServices, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a")
$publish = New-Object System.EnterpriseServices.Internal.Publish
$publish.GacInstall("C:\Data\Dev\AzureCP.dll")
```

## How to install AzureCP
Download [latest release](https://github.com/Yvand/AzureCP/releases).
Install and deploy the solution (that will automatically activate the "AzureCP" farm feature):
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

Finally, AzureCP must be registered as an application in Azure Active Directory to be able to query it, with permission "Read directory data".  
[This article](https://github.com/AzureADSamples/ConsoleApp-GraphAPI-DotNet) explains how to register the app (start from Step 3).

Once the app is created in Azure AD, information can be entered in Central Administration > Security > AzureCP Glogal configuration > "New Azure tenant" section.

## Proxy configuration
AzureCP makes HTTP requests to Azure under the process of the site (for people picker requests) and the SharePoint STS (for augmentation), but there are also requests made by lsass.exe to validate CRL and certificate chain of certificates returned by Azure.<br>
If SharePoint servers need to connect through a HTTP proxy, additional configuration is required to configure it:
### For AzureCP to be able to connect to Azure
Add the following [proxy configuration](https://msdn.microsoft.com/en-us/library/kd3cf2ex.aspx) in the web.config of:
- SharePoint sites that use AzureCP
- SharePoine central administration site
- SharePoint STS located in 15\WebServices\SecurityToken
- SharePoint Web Services root site
- Also create file owstimer.exe.config in 15\BIN of each SharePoint server to put proxy configuration
```xml
<system.net>
    <defaultProxy>
        <proxy proxyaddress="http://proxy.contoso.com:3128" bypassonlocal="true" />
    </defaultProxy>
</system.net>
```
### For certificate validation to succeed
- Display proxy configuration with this command:<br>
netsh winhttp show proxy
- Set proxy configuration with this command:<br>
netsh winhttp set proxy proxy-server="http=myproxy;https=sproxy:88" bypass-list="*.foo.com"


## Claims supported
Azure AD binds property UserPrincipalName (which identifies the user) to claim type "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/name_". Unfortunately this claim type is reserved by SharePoint and cannot be used.

So the STS (whatever it is ACS, ADFS or another one) must transform this claim type into another one.  
AzureCP assumes it will be either "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/emailaddress_" or "_http://schemas.xmlsoap.org/ws/2005/05/identity/claims/upn_" (you should define only 1 in the SPTrust), so both are binded to property **UserPrincipalName**.  
Properties **DisplayName, GivenName, Surname** are also used to query users, and **Mail, Mobile, JobTitle** are used as metadata for the permission created.  
For groups, property **DisplayName** is linked to _http://schemas.microsoft.com/ws/2008/06/identity/claims/role_  

AzureCP can augment Azure AD users with their group membership. This is configurable and enabled by default. Groups are stored in claim type "_http://schemas.microsoft.com/ws/2008/06/identity/claims/role_", but this can be changed. However there can be only 1 claim type used for the roles, otherwise augmentation will be disabled.

**All of this can be configured** in AzureCP pages added in central administration > security.

## How to update AzureCP
Run Update-SPSolution cmdlet to start a timer job that that will deploy the update. You can monitor the progression in farm solutions page in central administration.
```powershell
Update-SPSolution -GACDeployment -Identity "AzureCP.wsp" -LiteralPath "PATH TO WSP FILE"
```
Then run iisreset on each server. This is mandatory to ensure that all w3wp processes, including SecurityTokenServiceApplicationPool, get restarted (this one is not restarted by solution update process).

## How to remove AzureCP
For an unknown reason, randomly SharePoint 2013 doesn’t uninstall correctly the solution because it removes assembly from the GAC before calling the feature receiver... When this happens, the claims provider is not removed and that causes issues when you re-install it.  
To uninstall safely, deactivate the farm feature before retracting the solution:
```powershell
Disable-SPFeature -identity "AzureCP"
Uninstall-SPSolution -Identity "AzureCP.wsp"
Remove-SPSolution -Identity "AzureCP.wsp"
```
